# Standard library imports
import json
import logging
from datetime import datetime, timedelta, timezone, time
from urllib.parse import urlencode

# Azure imports
import azure.functions as func
from azure.identity.aio import DefaultAzureCredential

# Microsoft Graph imports
from msgraph import GraphServiceClient
from kiota_abstractions.base_request_configuration import RequestConfiguration
from kiota_abstractions.headers_collection import HeadersCollection
from msgraph.generated.models.event import Event
from msgraph.generated.models.date_time_time_zone import DateTimeTimeZone
from msgraph.generated.models.attendee import Attendee
from msgraph.generated.models.email_address import EmailAddress
from msgraph.generated.models.attendee_type import AttendeeType
from msgraph.generated.models.location import Location
from msgraph.generated.models.item_body import ItemBody
from msgraph.generated.models.body_type import BodyType

app = func.FunctionApp(http_auth_level=func.AuthLevel.FUNCTION)

# 固定オフセット（日本はDSTなし）
JST = timezone(timedelta(hours=9))

# 整形に使用するへヘルパー関数群
def _dtz_to_str(dtz):
    """Graph の DateTimeTimeZone から ISO 文字列のみを返す（例: '2025-09-17T09:00:00'）"""
    if not dtz:
        return None
    # 属性名の揺れ（date_time / dateTime）を吸収
    dt = getattr(dtz, "date_time", getattr(dtz, "dateTime", None))
    # datetime 型なら ISO に
    if hasattr(dt, "isoformat"):
        return dt.isoformat()
    # 既に str の場合想定
    return dt

def _event_to_min_dict(ev):
    """LLM 用に最小項目のみ抽出"""
    return {
        "subject": getattr(ev, "subject", None),
        "start": _dtz_to_str(getattr(ev, "start", None)),
        "end": _dtz_to_str(getattr(ev, "end", None)),
        "isAllDay": getattr(ev, "is_all_day", getattr(ev, "isAllDay", None)),
    }

def _parse_to_jst(dt_str: str):
    """ISO文字列を datetime へ。オフセット無ければ JST を付与。"""
    if not dt_str:
        return None
    try:
        dt = datetime.fromisoformat(dt_str)
    except Exception:
        try:
            dt = datetime.strptime(dt_str, "%Y-%m-%dT%H:%M:%S")
        except Exception:
            return None
    # タイムゾーン情報が無ければ JST を付与（Prefer:outlook.timezone でローカル返却想定）
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=JST)
    return dt

def _to_bool(v, default=False):
    if isinstance(v, bool): return v
    if isinstance(v, str):  return v.strip().lower() in ("true","1","yes","y","on")
    if isinstance(v, (int, float)): return bool(v)
    return default

@app.generic_trigger(
    arg_name="context",
    type="mcpToolTrigger",
    toolName="get_user_outlook_events",
    description="指定したユーザーのOutlook予定を取得します。",
    toolProperties=json.dumps([
        {
            "propertyName": "userPrincipalName",
            "propertyType": "string",
            "description": "予定を取得したいユーザーのUPN（例: user@example.com）",
            "required": True
        }
    ]),
)
async def get_user_outlook_events(context) -> str:
    try:
        content = json.loads(context)
        upn = content["arguments"].get("userPrincipalName")
        if not upn:
            return json.dumps({"error": "userPrincipalName is required."}, ensure_ascii=False)

        # JST の「明日 00:00」
        now_jst = datetime.now(JST)
        tomorrow_start_jst = datetime.combine(
            now_jst.date() + timedelta(days=1), time(0, 0, 0), tzinfo=JST
        )

        async with DefaultAzureCredential() as credential:
            client = GraphServiceClient(
                # あなたの環境は複数形の 'credentials' が正（単数 'credential' はNG）
                credentials=credential,
                scopes=["https://graph.microsoft.com/.default"]
            )

            # 期間：明日 00:00（JST）〜 30日後（必要に応じて調整）
            start = tomorrow_start_jst.isoformat()                         # 例: 2025-10-07T00:00:00+09:00
            end   = (tomorrow_start_jst + timedelta(days=30)).isoformat()  # 例: 2025-11-06T00:00:00+09:00

            # クエリは urlencode で正しくエンコード（'+' を '%2B' に）
            params = {
                "startDateTime": start,
                "endDateTime": end,
                "$select": "subject,start,end,isAllDay",
                "$orderby": "start/dateTime asc",
                "$top": "100",  # calendarView の $top は 1〜1000
            }
            cv_url = f"https://graph.microsoft.com/v1.0/users/{upn}/calendarView?{urlencode(params)}"

            # RequestBuilder を生URLで差し替え
            cv_builder = client.users.by_user_id(upn).calendar_view.with_url(cv_url)

            # Prefer で返却の start/end を JST 表示に
            headers = HeadersCollection()
            headers.add("Prefer", 'outlook.timezone="Tokyo Standard Time"')
            req = RequestConfiguration(headers=headers)

            # 期間内の「各回（occurrences / exceptions）」が展開されて返る
            resp = await cv_builder.get(request_configuration=req)

            # 最小項目へ整形
            items = []
            if resp and getattr(resp, "value", None):
                for ev in resp.value:
                    items.append(_event_to_min_dict(ev))

            # （保険）アプリ側でも明日以降に絞る
            filtered = []
            for it in items:
                sd = _parse_to_jst(it.get("start"))
                if sd and sd >= tomorrow_start_jst:
                    filtered.append(it)

            # （保険）開始時刻で昇順ソート
            filtered.sort(key=lambda it: _parse_to_jst(it.get("start")) or datetime.max.replace(tzinfo=JST))

            result = {
                "user": {"userPrincipalName": upn},
                "value": filtered
            }
            return json.dumps(result, ensure_ascii=False, indent=2)

    except Exception as e:
        logging.exception("Error retrieving events")
        return json.dumps({"error": str(e)}, ensure_ascii=False)

@app.generic_trigger(
    arg_name="context",
    type="mcpToolTrigger",
    toolName="create_simple_event",
    description="指定ユーザーのOutlookに簡易予定を作成（全員必須参加者・Teams対応・場所/本文対応）。",
    # ★ 重要：トップレベルは「配列」＋ json.dumps！
    toolProperties=json.dumps([
        {"propertyName": "userPrincipalName", "propertyType": "string",  "required": True},
        {"propertyName": "subject",           "propertyType": "string",  "required": True},
        {"propertyName": "start",             "propertyType": "string",  "required": True},
        {"propertyName": "end",               "propertyType": "string",  "required": True},
        # ★ 暫定：array ではなく string（CSV/セミコロン対応）
        {"propertyName": "attendees",         "propertyType": "string",  "required": False},
        {"propertyName": "isOnlineMeeting",   "propertyType": "boolean", "required": False},
        {"propertyName": "location",          "propertyType": "string",  "required": False},
        {"propertyName": "body",              "propertyType": "string",  "required": False}
    ]),
)
async def create_simple_event(context) -> str:
    try:
        args = json.loads(context).get("arguments", {})
        upn = args.get("userPrincipalName")
        subject = args.get("subject")
        start_str = args.get("start")
        end_str = args.get("end")

        raw_attendees = args.get("attendees", None)
        is_online = _to_bool(args.get("isOnlineMeeting", False))
        location_str = args.get("location")
        body_str = args.get("body")

        if not all([upn, subject, start_str, end_str]):
            return json.dumps({"error": "必須項目が不足しています"}, ensure_ascii=False)

        # 時刻は「ローカル時刻（オフセット無し）＋Tokyo Standard Time」
        def _to_local(dt: str) -> str:
            return dt.split("+")[0].split("Z")[0]

        start_dttz = DateTimeTimeZone(date_time=_to_local(start_str), time_zone="Tokyo Standard Time")
        end_dttz   = DateTimeTimeZone(date_time=_to_local(end_str),   time_zone="Tokyo Standard Time")

        # 参加者（全員必須）：配列 or 文字列（カンマ/セミコロン区切り）
        if isinstance(raw_attendees, list):
            emails = [e.strip() for e in raw_attendees if isinstance(e, str)]
        elif isinstance(raw_attendees, str):
            emails = [t.strip() for t in raw_attendees.replace(";", ",").split(",") if t.strip()]
        else:
            emails = []
        emails = list(dict.fromkeys([e for e in emails if "@" in e]))  # 重複除去＆簡易バリデーション

        attendees = [
            Attendee(email_address=EmailAddress(address=e), type="required")  # ← 文字列で安全に
            for e in emails
        ] or None

        # 場所・本文（BodyType は環境差対策付き）
        ev_location = Location(display_name=location_str) if location_str else None
        if body_str:
            try:
                ct = BodyType.Html if ("<" in body_str) else BodyType.Text  # PascalCase 版
            except AttributeError:
                ct = "html" if ("<" in body_str) else "text"               # 文字列フォールバック
            ev_body = ItemBody(content=body_str, content_type=ct)
        else:
            ev_body = None

        # ★ Provider 指定はしない（SDK型依存を排除）
        ev = Event(
            subject=subject,
            start=start_dttz,
            end=end_dttz,
            attendees=attendees,
            location=ev_location,
            body=ev_body,
            is_online_meeting=is_online
            # online_meeting_provider は渡さない
        )

        # Prefer: representation + JST 表示
        headers = HeadersCollection()
        headers.add("Prefer", 'return=representation')
        headers.add("Prefer", 'outlook.timezone="Tokyo Standard Time"')
        req = RequestConfiguration(headers=headers)

        async with DefaultAzureCredential() as credential:
            client = GraphServiceClient(credentials=credential, scopes=["https://graph.microsoft.com/.default"])
            created = await client.users.by_user_id(upn).events.post(ev, request_configuration=req)

        return json.dumps({
            "created": {
                "id": getattr(created, "id", None),
                "subject": getattr(created, "subject", None),
                "start": getattr(created, "start", None).date_time if getattr(created, "start", None) else None,
                "end": getattr(created, "end", None).date_time if getattr(created, "end", None) else None,
                "webLink": getattr(created, "web_link", None),
                "attendees": [
                    getattr(getattr(at, "email_address", None), "address", None)
                    for at in (getattr(created, "attendees", []) or [])
                ],
                "location": getattr(getattr(created, "location", None), "display_name", None),
                "isOnlineMeeting": getattr(created, "is_online_meeting", None)
            }
        }, ensure_ascii=False, indent=2)

    except Exception as e:
        logging.exception("Error creating simple event")