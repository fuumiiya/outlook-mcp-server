# Outlook MCP Server

Azure FunctionsにホスティングするMCPサーバーです。Microsoft Graph APIを使用してOutlookカレンダーにアクセスし、予定の取得と会議予約が可能です。

## 機能

- 📅 **予定取得**: 指定したユーザーの直近30日間の予定を取得
- 📝 **会議予約**: 新しい会議予定を作成（参加者、場所、Teams会議対応）
- 🌏 **日本時間対応**: JST（日本標準時）での時刻処理
- 🔐 **Azure認証**: Azure AD認証による安全なアクセス

## 利用可能なツール

### `get_user_outlook_events`
指定したユーザーのOutlook予定を取得します。

**パラメータ:**
- `userPrincipalName` (必須): 予定を取得したいユーザーのUPN（例: user@example.com）

**レスポンス例:**
```json
{
  "user": {
    "userPrincipalName": "user@example.com"
  },
  "value": [
    {
      "subject": "チーム会議",
      "start": "2025-01-15T09:00:00",
      "end": "2025-01-15T10:00:00",
      "isAllDay": false
    }
  ]
}
```

### `create_simple_event`
指定ユーザーのOutlookに簡易予定を作成します。

**パラメータ:**
- `userPrincipalName` (必須): 予定を作成するユーザーのUPN
- `subject` (必須): 予定の件名
- `start` (必須): 開始時刻（ISO形式）
- `end` (必須): 終了時刻（ISO形式）
- `attendees` (任意): 参加者のメールアドレス（カンマ区切り）
- `isOnlineMeeting` (任意): Teams会議の有無（boolean）
- `location` (任意): 会議場所
- `body` (任意): 予定の本文

**レスポンス例:**
```json
{
  "created": {
    "id": "event-id",
    "subject": "新しい会議",
    "start": "2025-01-15T09:00:00",
    "end": "2025-01-15T10:00:00",
    "webLink": "https://outlook.office.com/...",
    "attendees": ["user1@example.com", "user2@example.com"],
    "location": "会議室A",
    "isOnlineMeeting": true
  }
}
```

## 使用例

作成者は、このMCPサーバーをresponsesAPI経由で使用しています。

### MCPクライアントでの使用

#### 予定取得の例
```json
{
  "tool": "get_user_outlook_events",
  "arguments": {
    "userPrincipalName": "user@company.com"
  }
}
```

#### 会議予約の例
```json
{
  "tool": "create_simple_event",
  "arguments": {
    "userPrincipalName": "user@company.com",
    "subject": "プロジェクト会議",
    "start": "2025-01-20T14:00:00",
    "end": "2025-01-20T15:00:00",
    "attendees": "colleague1@company.com,colleague2@company.com",
    "isOnlineMeeting": true,
    "location": "会議室A",
    "body": "プロジェクトの進捗について話し合います。"
  }
}
```

### レスポンス例

#### 予定取得の成功レスポンス
```json
{
  "user": {
    "userPrincipalName": "user@company.com"
  },
  "value": [
    {
      "subject": "チーム会議",
      "start": "2025-01-15T09:00:00",
      "end": "2025-01-15T10:00:00",
      "isAllDay": false
    },
    {
      "subject": "プロジェクトレビュー",
      "start": "2025-01-16T14:00:00",
      "end": "2025-01-16T15:30:00",
      "isAllDay": false
    }
  ]
}
```

#### エラーレスポンス
```json
{
  "error": "userPrincipalName is required."
}
```

## セットアップ

### 1. Azure Functions プロジェクトの作成

```bash
func init outlook-mcp-server --python
cd outlook-mcp-server
```

### 2. 依存関係のインストール

```bash
pip install -r requirements.txt
```

### 3. 環境設定

`local.settings.json` に以下の設定を追加：

```json
{
  "IsEncrypted": false,
  "Values": {
    "FUNCTIONS_WORKER_RUNTIME": "python",
    "AzureWebJobsStorage": "",
    "AZURE_TENANT_ID": "your-tenant-id",
    "AZURE_CLIENT_ID": "your-client-id",
    "AZURE_CLIENT_SECRET": "your-client-secret"
  }
}
```

**注意**: 実際の値はAzure ADアプリケーション登録で取得した値に置き換えてください。

### 4. Azure AD アプリケーションの設定

1. Azure Portal でアプリケーションを登録
2. Microsoft Graph API の以下の権限を付与：
   - `Calendars.Read`
   - `Calendars.ReadWrite`
3. クライアントシークレットを生成
4. Azure Functions の設定に認証情報を追加

### 5. デプロイ

```bash
func azure functionapp publish <your-function-app-name>
```

## 技術スタック

- **Azure Functions**: サーバーレス実行環境
- **Microsoft Graph API**: Outlook カレンダーアクセス
- **Python 3.x**: 実装言語
- **Azure AD**: 認証・認可

## ライセンス

MIT License

## 貢献

プルリクエストやイシューの報告を歓迎します！

## 注意事項

- このサーバーはMCP（Model Context Protocol）に対応しています
- Azure AD認証が必要です
- Microsoft Graph APIの利用制限にご注意ください
- 本番環境では適切なセキュリティ設定を行ってください
- **現時点で終日予定取得できないので、今後カスタマイズが必要**
- **リファクタリングの余地が多くあります（笑）**
