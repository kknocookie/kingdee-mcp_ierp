# kingdee-mcp-ierp

An MCP server for Kingdee-ierp queries.

## Features

- Query sale orders
- Query manufacturing orders
- Query SO-MO relations
- Query delivery risk
- Query material shortage
- Export SO-MO Excel report

## Local Run

```bash
pip install kingdee-mcp-ierp
kingdee-mcp-ierp
```

## Environment Variables

Set these environment variables before starting the server:

```powershell
$env:KINGDEE_BASE_URL = "https://your-host/ierp/kapi"
$env:KINGDEE_CLIENT_ID = "your-client-id"
$env:KINGDEE_CLIENT_SECRET = "your-client-secret"
$env:KINGDEE_USERNAME = "your-username"
$env:KINGDEE_ACCOUNT_ID = "your-account-id"
$env:KINGDEE_LANGUAGE = "zh_CN"
```

Required variables:

- `KINGDEE_BASE_URL`
- `KINGDEE_CLIENT_ID`
- `KINGDEE_CLIENT_SECRET`
- `KINGDEE_USERNAME`
- `KINGDEE_ACCOUNT_ID`

Optional variables:

- `KINGDEE_LANGUAGE` (defaults to `zh_CN`)

You can copy `.env.example` as a local reference, but do not commit your real credentials.

```powershell
Copy-Item .env.example .env
```

## Trae MCP Config

Install the package first with `pip install kingdee-mcp-ierp`, then configure your MCP client like this.

```json
{
  "mcpServers": {
    "kingdee": {
      "command": "kingdee-mcp-ierp",
      "env": {
        "KINGDEE_BASE_URL": "https://your-host/ierp/kapi",
        "KINGDEE_CLIENT_ID": "your-client-id",
        "KINGDEE_CLIENT_SECRET": "your-client-secret",
        "KINGDEE_USERNAME": "your-username",
        "KINGDEE_ACCOUNT_ID": "your-account-id",
        "KINGDEE_LANGUAGE": "zh_CN"
      }
    }
  }
}
```
