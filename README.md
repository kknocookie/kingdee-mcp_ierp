# kingdee-mcp

一个用于查询 Kingdee-ierp(金蝶云苍穹) 的 MCP 服务，基于 `uvx` 打包。

## 功能

- 查询销售订单
- 查询生产工单
- 查询 SO-MO 关联关系
- 查询交付风险
- 查询物料短缺
- 导出 SO-MO Excel 报表

## 环境变量

在启动服务之前，请先设置以下环境变量：

```powershell
$env:KINGDEE_BASE_URL = "https://your-host/ierp/kapi"
$env:KINGDEE_CLIENT_ID = "your-client-id"
$env:KINGDEE_CLIENT_SECRET = "your-client-secret"
$env:KINGDEE_USERNAME = "your-username"
$env:KINGDEE_ACCOUNT_ID = "your-account-id"
$env:KINGDEE_LANGUAGE = "zh_CN"
```

必填变量：

- `KINGDEE_BASE_URL`
- `KINGDEE_CLIENT_ID`
- `KINGDEE_CLIENT_SECRET`
- `KINGDEE_USERNAME`
- `KINGDEE_ACCOUNT_ID`


## 第一步
1.第三方应用——新增
2.获取CLIENT_ID，CLIENT_SECRET，ACCOUNT_ID

## 第二步

```bash
pip install kingdee-mcp-ierp
```

## 第三步

设置配置

## Trae MCP 配置

你也可以直接复制 `.mcp.json.example`，然后填入你自己的配置值。

```json
{
  "mcpServers": {
    "kingdee": {
      "command": "uvx",
      "args": ["kingdee-mcp-ierp"],
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
