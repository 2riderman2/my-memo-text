{
  "$schema": "https://developer.microsoft.com/json-schemas/sp/v2/column-formatting.schema.json",
  "elmType": "button",
  "txtContent": "✓ 完了",
  "style": {
    "display": "=if([$Status] == '完了', 'none', 'inline-flex')",
    "cursor": "pointer"
  },
  "customRowAction": {
    "action": "setValue",
    "actionInput": {
      "Status": "完了",
      "CompletedDate": "@now"
    }
  }
}
