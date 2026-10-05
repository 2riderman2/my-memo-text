{
  "$schema": "https://developer.microsoft.com/json-schemas/sp/v2/row-formatting.schema.json",
  "hideSelection": true,
  "hideColumnHeader": true,
  "groupProps": {
    "hideFooter": true,
    "headerFormatter": {
      "elmType": "div",
      "style": {
        "display": "flex",
        "flex-direction": "column",
        "align-items": "center",
        "width": "100%",
        "padding": "12px",
        "box-sizing": "border-box",
        "background-color": "white",
        "border-radius": "8px",
        "margin-bottom": "6px"
      },
      "children": [
        {
          "elmType": "img",
          "attributes": {
            "src": "https://○○.sharepoint.com/sites/○○/SiteAssets/PC01.jpg",
            "alt": "パソコン①"
          },
          "style": {
            "width": "100%",
            "max-width": "180px",
            "height": "110px",
            "object-fit": "contain",
            "margin-bottom": "8px"
          }
        },
        {
          "elmType": "div",
          "txtContent": "パソコン①",
          "style": {
            "font-size": "18px",
            "font-weight": "600",
            "margin-bottom": "10px"
          }
        },
        {
          "elmType": "div",
          "txtContent": "今後の予約",
          "style": {
            "width": "100%",
            "font-size": "13px",
            "font-weight": "600",
            "padding-top": "8px",
            "border-top": "1px solid #e1dfdd"
          }
        }
      ]
    }
  },
  "rowFormatter": {
    "elmType": "div",
    "style": {
      "display": "=if([$KeepDisplay] == true, 'none', 'flex')",
      "flex-direction": "column",
      "width": "100%",
      "box-sizing": "border-box",
      "padding": "8px 10px",
      "margin-bottom": "5px",
      "border-radius": "6px",
      "background-color": "#f8f8f8",
      "border": "1px solid #edebe9"
    },
    "children": [
      {
        "elmType": "div",
        "style": {
          "display": "flex",
          "align-items": "center",
          "margin-bottom": "5px"
        },
        "children": [
          {
            "elmType": "span",
            "txtContent": "=if(@now < [$StartDate], '予約', if(@now <= [$EndDate], '利用中', '期限超過'))",
            "style": {
              "font-size": "12px",
              "font-weight": "600",
              "padding": "2px 8px",
              "border-radius": "12px",
              "margin-right": "8px",
              "color": "=if(@now < [$StartDate], '#8a6d00', if(@now <= [$EndDate], '#a4262c', '#a4262c'))",
              "background-color": "=if(@now < [$StartDate], '#fff4ce', if(@now <= [$EndDate], '#fde7e9', '#fde7e9'))"
            }
          },
          {
            "elmType": "span",
            "txtContent": "[$ReservedBy.title]",
            "style": {
              "font-size": "14px",
              "font-weight": "600"
            }
          }
        ]
      },
      {
        "elmType": "div",
        "txtContent": "=toLocaleDateString([$StartDate]) + ' ' + toLocaleTimeString([$StartDate]) + ' ～ ' + toLocaleTimeString([$EndDate])",
        "style": {
          "font-size": "13px",
          "color": "#323130"
        }
      },
      {
        "elmType": "div",
        "txtContent": "[$Purpose]",
        "style": {
          "font-size": "12px",
          "color": "#605e5c",
          "margin-top": "3px"
        }
      }
    ]
  }
}
