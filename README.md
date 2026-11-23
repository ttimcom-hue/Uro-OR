# 泌尿科 刀房調度板＋科務入口

## 網站結構
- `index.html`：刀房調度板（原本的網址不變）
- `hub/index.html`：科務入口（首頁、人員名冊、常用連結），網址＝原網址後面加 `hub/`
- `or/index.html`：只是轉址回調度板（之前若用過 /or/ 網址，不會失效）
- 之後的新工具各放一個資料夾，例如 `cancer/`、`meeting/`、`roster/`、`teaching/`

## 共用資料（同一個 Google Apps Script 同步網址）
| 資料鍵 | 內容 | 誰寫入 |
|---|---|---|
| `common/people` | 人員名冊：簡稱、全名、身分（VS／R／F）、級別、在職 | 入口、調度板 |
| `common/links` | 常用連結 | 入口 |
| `config/main` | 調度板設定，含值班 `duty`、住院醫師休假 `leaves`、主治休假 `vsLeaves` | 調度板 |
| `days/YYYY-MM-DD` | 當日刀表與人力 | 調度板 |
| `er/list` | 急診刀清單 | 調度板 |

雲端只放人員姓名、術式、日期，不放病人姓名與病歷號。
