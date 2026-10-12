# 07 — Adobe Form `ZSD_EPACK_AF` (SFP → Form)

**Description:** Export consolidated packing list
**Interface:** `ZSD_EPACK_AF_IF`
**Layout reference:** `PackingList_90000014_Tarmac_UAE_BySea.pdf` (without notify party) and
`PackingList_90000081*.pdf` (with notify party) in the repository root. **One form for both.**

### Context
Drag `IS_HEADER` and `IT_ITEM` from the interface into the context (all fields).

### Layout (Adobe LiveCycle Designer)
Page A4 portrait, header repeated on every page, item table flowing over pages.


| Area | Content | Shown |
|---|---|---|
| Master page | Title "PACKING LIST", exporter box, consignee / buyer, notify parties, right box (PL no., date, exporter's ref., CIN, GST, order no., buyer's ref., countries, terms), shipping box, marks / containers / packages / goods, item column headings, page "Page n of m" | every page |
| Consignee | `SAME_PARTY` = X: one "CONSIGNEE :" block (`CNS_*`). Otherwise "SOLD TO / BILL TO PARTY (BUYER) :" (`BUY_*`) and "SHIP TO PARTY (CONSIGNEE) :" (`CNS_*`) | presence script on `SAME_PARTY` |
| Notify parties | Subform with "1ST … 5TH NOTIFY PARTY :" + `NP1`…`NP5`; each party subform hidden when its field is empty | presence script on `FORMAT` = `N` |
| Item table | Bound to `IT_ITEM`. Three body rows: heading (`ROW_TYPE` = `H`, `ROW_TEXT` across the table), item (`I`), subtotal (`S`, `ARTICLES` + `ROW_TEXT`) — each hidden unless its row type matches. Header row repeated on each page; rows may break between pages | every page |
| Totals row | `TOT_ART`, `TOT_TEXT`, `TOT_QTY`, `TOT_GWT`, `TOT_NWT` | table footer, last page |
| Declaration / signature | "Declaration :" + `DECLARATION`; "Signature & Date", "For " + `SIGN_FOR`, "Authorized Signatory" | last page |

### Static elements


| Element on the sample | Handling |
|---|---|
| Title "PACKING LIST", all labels, column headings | Static text in the form |
| Page "Page 1 of 4" | Form system fields (current page / total pages) |
| Company stamp and signature | Not printed (signed by hand), or an MIME graphic if the business wants it |
| Diagonal line through the empty part of the table on the last page | Form layout (as on the sample; optional) |
| Header repeated on every page | Header subform placed on the master page |


### Scripting needed (JavaScript, `presence`)
* Consignee: `if (IS_HEADER.SAME_PARTY.rawValue == "X")` show one CONSIGNEE block, else BUYER + CONSIGNEE.
* Notify-party subform: visible only when `IS_HEADER.FORMAT.rawValue == "N"`; each NPn subform hidden when empty.
* Item body rows: row `H` / `I` / `S` visible only when `ROW_TYPE` matches.
* "Merged" cells: bottom border of Articles / weights only when `ART_LAST` / `WT_LAST` = `X`.

The report calls the form with `FP_FUNCTION_MODULE_NAME`, `FP_JOB_OPEN`, the generated
function module and `FP_JOB_CLOSE` — one job per packing list (FORM `SEND_FORM`).
