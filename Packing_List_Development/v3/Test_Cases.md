# Test cases — WRICEF 102-B (version 3)

Preparation: a sales order **SO1** (quotation reference, 2+ items, delivered and goods issued), a sales order
**SO2** (other customer), a sales order without delivery **SO3**, a user without `V_VBAK_VKO` for SO2's sales area.
Reference data: sample `test_data/Sample_data_ABAP.xlsx`, PDFs `PackingList_90000014 …` (without NP) and `PackingList_90000081 …` (with NP).

## Report 1 — ZSD_EPACK_EXTRACT (T-code ZSD_EPACK)

| # | Test | Steps | Expected result |
|---|---|---|---|
| R1-01 | Single order | `S_VBELN` = SO1 | ALV: one row per VBAP item, 46 columns in Z-table order (packing list no … declaration, 2nd packing list no not shown) with the sheet headings |
| R1-02 | Range / multiple | SO1 – SO2 and multiple selection | Rows of all orders, sorted by order and item |
| R1-03 | Order does not exist | single value 9999999999 | Popup: "Sales order does not exist"; other orders displayed |
| R1-04 | No authorization | user without V_VBAK_VKO for SO2 | Popup: "No authorization for sales area …"; SO2 rows missing |
| R1-05 | Exporter | compare with VKORG → TVKO → T001 → ADRC | `NAME, STREET, STREET2, CITY, POSTCODE, REGION TEXT, COUNTRY TEXT`; empty part kept between commas |
| R1-06 | IEC / CIN / GST | compare with ADRC-BUILDING, T001Z (ZCIN), J_1BBRANCH of the plant business place | Values filled; missing → warning in popup |
| R1-07 | Order no. & date | SO1 with quotation | VBKD-BSTKD / BSTDK of the quotation; order without VGBEL → warning, empty |
| R1-08 | Sold-to / ship-to | compare with KNA1 / KNVP (WE) / ADRC / ADR6 / ADR2 | 10 fixed positions incl. e-mail, telephone (R3_USER 1), mobile (R3_USER 3); country of final destination = ship-to country text |
| R1-09 | Payment / Incoterms | | TVZBT-VTEXT; `INCO1 INCO2` |
| R1-10 | Item data | | Part No. (external), MAKT description, KWMENG, MARC-STEUC |
| R1-11 | Editability | try to change Part No., Quantity, Exporter | Not possible (protected); manual columns (packing list no, Sl.No., weights, NP …) editable |
| R1-12 | Insert / delete row | ALV toolbar | Not offered (A16) |
| R1-13 | Date / number entry | type `31.02.2026` in Packing list date, `abc` in Sl.No. | ALV error, cell marked |
| R1-14 | GW < NW in ALV | Gross 10, Net 12 | Error "Invalid Weight: Gross Weight cannot be less than Net Weight." on the cell |
| R1-15 | Download with errors | R1-14 open, press Download Excel | "Correct the marked entries before the download"; no file |
| R1-16 | Download | fill some manual cells without leaving the last cell, press Download Excel | File saved; the last entry is in the file (CHECK_CHANGED_DATA); row 1 = 46 headings exactly as sheet "Z table"; column P = column A; zero values empty; dates DD.MM.YYYY |
| R1-17 | No database update | SE16 ZSD_EPACK_DATA before/after | Unchanged |
| R1-18 | Column width | open the file | Headings complete, widths = heading or content (max. 60) |

## Report 2 — ZSD_EPACK_UPLOAD

| # | Test | Steps | Expected result |
|---|---|---|---|
| R2-01 | Wrong extension | `x.csv` | Error on the selection screen (V02) |
| R2-02 | File missing | non-existing path | Error on the selection screen (V03) |
| R2-03 | Not an Excel file | rename a .txt to .xlsx | "The file is not a valid Excel (.xlsx) file" (V04) |
| R2-04 | Changed heading | rename heading of column K, or move a column | Error per wrong column with letter, found and expected heading; nothing else processed (V06) |
| R2-05 | Happy path, test run | file of R1-16 completely filled, Test run ✔ | Result list "TEST RUN – nothing saved …", success line "No errors – run again …" |
| R2-06 | Happy path, save + print (without NP) | same file, Test run off, format without NP | Packing list(s) saved (SE16: one row per Excel row, MEINS from order, GEWEI KG, ERNAM/ERDAT/ERZET); print dialog; one form per packing list; result "saved", "printed" |
| R2-07 | With NP | NP1, NP2 filled, format with NP | `ZSD_EPACK_SF_NP` printed; NP3–NP5 crossed out |
| R2-08 | NP filled, format without NP | | Error V19, nothing saved (all-or-nothing) |
| R2-09 | Rerun same file | repeat R2-06 | Error "Packing list is already saved – tick Overwrite"; table unchanged (no duplicates) |
| R2-10 | Overwrite | R2-09 with Overwrite ✔, one row deleted in Excel | Warning "will be replaced"; old rows deleted, new rows inserted (row count = file) |
| R2-11 | Mandatory keys | empty packing list no / sales order / Sl.No. in one row each | 3 errors with row and column (V11) |
| R2-12 | Duplicate key | two rows same packing list + order + Sl.No. | Error "… also in Excel row n" (V12) |
| R2-13 | Column P differs | P ≠ A | Error V09 |
| R2-14 | Wrong types | date `32.13.2026`, Sl.No. `1.5`, gross weight `-5`, Article No. 16 characters, declaration 1 600 characters | One error each with row and column (V08) |
| R2-15 | Order does not exist | 9999999999 | Error V13 |
| R2-16 | Material not in order | Part No. of another material | Error V15 |
| R2-17 | **Quantity > delivered** | material delivered 100 (two deliveries 60 + 40); file: two rows 70 + 40 | Error "Uploaded Quantity should not be greater than Delivered Quantity (… uploaded 110, delivered 100)" on both rows (V16) |
| R2-18 | Quantity = delivered | 60 + 40 | No error |
| R2-19 | Order not delivered | SO3, quantity 10 | Error V16 (delivered 0) |
| R2-20 | **GW < NW** | gross 2100, net 2107.2 | Error "Invalid Weight: Gross Weight cannot be less than Net Weight." (V17) |
| R2-21 | All-or-nothing | file with 2 packing lists, error in one, save mode default | Nothing saved, nothing printed (V22) |
| R2-22 | Per packing list | same file, "Save correct packing lists only" | Correct packing list saved and printed, the other not |
| R2-23 | Lock | user A runs SE16N-edit / a second session holding the lock (debug stop after ENQUEUE) | "Packing list is locked by another user – not saved" |
| R2-24 | Header differs | Pre-Carriage different in row 5 | Warning V21; form prints the value of the first row |
| R2-25 | Header only on first row | header columns filled only in the first row of each packing list | No warning; form shows header values |
| R2-26 | Totals / order | rows not sorted by Sl.No. | Form items sorted by Sl.No.; totals = sums (sample: qty 12146, gross 116323.8, net 115386.78, articles 4074) |
| R2-27 | Smart Form missing | before the form is activated | "Smart Form … not found – data saved, nothing printed" |
| R2-28 | Print dialog cancelled | Cancel in the dialog | Warning; data stays saved |
| R2-29 | Excel date number / text dates | dates typed as dates and as text `10-09-2026` | Both accepted |
| R2-30 | Several sales orders in one cell | `12345, 12346` | Stored `0000012345,0000012346`; quantity check against both orders |
| R2-31 | Background | run in background | "run the report in dialog" |
| R2-32 | Performance | 1 000 rows, 20 orders | Runtime within seconds; SQL trace (ST05) shows one statement per table, no statement per row |

## Code checks before transport

* ATC / SLIN / Code Inspector clean (the unused FORMs of the shared include in each program are expected).
* ST05 trace of R1-01 and R2-06: no SELECT per row.
* Unit test results to be recorded in FS section 3.1.
