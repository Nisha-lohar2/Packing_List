# ZSD_EPACK_REPORT — process, relationships, assumptions (version 2)

Report: `v2/ZSD_EPACK_REPORT.prog.abap`. Sources: FS `Object ID- 102(B)Export consolidated packing
list.docx`, `Logic_sheet1.xlsx` (rows Bxx), sample `test_data/Sample_data_ABAP.xlsx`.

**Status:** checked offline (balanced blocks, every PERFORM has a FORM, no SELECT or SORT inside a
loop, every Adobe Form field used exists in `src/Adobe_Form_Interface_Structures.md`).
**Not compiled or run in S2A yet.**

---

## 1. Process

| Step | User does | Program does |
|---|---|---|
| 1 | Enters **sales order(s)**, leaves the file empty, runs | Reads the SAP data whose logic is given, shows an **ALV with all packing-list columns** (filled or empty). One row per billing item of the sales orders. |
| 2 | ALV → *Export → Spreadsheet* | Standard ALV function |
| 3 | In Excel: fills the empty columns (packing list no., date, notifiers, shipment data, Sl.No., articles, quantity, weights …), deletes / copies rows as needed | — |
| 4 | Enters the **filled file**, runs | Reads the file back (columns found by heading), groups the rows by **Packing List No.** and prints the Adobe Form `ZSD_EPACK_AF` once per packing list (print preview on `LP01`) |

Selection screen: `S_VBELN` (sales order, select-options) and `P_FILE` (Excel file, F4). Both optional;
one of them must be filled. File filled → step 4, otherwise step 1.

## 2. Relationships

| Link | Implemented as | Logic |
|---|---|---|
| Sales order → billing documents | `VBRP-AUBEL IN S_VBELN`, joined to `VBRK` | B13 uses AUBEL |
| Exporter | `VBRK-BUKRS` → `T001-ADRNR` → `ADRC` | B3 |
| GST No. | `J_1BBRANCH-GSTIN` for `BUKRS` + `BRANCH` = `VBRK-BUPLA` | B6 |
| CIN | `T001Z-PAVAL`, `PARTY` = `ZCIN` | B7 |
| IEC | `ADRC-BUILDING` of the company address | B11 |
| Order no. & date | `VBRP-AUBEL` → `VBAK-VGBEL` → `VBKD` (`POSNR` 000000) `BSTKD`, `BSTDK` | B13 |
| Sold-to / ship-to | `VBRK-KUNAG` + sales area → `KNVP` (SP / SH) → `KUNN2` → `KNA1-ADRNR` → `ADRC`, `ADR6`, `ADR2` | B20 – B29 |
| Country of final destination | ship-to `ADRC-COUNTRY` → `T005T-LANDX` | B17 |
| Description of goods | `VBRP-MATNR` + `WERKS` → `MARC-STEUC` → `T604N-TEXT1` (E, IN) | B42 |
| Payment term / Incoterms | `VBRK-ZTERM` → `TVZBT-VTEXT`; `VBRK-INCO1`, `INCO2` | B43, B44 |
| Part No., description, HSN (pre-filled) | `VBRP-MATNR`; `MAKT-MAKTX`; `MARC-STEUC` | B53 (step 53), B54, B58 |
| `ZSD_PACKNO` fields | Empty in step 1, entered by the user in Excel, read in step 4 | B9, B10, B14, B16, B31 – B41, B47 – B57, B61 |

## 3. Assumptions

| # | Assumption | Why |
|---|---|---|
| A1 | The upload file is the ALV export of step 1. Columns are found by their **heading** (ALV text, or the technical field name), so column order and extra columns do not matter. `Packing List No.` is mandatory. | Process step 4 |
| A2 | Cancelled billing documents (`FKSTO` = X) and cancellation documents (`SFAKN` filled) are excluded. | Otherwise the cancellation pair would appear |
| A3 | Fields not in the named table are read from the table that holds them: region text from **T005U** (sheet T005S), country text from **T005T** (sheet T005), `MAKTX` from **MAKT** (sheet MARA). | Required fields unavailable in the named tables |
| A4 | KNVP partner functions SP / SH = internal **AG / WE**; partner = `KNVP-KUNN2`; sales area of the billing document; several partners → lowest `PARZA`. | KNVP structure |
| A5 | Step 1 has **one ALV row per billing item**; header data is read per billing document. Part No. = billed material, description = `MAKT`, HSN = `MARC-STEUC` (logic B54 / B58 "OR"); the user may overwrite them. | Line items come from ZSD_PACKNO, which does not exist yet in step 1 |
| A6 | Description of goods (B42) in step 1: each distinct HSN of **all billing documents of the selection** once, "`TEXT1` - `HSN`", joined with " / ". | "Club common HSN code and description in all billing documents" |
| A7 | Total Articles (B52) is entered by the user. If empty, the form total is the sum of Articles of the packing list. | Totals row of the sample PDFs |
| A8 | Payment term text: first `TVZBT` line of the payment term (language E). | Several `ZTAGG` lines possible |
| A9 | Addresses: latest valid `ADRC` (`DATE_FROM` ≤ today, `NATION` = space); e-mail / phone: default entry, else lowest sequence. | ADRC / ADR2 / ADR6 structure |
| A10 | Step 4: header fields of a packing list = first filled value of its rows (the user may fill them on the first row only). | Convenience; no business rule changed |
| A11 | Step 4 prints the data **as in the file**, without re-reading SAP. | "print the form using the same data" |
| A12 | Notify-party section of the form shown when any Notifier Party is filled (`FORMAT` = N). Sold-to = ship-to → one CONSIGNEE block. | FS 1.1, sample 90000014 |
| A13 | Excel dates: Excel date number, `YYYYMMDD`, `DD.MM.YYYY`, `DD-MM-YYYY`, `DD/MM/YYYY`, `YYYY-MM-DD`. Numbers: decimal point, thousands commas removed. | ALV export formats |
| A14 | Form totals: quantity, gross and net weight = sums of the packing list rows. | Totals row of the sample PDFs |
| A15 | Form rows: an HSN heading row whenever the HSN changes; an empty Articles / weight cell belongs to the package above (merged cell on the samples). | Sample PDFs |
| A16 | Form output: print preview on device `LP01` (constant `GC_DEST`). | Decision of 12.10. |

## 4. Error handling

| Situation | Behaviour |
|---|---|
| Neither sales order nor file entered | Error on the selection screen |
| No sales order / no billing document found | Error message, stop |
| Single sales order that does not exist; sales order without billing document | Entry in the message log |
| Missing address, GST, CIN, partner | Warning in the log, field empty |
| File not readable / not `.xlsx` / empty / no `Packing List No.` column | Error message, stop |
| Row without packing list number | Warning, row not printed |
| Value not convertible (number / date) | Warning with row and column, field empty |
| Packing list without line items; form missing; form error | Error in the log for that packing list, next one continues |

The message log is shown as a popup (step 1 before the ALV, step 4 after printing).

## 5. Objects needed

Program `ZSD_EPACK_REPORT`; for step 4 the Adobe Form `ZSD_EPACK_AF` with interface `ZSD_EPACK_AF_IF`
(`IS_HEADER` TYPE `ZSD_S_EPACK_AF_HDR`, `IT_ITEM` TYPE `ZSD_TT_EPACK_AF_ITM`) as specified in
`src/Adobe_Form_Interface_Structures.md`. Step 1 needs no Z object.

## 6. Performance / clean core

16 bulk SELECTs (JOIN / FOR ALL ENTRIES on de-duplicated keys) before any output loop; hashed / sorted
lookups; no SELECT or SORT inside a loop; no `TABLES` statement; no modification.

## 7. Open points

1. Confirm A2 (cancelled invoices excluded) and A7 (Total Articles = sum when empty).
2. Check in S2A: `J_1BBRANCH-GSTIN`, `T604N` key (`SPRAS`, `LAND1`, `STEUC`), `ADR2-R3_USER` values 1 / 3, device `LP01`.
3. Contact person and FAX on the consignee block are not in the logic sheet — not printed.
