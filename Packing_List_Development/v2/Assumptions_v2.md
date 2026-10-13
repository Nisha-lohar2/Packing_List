# ZSD_EPACK_REPORT — relationships, assumptions and open points (version 2)

Report: `v2/ZSD_EPACK_REPORT.prog.abap` (new development, replaces the earlier versions).
Sources: FS `Object ID- 102(B)Export consolidated packing list.docx`, `Logic_sheet1.xlsx`
(rows Bxx below), sample upload file `test_data/Sample_data_ABAP.xlsx`.

**Status:** written and checked offline (balanced blocks, every PERFORM has a FORM, no SELECT
or SORT inside a loop). **Not compiled or run in S2A yet.**

---

## 1. Selection screen

| Field | Type | Use |
|---|---|---|
| `S_VBELN` | SELECT-OPTIONS on `VBAK-VBELN`, mandatory | Sales orders |
| `P_FILE` | File path, mandatory, F4 file browser | Packing list Excel (`.xlsx`) |

No other fields.

## 2. Relationships — validated before coding

| Link | Implemented as | Basis |
|---|---|---|
| Sales order → billing documents | `VBRP-AUBEL IN S_VBELN`, joined to `VBRK` on `VBELN` | `AUBEL` is the sales order of the billing item; the logic sheet itself uses `AUBEL` (B13) |
| Packing list ↔ billing documents | Excel column B "SAP Invoice Number" = `VBRK-VBELN` (one or several per cell, e.g. `1000000348, 1000000349, 1000000350`) | Sample file; FS: "packing list for multiple tax invoices" |
| Packing list ↔ sales orders | A packing list is reported when **at least one** of its invoices belongs to the entered sales orders | Result of the two links above |
| `ZSD_PACKNO` | The uploaded Excel. "Pass ZSD_PACKNO get X" = value X of that packing list in the file | Logic sheet B9 – B61; the file has exactly these fields |
| Exporter | `VBRK-BUKRS` → `T001-ADRNR` → `ADRC` | B3 |
| GST No. | `VBRK-BUKRS` + `VBRK-BUPLA` → `J_1BBRANCH-GSTIN` (`BRANCH` = `BUPLA`) | B6 |
| CIN | `T001Z-PAVAL` for `BUKRS`, `PARTY` = `ZCIN` | B7 |
| IEC | `T001-ADRNR` → `ADRC-BUILDING` | B11 |
| Order no. & date | `VBRP-AUBEL` → `VBAK-VGBEL` → `VBKD` (`VBELN` = `VGBEL`, `POSNR` = `000000`) `BSTKD`, `BSTDK` | B13 ("Quotation – customer reference and date") |
| Sold-to / ship-to | `VBRK-KUNAG` + sales area of the billing document → `KNVP` (`PARVW` SP / SH) → `KUNN2` → `KNA1-ADRNR` → `ADRC`, `ADR6`, `ADR2` | B20 – B29 |
| Description of goods | `VBRP-MATNR` + `VBRP-WERKS` → `MARC-STEUC` → `T604N-TEXT1` (`SPRAS` E, `LAND1` IN) | B42 |
| Payment term | `VBRK-ZTERM` → `TVZBT-VTEXT` | B43 |
| Terms of shipment | `VBRK-INCO1`, `VBRK-INCO2` | B44 |

## 3. Assumptions

| # | Assumption | Why |
|---|---|---|
| A1 | The Excel layout is the one of `Sample_data_ABAP.xlsx`: row 1 headings, columns A – AK in this order (A Packing list Number, B SAP Invoice Number, C Sl.No., … AH Declaration of Packing list). Columns are read **by position**. | Provided file |
| A2 | Cancelled billing documents (`VBRK-FKSTO` = X) and cancellation documents (`VBRK-SFAKN` filled) are excluded. Otherwise the cancellation pair of an invoice would be reported as well. | Standard SD behaviour; not stated in the FS |
| A3 | **Fields not in the named table** — the table that holds the field is read: region text `BEZEI` from **T005U** (sheet: T005S, which has no text); country text `LANDX` from **T005T** (sheet: T005, which has no text); `MAKTX` from **MAKT** (sheet: MARA, which has no description). | Required fields unavailable in the named tables |
| A4 | KNVP stores the partner functions SP / SH with the internal codes **AG / WE**. "GET KUNNR" = the partner number **`KNVP-KUNN2`** (KNVP-KUNNR is the customer passed in). The sales area of the billing document is used (KNVP key). Several partners of one function: the one with the lowest counter `PARZA`. | KNVP structure |
| A5 | VBRK header values (exporter, GST, CIN, IEC, sold-to, ship-to, payment term, Incoterms) are taken from the **first** billing document of the packing list (lowest number). The FS has one value per packing list; differing values across its invoices are not checked. | One value per printed field |
| A6 | Header fields of the file (packing list date, buyer's reference, notifiers, shipment fields, declaration …) are taken from the **first filled** cell of the packing list's rows. In the sample they are filled on the first row only. Invoice numbers are collected from all rows of the packing list. A row without packing list number belongs to the packing list above. | Sample file |
| A7 | **Total Articles (B52, `ZSD_TART`)**: the Excel file has no column for it → the output column stays **empty**. No calculation is invented. | Logic: "Pass ZSD_PACKNO get ZSD_TART", field not in the file |
| A8 | Product description (B54): file value, else `MAKT-MAKTX` (language E) of the Part No. HSN (B58): file value, else `MARC-STEUC` of the Part No. in the plant where it is billed in the packing list invoices. Part No. = material number (`CONVERSION_EXIT_MATN1_INPUT`). Cell texts starting with "Fetch from …" count as empty. | Logic B54 / B58 "OR"; placeholder texts of the sample |
| A9 | Description of goods: each **distinct** HSN of all billing items of the packing list's invoices once, "`<T604N-TEXT1>` - `<HSN>`", joined with " / ". Excel column AG is not used (B42 has SAP logic). | Logic B42 "club common HSN code and description in all billing documents" |
| A10 | Exporter Ref. (B11) is the IEC from `ADRC-BUILDING`; Excel column O is not used. Columns AI – AK (advance amount, commercial terms, invoice declaration) are commercial-invoice fields, not in the packing list logic, and are not output. | Logic sheet |
| A11 | Order no. & date: all distinct values of the packing list's sales orders, joined with " & ". Dates DD.MM.YYYY. | Several sales orders per packing list possible |
| A12 | Addresses: latest valid `ADRC` version (`DATE_FROM` ≤ today, international version `NATION` = space). E-mail / telephone / mobile: default entry (`FLGDEFAULT`), else lowest sequence. Texts in language E. | ADRC / ADR2 / ADR6 structure |
| A13 | Excel dates: Excel date number or `YYYY-MM-DD`, `DD.MM.YYYY`, `DD-MM-YYYY`, `DD/MM/YYYY`. Numbers: decimal point, thousands commas removed. | Sample file |

## 4. Error handling

| Situation | Behaviour |
|---|---|
| File cannot be read / not `.xlsx` / empty | Error message, stop |
| No sales order found for the selection | Error message, stop |
| Single sales order entered that does not exist | Error in the message log |
| Sales order without billing document | Warning in the log |
| No billing document for all entered sales orders | Error message, stop |
| Invoice of the file not belonging to the entered sales orders | Warning, invoice ignored; packing list skipped if none belongs |
| Packing list without invoice number / without line items | Error in the log, packing list skipped |
| Missing address, GST, CIN, partner, quotation reference, HSN, description | Warning in the log, field empty |
| Value not convertible (number / date) | Warning in the log, field empty |

The message log is shown in a popup before the result list.

## 5. Output (ALV grid, standard functions incl. export to spreadsheet)

One row per packing list line item; header values repeated on each row. Column order:
reference (sales orders, invoices) → exporter (B3 – B7) → packing list no., date, IEC, order
no. & date, buyer's reference, country of origin / destination (B9 – B17) → sold-to (B20 – B23)
→ ship-to (B26 – B29) → notifier 1 – 5 (B31) → shipment (B33 – B41) → description of goods,
payment term, Incoterms (B42 – B44) → line item (B47 – B58) → declaration (B61).

## 6. Performance / clean core

* 16 bulk SELECTs per run (JOIN or FOR ALL ENTRIES on de-duplicated keys), all before the output
  loop; lookups on hashed / sorted tables. No SELECT and no SORT inside a loop.
* No `TABLES` statement, no modification. Classes used: `CL_GUI_FRONTEND_SERVICES`,
  `CL_FDT_XL_SPREADSHEET`, `CL_BCS_CONVERT`, `CL_SALV_TABLE`.
* `VBRP-AUBEL` is not a standard index field; on HANA this is acceptable. On a high-volume
  system an access via `VBFA` (`VBELV` = sales order) can be used instead — same result.

## 7. Open points for the business

1. A7 — Total Articles: add a column to the file, or may it be calculated (Σ Articles)?
2. A5 — may one packing list contain invoices with different company code / customer / payment terms?
3. A2 — confirm exclusion of cancelled invoices.
4. Check in S2A: `J_1BBRANCH-GSTIN`, `T604N` key (`SPRAS`, `LAND1`, `STEUC`), `TVZBT-ZTAGG`, `ADR2-R3_USER` values 1 / 3.
