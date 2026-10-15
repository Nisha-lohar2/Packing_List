# WRICEF 102-B Export consolidated packing list — analysis, approach, assumptions, open questions (version 3)

**Sources (source of truth):** revised `Logic_sheet1.xlsx` (sheets *Input*, *Output screen*, *Z table*,
*Data From Excel*, *Data From SAP*, *Validation* — copy in `v3/input/`) and the FS
`Object ID- 102(B)Export consolidated packing list.docx` (V.01) with the sample PDFs in the repository root.
Where FS text and the revised Logic sheet differ, the revised Logic sheet is followed (it is newer).

**Status:** code written and checked offline (balanced blocks, every `PERFORM` has a `FORM`, no `SELECT`
inside a loop). **Not yet compiled or run in S2A.** Points marked *SAP-VERIFY* must be checked at activation.

---

## 1. Analysis of the files

| File / sheet | What it says | Used for |
|---|---|---|
| Logic sheet — *Input* | T-code `ZSD_EPACK`; **sales order from / to, multiple selection** | Report 1 selection screen |
| *Output screen* | 43 columns (A – AQ) and a 77-line sample with totals | Content check; column **order is not used** (see A02) |
| *Z table* | `ZSD_EPACK_DATA`: 46 columns B – AU with heading, technical name, length, type; keys "Packing list no" and "SAP Sales Order No" | Table, ALV, Excel layout (order and headings) |
| *Data From Excel* | Fields entered by the user (header, line item, footer) | Editable ALV columns |
| *Data From SAP* | Logic of the SAP fields per sales order | Report 1 data retrieval |
| *Validation* | Uploaded quantity ≤ delivered quantity (VBAP, LIPS by VGBEL / VGPOS) | Report 2, V16 |
| FS 1.1 | Two Smart Form formats (with / without notifier party, up to 5, unused crossed out); user selects the format when NP are maintained | Report 2 selection + Smart Forms |
| FS "Additional requirement" | GW < NW → message "Invalid Weight: Gross Weight cannot be less than Net Weight.", no form output | Report 1 (ALV check), report 2 (V17, V18) |
| FS manual/SAP table, embedded workbooks | Still the **old invoice-based** version (VBRK_VBELN, advance amount, commercial terms, invoice declaration) | Not used — superseded by the revised sheet |
| FS messages / authorization sections | Empty | Standard messages; standard authorization `V_VBAK_VKO` (A17) |
| Sample PDFs 90000014 (without NP), 90000081 (with NP) | Approved layout | Smart Form layout |

## 2. Approach

```
Report 1  ZSD_EPACK_EXTRACT (T-code ZSD_EPACK)
  sales orders ──► bulk read (VBAK/TVKO, VBAP, VBKD, T001, ADRC, ADR6, ADR2, T005U, T005T,
                   T001Z, T001W, J_1BBRANCH, KNA1, KNVP, TVZBT, MARC, MAKT)
              ──► editable ALV (layout = sheet "Z table"; manual columns editable, SAP columns protected,
                   DATA_CHANGED checks GW ≥ NW)
              ──► "Download Excel": CHECK_CHANGED_DATA, .xlsx with headings/order of sheet "Z table"
                   (no database update)
Report 2  ZSD_EPACK_UPLOAD
  .xlsx ──► path / format / headings by position ──► rows typed like ZSD_EPACK_DATA (+ Excel row no.)
       ──► validations V01 – V25, errors with row and column
       ──► save (all-or-nothing default / per packing list), lock, COMMIT / ROLLBACK, overwrite only on request
       ──► Smart Form per saved packing list (format chosen on the selection screen)
       ──► result list (summary + all messages)
```

Both reports share the include `ZSD_EPACK_COMMON`: the Excel layout is defined **once**, so the file of
report 1 always matches what report 2 expects. Procedural (FORM routines); standard classes only.
Clean core: no modification, no enhancement, standard APIs, bulk reads.

## 3. Objects

See `DDIC_and_SmartForm.md` §1: 32 data elements, table `ZSD_EPACK_DATA`, lock object `EZSD_EPACK`,
2 structures + 1 table type, 2 Smart Forms, include `ZSD_EPACK_COMMON`, reports `ZSD_EPACK_EXTRACT` and
`ZSD_EPACK_UPLOAD`, GUI status `ZEPACK_STD`, 2 T-codes.

## 4. Assumptions applied (code references `Axx`)

| # | Assumption | Reason |
|---|---|---|
| A01 | Input = **sales orders** (revised sheet *Input*), not tax invoices (FS text). | Revised Logic sheet is the source of truth |
| A02 | Column order and headings of ALV and Excel = sheet **"Z table"** (B – AU), not sheet *Output screen*. Headings are taken as written; the line break in "Packing list no / KEY FIELD" is a space. Report 2 also accepts the key headings without "KEY FIELD". | Z-table format defines table, ALV and download |
| A03 | **`ZSD_SN` (Sl.No.) is added to the table key** (`MANDT`, `ZSD_PACKNO`, `VBAK_VBELN`, `ZSD_SN`) and is mandatory in the upload. | The sheet keys (packing list + sales order) allow only one row per sales order, but a packing list has many lines per order (sample: 77 lines). Sl.No. is unique per packing list in the sample. |
| A04 | Sheet column Q (2nd "Packing list no", same technical name `ZSD_PACKNO`) is written to the Excel file with the same value as column B and checked on upload (must be empty or equal). It is not a second table field. | A table cannot contain the same field twice |
| A05 | `VBAK_VBELN` (CHAR 255) holds the sales order of the row; several orders may be entered separated by `,` `;` or blanks. Stored in internal format (leading zeros), sorted, `,`-separated. Report 1 fills one order per row (external format in Excel). | Length 255 in the sheet = several orders possible |
| A06 | Quantity check: Part No. → material; uploaded quantities are **added up per sales order(s) + material over all rows of the file** and compared with the sum of `LIPS-LFIMG` of all delivery items with `VGBEL` = sales order and the same material. Rows with quantity 0 are not checked. Quantities are compared without unit conversion. | Sheet *Validation*; the upload has no item number (POSNR) and a material can be split over several packages (sample) |
| A07 | Header values of a packing list (all non-item columns) are taken from the first row (lowest Sl.No.) that has a value; rows with a different value get a warning. | Header data is repeated on every row |
| A08 | A packing list already in `ZSD_EPACK_DATA` is **rejected** unless "Overwrite" is ticked; with Overwrite its old rows are deleted and the new rows inserted in the same LUW. → A rerun with the same file never creates duplicates. | No rule in FS / sheet |
| A09 | Save mode: **all-or-nothing (default)** — nothing is saved or printed if any error exists; alternative "per packing list" — packing lists without errors are saved (one LUW each), others are not. Never a part of a packing list. **Test run is the default.** | Not specified — to confirm (Q01) |
| A10 | GST No.: VBAK has no `BUPLA`. The business place is taken from the **item plant** (`T001W-J_1BBRANCH`) → `J_1BBRANCH-GSTIN` with the company code of the sales organization. *SAP-VERIFY* | Sheet names a field that does not exist (Q04) |
| A11 | Fields read from the table that holds them: region text `T005U-BEZEI` (sheet T005S), country text `T005T-LANDX` (sheet T005), `MAKT-MAKTX` (sheet MARA). Language `E`. | Fields not in the named tables |
| A12 | Address = latest valid `ADRC` version (`DATE_FROM` ≤ today, `NATION` blank); e-mail / telephone / mobile = default entry, else lowest sequence. | ADRC / ADR6 / ADR2 structure |
| A13 | Ship-to as in the sheet: customer master `KNVP` (sold-to, sales area of the order, `PARVW` SH = internal `WE`) → `KUNN2` → `KNA1`; several → lowest `PARZA`. **Not** the order partner `VBPA`. | Sheet text (Q05) |
| A14 | Exporter (CHAR 100): fixed positions `name, street, street 2, city, postal code, region, country`. Sold-to / ship-to (CHAR 150): `name, street, street 2, city, district, postal code, country, e-mail, telephone, mobile`. Missing parts stay empty between the commas; commas inside a part are replaced by a space; too long → cut off with a warning. The Smart Form gets the parts split and prints the labels "Telephone" / "Mobile number". | Previous decision on address columns; columns of the sheet are single fields |
| A15 | Technical fields added to the table: `ZSD_MEINS` (unit of `ZSD_QTY`, from `VBAP-VRKME`, default `ST`), `ZSD_GEWEI` (`KG`), `ERNAM`, `ERDAT`, `ERZET`. `ZSD_PL_DECL` = STRING (max. 1500), because CHAR table fields are limited to 1333. | QUAN needs a unit field; DDIC limit |
| A16 | Report 1: one ALV row per sales order item (all `POSNR`, including rejected items). Rows cannot be inserted or deleted in the ALV (they are the SAP items); copying / splitting rows is done in Excel. | Sheet "for all POSNR" (Q07) |
| A17 | Authorization: display check `V_VBAK_VKO` (activity 03) per sales area in report 1; report 2 only by its T-code. | FS authorization section empty (Q10) |
| A18 | Description of goods (`ZSD_HSN_DESC`) and Total Articles (`ZSD_TART`) are **manual** (sheet *Data From Excel*); nothing pre-filled. | Revised sheet |
| A19 | SAP columns may be changed in Excel; report 2 saves and prints the uploaded values (no re-read from SAP). | The uploaded file is the source of the Z table and the form (FS: data maintained through the Excel upload) |
| A20 | Excel dates accepted as Excel date number, `DD.MM.YYYY`, `DD-MM-YYYY`, `DD/MM/YYYY`, `YYYY-MM-DD`, `YYYYMMDD`; numbers with decimal point, thousands commas removed. Report 1 writes dates `DD.MM.YYYY`, zero → empty cell. | Excel behaviour |

## 5. Open questions — please confirm

| # | Question | Current implementation |
|---|---|---|
| Q01 | Save: all-or-nothing, or save correct packing lists only? | Both on the selection screen, default all-or-nothing (A09) |
| Q02 | T-code for report 2 (the Input sheet names only `ZSD_EPACK`) | `ZSD_EPACK` = report 1, `ZSD_EPACK_UPL` = report 2 |
| Q03 | Table key: may Sl.No. be added to the key? Should "SAP Sales Order No" really be CHAR 255 in the key (one order per row would be `VBELN_VA` CHAR 10)? | A03, A05 |
| Q04 | GST No.: which business place? (VBAK has no `BUPLA`) | Plant business place (A10) |
| Q05 | Ship-to: customer master `KNVP` (as in the sheet) or the ship-to of the sales order (`VBPA`)? | KNVP (A13) |
| Q06 | Total Articles on the form: the entered `ZSD_TART` or the sum of Articles? | Both passed to the form (`TART`, `TOT_ART`) |
| Q07 | Rejected sales order items / items with quantity 0 — include in report 1? | Included (A16) |
| Q08 | Quantity check: also count quantities of packing lists **already saved** for the same order + material? Only deliveries with goods issue posted (PGI), or all deliveries? Batch-split sub-items? | File only; all delivery items with `VGBEL` = order (A06). *SAP-VERIFY* batch split: if main item and sub-items both carry `LFIMG`, exclude sub-items (`UECHA` ≠ 0) |
| Q09 | Smart Form names | `ZSD_EPACK_SF` (without NP), `ZSD_EPACK_SF_NP` (with NP) |
| Q10 | Authorization for upload / save | T-code only (A17) |
| Q11 | Should the **Part No. must be an item of the sales order** check (V15) be an error or a warning? | Error — derived from the VBAP step of sheet *Validation* |
| Q12 | Sheet *Data From Excel* footer names `ZSD_DECL`; the Z table names `ZSD_PL_DECL` | `ZSD_PL_DECL` |
| Q13 | Terms of shipment `INCO1` + `INCO2` (up to 31 characters) in a CHAR 20 field; Order No `BSTKD` (35) in CHAR 20 | Cut off with a warning in report 1 |

## 6. Not in scope / not done

* No Adobe Form any more (FS: Smart Form). The earlier `v2` program and its Adobe Form stay in the repository unchanged.
* The FS embedded workbooks (old invoice-based Z table with advance amount, commercial terms, invoice declaration) are not used.
