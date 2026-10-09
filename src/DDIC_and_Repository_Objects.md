# WRICEF 102-B "Export Consolidated Packing List" — DDIC & repository objects

Package: **`ZSD_EPACK`** — all objects new. No SAP object is modified,
appended or enhanced anywhere in this development.

Assumption references (`Axx`) point to `docs/Development_Assumptions.md`;
clarification references (`CL-xx`) to
`docs/Packing_List_Requirement_Analysis_and_Development_Approach.md`.

---

## 0. Deviation from the FS data model (read first)

The FS names one table, `ZSD_EPACK_DATA`, keyed on `ZSD_PACKNO` +
`VBRK_VBELN` (CHAR 255, "store multiple tax invoice"). That design cannot hold
more than one line item per packing list and cannot be searched by invoice
(analysis §6.5, CL-01). It is therefore split into three tables — **the FS
name `ZSD_EPACK_DATA` is kept for the line items**:

| FS | Implemented | Why |
|---|---|---|
| `ZSD_EPACK_DATA` (header + items + invoices in one row) | `ZSD_EPACK_HDR` — one row per packing list | Header data stored once |
| `VBRK_VBELN` CHAR 255 (several invoices concatenated) | `ZSD_EPACK_INV` — one row per invoice | Selection by invoice (FS input), index, validation |
| — | `ZSD_EPACK_DATA` — one row per line item, key `ZSD_PACKNO` + `ZSD_SN` | Many lines per packing list |
| `ZSD_TART` (Total Articles) | not stored — calculated | Sample totals = Σ Articles (A08) |
| `ZSD_VSl_FLT` | `ZSD_VSL_FLT` | DDIC names are upper case |
| NP / container / packages / goods / declaration as CHAR 30…1500 | `SSTRING` / `STRING` | Sample texts exceed the FS lengths; CHAR 1500 exceeds the table limit (A19) |
| — | Package-group fields `ZSD_ART_GRP`, `ZSD_WT_GRP`, `ZSD_PKG_TYPE` | Merged cells / subtotals of the samples (A06, A09) |
| — | Audit / revision / print fields | Traceability, "Rev.1" sample (A22) |

**This deviation needs functional sign-off (CL-01).**

---

## 1. Domains and data elements — new

Data element name = field name (FS "Tech name"), so that the FS template
`Z table Format1.xlsx › Technical detail` maps 1:1.

| Data element | Domain / type | Len | Dec | Description | FS |
|---|---|---:|---:|---|---|
| `ZSD_PACKNO` | `ZSD_PACKNO` CHAR, conversion `ALPHA`, value table `ZSD_EPACK_HDR` | 10 | | Packing list number | ✔ |
| `ZSD_PACKDT` | `DATS` | 8 | | Packing list date | ✔ |
| `ZSD_REVNO` | NUMC | 3 | | Revision (0 = original) | new |
| `ZSD_BUKRS` | ref. `BUKRS` | 4 | | Company code of the invoices | new |
| `ZSD_EXPREF` | CHAR | 20 | | Exporter's ref. — kept for the FS structure, **not filled** (Technical detail: "Not Req – fetched from SAP") | ✔ |
| `ZSD_PARTY_REF` | CHAR | 200 | | Buyer's reference no. | ✔ |
| `ZSD_REF_DT` | `DATS` | 8 | | Buyer's reference date | ✔ |
| `ZSD_CNTY_ORGN` | CHAR | 20 | | Country of origin of goods | ✔ |
| `ZSD_NP` | `SSTRING` | 1333 | | Notify party (multi-line). Used for `ZSD_NP1`…`ZSD_NP5`. | ✔ (FS 200) |
| `ZSD_PREC` | CHAR | 30 | | Pre-carriage | ✔ |
| `ZSD_REC_PC` | CHAR | 30 | | Place of receipt by pre-carrier | ✔ |
| `ZSD_POL` / `ZSD_POD` / `ZSD_PLD` | CHAR | 30 | | Port of loading / discharge / place of delivery | ✔ |
| `ZSD_VSL_FLT` | CHAR | 30 | | Vessel / flight no. | ✔ |
| `ZSD_MARK` | `SSTRING` | 255 | | Marks & Nos. | ✔ (FS 30) |
| `ZSD_CONT` | `SSTRING` | 1333 | | Container no(s). | ✔ (FS 100) |
| `ZSD_NOPACK` | `SSTRING` | 255 | | No. of packages (text) | ✔ (FS 30) |
| `ZSD_HSN_DESC` | `SSTRING` | 1333 | | Description of goods | ✔ (FS 100) |
| `ZSD_PL_DECL` | `STRING` | | | Declaration of packing list | ✔ (FS 1500) |
| `ZSD_ADV_AMT` | CURR | 15 | 2 | Advance amount received (commercial invoice) | ✔ |
| `ZSD_CMMT` | `SSTRING` | 1333 | | Commercial terms of invoice | ✔ (FS 500) |
| `ZSD_INV_DECL` | `STRING` | | | Declaration of invoice | ✔ (FS 1500) |
| `ZSD_FILENAME` | `SSTRING` | 255 | | Uploaded file name | new |
| `ZSD_SN` | NUMC | 6 | | Sl.No. | ✔ |
| `ZSD_INV_SN` | NUMC | 6 | | Sr.No. of commercial-invoice line | ✔ |
| `ZSD_ARTNO` | CHAR | 15 | | Article no. ("1 TO 60", "STACKER") | ✔ |
| `ZSD_STD` | NUMC | 6 | | STD (pieces per article) | ✔ |
| `ZSD_ART` | NUMC | 6 | | Articles (first line of article group only) | ✔ |
| `ZSD_ART_GRP` | NUMC | 4 | | Article group (merged cells) | new |
| `ZSD_WT_GRP` | NUMC | 4 | | Weight group (merged cells) | new |
| `ZSD_PKG_TYPE` | CHAR | 20 | | Package type (Loose Pipes / Boxes / Bundles) | new |
| `ZSD_PARTNO` | CHAR | 40 | | Part no. (= material no.) | ✔ |
| `ZSD_MATDESC` | CHAR | 40 | | Product description | ✔ |
| `ZSD_QTY` | QUAN | 13 | 3 | Quantity | ✔ |
| `ZSD_GWT` / `ZSD_NWT` | QUAN | 13 | 3 | Gross / net weight (first line of weight group only) | ✔ |
| `ZSD_QTY_UOM` | ref. `MEINS` | 3 | | Unit of quantity (`ST`) | new |
| `ZSD_WT_UOM` | ref. `GEWEI` | 3 | | Unit of weight (`KG`) | new |
| `ZSD_HSN` | CHAR | 17 | | H.S. code | ✔ |
| `ZSD_EPACK_BLOCK` | CHAR | 10 | | Form text block id | new |
| `ZSD_EPACK_ROW_TYPE` | CHAR, fixed values `H` heading / `I` item / `S` subtotal | 1 | | Printed row type | new |
| `ZSD_EPACK_STATUS` | CHAR, fixed values `S` success / `W` warning / `E` error | 1 | | Processing status (log) | new |

Audit fields reuse standard data elements: `ZSD_ERNAM`/`ZSD_AENAM`/`ZSD_PRNAM` → `ERNAM`/`AENAM`/`USNAM`,
`*DAT` → `DATS`, `*ZET` → `TIMS`.

> ⚠ **SAP-VERIFY** — the maximum length of `SSTRING` in a transparent table
> is 1333 characters; tables containing `STRING` fields cannot be buffered
> (not needed here).

---

## 2. Transparent tables — new

Common technical settings: delivery class **`A`**, data class `APPL1`,
size category `0`, **no buffering**, log data changes **yes**,
table maintenance **not allowed** (written only by `ZSD_EPACK`).

### 2.1 `ZSD_EPACK_HDR` — Packing list header

| # | Field | Key | Data element | Ref. field |
|---|---|:---:|---|---|
| 1 | `MANDT` | ✔ | `MANDT` | |
| 2 | `ZSD_PACKNO` | ✔ | `ZSD_PACKNO` | |
| 3 | `ZSD_PACKDT` | | `ZSD_PACKDT` | |
| 4 | `ZSD_REVNO` | | `ZSD_REVNO` | |
| 5 | `ZSD_BUKRS` | | `BUKRS` | |
| 6 | `ZSD_EXPREF` | | `ZSD_EXPREF` | |
| 7 | `ZSD_PARTY_REF` | | `ZSD_PARTY_REF` | |
| 8 | `ZSD_REF_DT` | | `ZSD_REF_DT` | |
| 9 | `ZSD_CNTY_ORGN` | | `ZSD_CNTY_ORGN` | |
| 10–14 | `ZSD_NP1` … `ZSD_NP5` | | `ZSD_NP` | |
| 15 | `ZSD_PREC` | | `ZSD_PREC` | |
| 16 | `ZSD_REC_PC` | | `ZSD_REC_PC` | |
| 17 | `ZSD_VSL_FLT` | | `ZSD_VSL_FLT` | |
| 18 | `ZSD_POL` | | `ZSD_POL` | |
| 19 | `ZSD_POD` | | `ZSD_POD` | |
| 20 | `ZSD_PLD` | | `ZSD_PLD` | |
| 21 | `ZSD_MARK` | | `ZSD_MARK` | |
| 22 | `ZSD_CONT` | | `ZSD_CONT` | |
| 23 | `ZSD_NOPACK` | | `ZSD_NOPACK` | |
| 24 | `ZSD_HSN_DESC` | | `ZSD_HSN_DESC` | |
| 25 | `ZSD_PL_DECL` | | `ZSD_PL_DECL` | |
| 26 | `ZSD_ADV_AMT` | | `ZSD_ADV_AMT` | `ZSD_ADV_WAERS` |
| 27 | `ZSD_ADV_WAERS` | | `WAERS` | |
| 28 | `ZSD_CMMT` | | `ZSD_CMMT` | |
| 29 | `ZSD_INV_DECL` | | `ZSD_INV_DECL` | |
| 30 | `ZSD_FILENAME` | | `ZSD_FILENAME` | |
| 31–33 | `ZSD_ERNAM`, `ZSD_ERDAT`, `ZSD_ERZET` | | `ERNAM`, `DATS`, `TIMS` | |
| 34–36 | `ZSD_AENAM`, `ZSD_AEDAT`, `ZSD_AEZET` | | `AENAM`, `DATS`, `TIMS` | |
| 37–39 | `ZSD_PRNAM`, `ZSD_PRDAT`, `ZSD_PRZET` | | `USNAM`, `DATS`, `TIMS` | |

Foreign key `ZSD_BUKRS` → `T001`.
`ZSD_ADV_AMT`, `ZSD_CMMT` and `ZSD_INV_DECL` are commercial-invoice fields from the FS
template. They are stored but not printed on the packing list (CL-17).

### 2.2 `ZSD_EPACK_INV` — Packing list ↔ tax invoice

| # | Field | Key | Data element |
|---|---|:---:|---|
| 1 | `MANDT` | ✔ | `MANDT` |
| 2 | `ZSD_PACKNO` | ✔ | `ZSD_PACKNO` |
| 3 | `VBELN` | ✔ | `VBELN_VF` |

* Foreign keys: `ZSD_PACKNO` → `ZSD_EPACK_HDR`, `VBELN` → `VBRK`.
* **Secondary index `Z01`**: `MANDT`, `VBELN` (non-unique). Used for the
  selection by invoice in `ZSD_EPACK`.
* An invoice may belong to **one** packing list only. This is enforced by the
  upload, not by a unique index, so that the error message names the
  conflicting packing list (A01).

### 2.3 `ZSD_EPACK_DATA` — Packing list line items (FS table name)

| # | Field | Key | Data element | Ref. field |
|---|---|:---:|---|---|
| 1 | `MANDT` | ✔ | `MANDT` | |
| 2 | `ZSD_PACKNO` | ✔ | `ZSD_PACKNO` | |
| 3 | `ZSD_SN` | ✔ | `ZSD_SN` | |
| 4 | `ZSD_INV_SN` | | `ZSD_INV_SN` | |
| 5 | `ZSD_ARTNO` | | `ZSD_ARTNO` | |
| 6 | `ZSD_STD` | | `ZSD_STD` | |
| 7 | `ZSD_ART` | | `ZSD_ART` | |
| 8 | `ZSD_ART_GRP` | | `ZSD_ART_GRP` | |
| 9 | `ZSD_WT_GRP` | | `ZSD_WT_GRP` | |
| 10 | `ZSD_PKG_TYPE` | | `ZSD_PKG_TYPE` | |
| 11 | `ZSD_PARTNO` | | `ZSD_PARTNO` | |
| 12 | `ZSD_MATDESC` | | `ZSD_MATDESC` | |
| 13 | `ZSD_QTY` | | `ZSD_QTY` | `ZSD_QTY_UOM` |
| 14 | `ZSD_QTY_UOM` | | `MEINS` | |
| 15 | `ZSD_GWT` | | `ZSD_GWT` | `ZSD_WT_UOM` |
| 16 | `ZSD_NWT` | | `ZSD_NWT` | `ZSD_WT_UOM` |
| 17 | `ZSD_WT_UOM` | | `GEWEI` | |
| 18 | `ZSD_HSN` | | `ZSD_HSN` | |

Foreign key `ZSD_PACKNO` → `ZSD_EPACK_HDR`.

### 2.4 `ZSD_EPACK_LOG` — processing log (one row per packing list per run)

Written by `ZCL_SD_EPACK_STORE->WRITE_LOG` at the end of every run of
`ZSD_EPACK`, except "Validate only". It covers successful and failed packing lists.
File-level problems get one extra row with an empty packing list number.
Delivery class `A`, no buffering. **Request: "at last add the data in the Z table that was created for log"** (Q16).

| # | Field | Key | Data element / type | Content |
|---|---|:---:|---|---|
| 1 | `MANDT` | ✔ | `MANDT` | |
| 2 | `RUNID` | ✔ | `SYSUUID_C32` | One id per program run |
| 3 | `ZSD_PACKNO` | ✔ | `ZSD_PACKNO` | Packing list (empty = file-level row) |
| 4 | `SOURCE` | | CHAR 1 | `U` Excel upload / `R` reprint of saved data |
| 5 | `FILENAME` | | `ZSD_FILENAME` | Uploaded file |
| 6 | `XLS_ROWS` | | CHAR 20 | Excel rows of this packing list, e.g. "2-78" |
| 7 | `INVOICES` | | `SSTRING` 1333 | Invoice numbers, comma-separated |
| 8 | `INV_CNT` | | `INT4` | Number of invoices |
| 9 | `ITEM_CNT` | | `INT4` | Number of line items |
| 10 | `FORMNAME` | | `TDSFNAME` | Smart Form used |
| 11 | `OUTPUT` | | CHAR 1 | `V` preview / `P` print / `F` PDF |
| 12 | `OUTPUT_DONE` | | `XFELD` | Form produced |
| 13 | `SAVED` | | `XFELD` | Data written to `ZSD_EPACK_HDR/INV/DATA` |
| 14 | `STATUS` | | `ZSD_EPACK_STATUS` | `S` / `W` / `E` |
| 15 | `MSG_TYPE` | | `BAPI_MTYPE` | Type of the most important message |
| 16 | `MESSAGE` | | `BAPI_MSG` | First error, else first warning, else last success text |
| 17 | `MSG_CNT` | | `INT4` | Number of messages |
| 18 | `ERNAM` / `ERDAT` / `ERZET` | | `ERNAM` / `DATS` / `TIMS` | Who / when |

Secondary index `Z01`: `MANDT`, `ZSD_PACKNO`. Display it with SE16N or the table view (no maintenance).

---

## 3. Structures and table types — Smart Form interface

Both Smart Forms have the **same** form interface:

| Parameter | Type | Content |
|---|---|---|
| `IS_HEADER` | `ZSD_S_EPACK_PRT_HDR` | Single values of the header |
| `IT_TEXT` | `ZSD_TT_EPACK_PRT_TXT` | All multi-line blocks (addresses, notify parties, …) |
| `IT_ITEM` | `ZSD_TT_EPACK_PRT_ITM` | Printed item rows incl. headings and subtotals |

### 3.1 `ZSD_S_EPACK_PRT_HDR`

| Component | Type | Content |
|---|---|---|
| `PACKNO` | `ZSD_PACKNO` | Packing List No. |
| `PACKDT_TXT` | CHAR 10 | Date, DD-MM-YYYY |
| `REVNO` | `ZSD_REVNO` | Revision (not printed by default, A22) |
| `IEC` | CHAR 40 | Exporter's Ref. "IEC:…" |
| `GSTIN` | CHAR 18 | GST No. |
| `CIN` | CHAR 40 | CIN No. |
| `CNTY_ORGN` | CHAR 20 | Country of origin |
| `CNTY_DEST` | CHAR 50 | Country of final destination (ship-to) |
| `SAME_PARTY` | `XFELD` | Sold-to = ship-to → one "CONSIGNEE" block (A15, Q09) |
| `PREC`, `REC_PC`, `VSL_FLT`, `POL`, `POD`, `PLD` | CHAR 30 | Shipping fields |
| `NP_COUNT` | `INT1` | Number of filled notify parties (0–5) |
| `TOT_ART_TXT` | CHAR 10 | Total articles / packages |
| `TOT_QTY_TXT` | CHAR 17 | Total quantity |
| `TOT_GWT_TXT`, `TOT_NWT_TXT` | CHAR 17 | Total gross / net weight, 3 decimals |
| `TOT_TEXT` | CHAR 132 | "(Total 2342 Loose Pipes + 1073 Boxes = 3415 Packages)" |

### 3.2 `ZSD_S_EPACK_PRT_TXT` / `ZSD_TT_EPACK_PRT_TXT` (standard table)

| Component | Type |
|---|---|
| `BLOCK` | `ZSD_EPACK_BLOCK` — `EXPORTER`, `SOLDTO`, `SHIPTO`, `NP1`…`NP5`, `ORDERREF`, `BUYERREF`, `PAYMENT`, `MARKS`, `CONTAINER`, `PACKAGES`, `GOODS`, `DECL` |
| `LINE_NO` | NUMC 3 |
| `TEXT` | CHAR 255 |

### 3.3 `ZSD_S_EPACK_PRT_ITM` / `ZSD_TT_EPACK_PRT_ITM` (standard table)

| Component | Type | Content |
|---|---|---|
| `ROW_TYPE` | `ZSD_EPACK_ROW_TYPE` | `H` / `I` / `S` |
| `ROW_TEXT` | CHAR 132 | Heading text or "<=Loose Pipes" |
| `SN_TXT`, `INV_SN_TXT`, `STD_TXT` | CHAR 6 | Without leading zeros |
| `ARTNO` | `ZSD_ARTNO` | |
| `ART_TXT` | CHAR 10 | Only on the first line of an article group |
| `PARTNO` | `ZSD_PARTNO` | |
| `MATDESC` | `ZSD_MATDESC` | |
| `QTY_TXT` | CHAR 17 | |
| `GWT_TXT`, `NWT_TXT` | CHAR 17 | Only on the first line of a weight group |
| `HSN` | `ZSD_HSN` | |
| `ART_FIRST`, `ART_LAST`, `WT_FIRST`, `WT_LAST` | `XFELD` | For the borders of the "merged" cells |
| `ART_GRP`, `WT_GRP` | NUMC 4 | |

---

## 4. Lock object `EZSD_EPACK` — new (SE11)

Primary table `ZSD_EPACK_HDR`, mode `E`, lock parameters `MANDT`, `ZSD_PACKNO`.
Generates `ENQUEUE_EZSD_EPACK` / `DEQUEUE_EZSD_EPACK` (parameter
`MODE_ZSD_EPACK_HDR`).

---

## 5. Message class `ZSD_EPACK` — new (SE91)

| No. | Text | Raised by |
|---|---|---|
| `001` | `Invalid Weight: Gross Weight cannot be less than Net Weight.` **(FS 2.1, exact text)** | rules |
| `002` | `Package from Sl.No. &1: gross weight &2 kg is less than net weight &3 kg` | rules |
| `003` | `Enter at least one tax invoice or a packing list number` | report / data |
| `004` | `Invoice &1 is not assigned to any saved packing list` | data (reprint) |
| `006` | `Packing list &1 also contains invoice &2 - printed as part of it` | data (reprint) |
| `008` | `Invoice &1 does not exist` | data |
| `009` | `Invoice &1 is cancelled or is a cancellation document` | data |
| `010` | `Billing type &2 of invoice &1 is not allowed for export packing lists` | data |
| `011` | `Invoices &3 differ in &1 (&2) - one packing list needs one value` | data |
| `012` | `No authorization: &1 &2, activity &3` | data / store |
| `013` | `Notify parties exist for packing list &1 - use format with notify party` | data |
| `014` | `Packing list &1 is locked by user &2 - not saved` | store |
| `015` | `File &1 could not be read: &2` | upload |
| `016` | `File is not the packing list template (column &1 heading "&2")` | upload |
| `017` | `Row &1: mandatory field &2 is empty` | upload |
| `018` | `Row &1: "&2" in column &3 is not a valid number` | upload |
| `019` | `Row &1: "&2" in column &3 is not a valid date` | upload |
| `020` | `Row &1: &2 differs from the value in row &3` | upload |
| `021` | `Row &1: Sl.No. &2 occurs more than once` | rules |
| `022` | `Row &1: text in column &2 is longer than &3 characters` | upload |
| `027` | `Row &1: first line item has no &2 (start of a package group)` | rules |
| `028` | `Row &1: enter gross and net weight together` | rules |
| `029` | `Sl.No. &1: quantity &2 differs from STD &3 x Articles &4` | rules (warning) |
| `030` | `Sl.No. &1: part no. &2 is not billed in the packing list invoices` | data (warning) |
| `031` | `Notify party &1 is filled but notify party &2 is empty` | upload |
| `032` | `Packing list &1 saved (revision &2, &3 lines, &4 invoices)` | store |
| `036` | `The file contains no line items` | upload |
| `038` | `PDF saved: &1` | report |
| `039` | `Error &1 in &2` | output |
| `040` | `Smart Form &1 does not exist or is not active` | output |
| `041` | `Row &1: Sl.No. &2 is not in ascending order` | rules |
| `042` | `No &1 partner in KNVP for customer &2, sales area &3` | data |
| `044` | `Sl.No. &1: no HSN code in upload or material master` | data (warning) |
| `046` | `No packing list found for the selection` | data (reprint) |
| `047` | `Packing list &1 has no valid invoices or no line items` | data |
| `048` | `Sl.No. &1: no product description in upload or material master` | data (warning) |
| `049` | `&1 not found - field stays empty on the output` | data (warning) |
| `050` | `&1 rows without Sl.No. were ignored (e.g. total row)` | upload (info) |
| `052` | `Packing list &1 could not be saved - database error` | store |
| `053` | `File &1 could not be written` | output |
| `054` | `Enter a folder for the PDF download` | report |
| `055` | `Row &1: no packing list number in or above this row - row ignored` | upload |
| `056` | `&1 repeated invoice number(s) counted once` | upload (info) |
| `057` | `Row &1 is an exact duplicate of row &2 and was ignored` | upload (warning) |
| `058` | `Invoice &1 is also assigned to packing list &2 in the same file` | upload |
| `059` | `No packing list number found in the file` | upload / report |
| `060` | `Customer &1 has &2 &3 partners in KNVP - &4 is used` | data (warning) |
| `061` | `Sales order &1 has no reference document - no order no. / date` | data (warning) |
| `062` | `The processing log could not be written (ZSD_EPACK_LOG)` | store |
| `063` | `The result list could not be displayed` | output |
| `064` | `Output of packing list &1 cancelled by the user` | report (warning) |
| `065` | `Form &1 output for packing list &2` | report |
| `066` | `&1 of &2 packing lists output, &3 saved, &4 with errors` | report |
| `067` | `Validation only: &1 packing lists checked, &2 with errors` | report |

---

## 6. Repository objects — new

| Object | Type | Purpose |
|---|---|---|
| `ZSD_EPACK` | Package | Container |
| `ZIF_SD_EPACK` | Interface | Types, constants |
| `ZCX_SD_EPACK` | Exception class | Errors with message list |
| `ZCL_SD_EPACK_UPLOAD` | Class | Excel parsing, split by packing list number (**+ 20 unit tests**) |
| `ZCL_SD_EPACK_RULES` | Class | Package groups, FS weight check, totals, formatting, parsing (**+ 26 unit tests**) |
| `ZCL_SD_EPACK_DATA` | Class | Logic-sheet data retrieval for all packing lists in bulk, checks, form data |
| `ZCL_SD_EPACK_OUTPUT` | Class | Smart Form call, PDF download, result list |
| `ZCL_SD_EPACK_STORE` | Class | Save to `ZSD_EPACK_HDR/INV/DATA`, log to `ZSD_EPACK_LOG` |
| `ZSD_EPACK_PRINT` | Executable program | Selection screen, run control |
| `ZSD_EPACK_WITH_NP` | Smart Form | Format with notify parties (see `docs/ZSD_EPACK_SmartForm_Specification.md`) |
| `ZSD_EPACK_WITHOUT_NP` | Smart Form | Format without notify parties |
| `ZSD_EPACK_STYLE` | Smart Style | Shared paragraph / character formats |
| `ZSD_EPACK_CROSS` | SE78 graphic (BMON) | Diagonal line for unused notify-party sections (A26) |

### Transaction code (SE93)

| T-code | Type | Program | Text |
|---|---|---|---|
| **`ZSD_EPACK`** (FS) | Report transaction, screen 1000 | `ZSD_EPACK_PRINT` | Export Packing List |

The earlier separate upload transaction `ZSD_EPACK_UPL` / program `ZSD_EPACK_UPLOAD` has been withdrawn.
Upload, form output and saving are now one run of `ZSD_EPACK`.

---

## 7. Authorization (FS 2.3 is empty — A24)

| Object | Fields | Where |
|---|---|---|
| `S_TCODE` | `ZSD_EPACK` | Transaction start |
| **`ZSD_EPACK`** (new, SU21, class `SD`) | `ACTVT` (01 create/save, 04 print), `BUKRS` | 04 per company code before a form is output; 01 before the data is saved |
| `V_VBRK_VKO` | `VKORG`, `ACTVT` = 04 | Print — for each sales org. of the invoices |

> Placeholder until security defines the concept (CL-22 / Q14). The checks are
> isolated in `ZCL_SD_EPACK_DATA->CHECK_AUTHORITY` and
> `ZCL_SD_EPACK_STORE->SAVE_PACKING_LISTS`.

---

## 8. Customizing / parameters

| Object | Entry | Purpose |
|---|---|---|
| `TVARVC` select-option `ZSD_EPACK_FKART` | Export tax-invoice billing types (e.g. `I EQ ZEXP`) | Only these billing types are accepted. **Empty = all accepted** (A11). |
| `T001Z` party `ZCIN` | CIN per company code | FS logic B7 — must exist (CL-27) |
| Company address field `BUILDING` | IEC number | FS logic B11 (Q08) |

---

## 9. SAP objects consumed — read only (logic sheet rows)

Every table is read **once per run** for all packing lists (FOR ALL ENTRIES).

| Table | Fields | Logic sheet |
|---|---|---|
| `VBRK` | `VBELN FKART FKSTO SFAKN BUKRS VKORG VTWEG SPART KUNAG ZTERM INCO1 INCO2 BUPLA` | B3, B6, B11, B20, B26, B43, B44 |
| `VBRP` | `VBELN POSNR MATNR WERKS AUBEL` | B13, B42 |
| `KNVP` | `KUNNR VKORG VTWEG SPART PARVW PARZA KUNN2 DEFPA` — `PARVW` `AG` (= SP), `WE` (= SH) | B20, B26 |
| `KNA1` | `KUNNR ADRNR` | B20, B26 |
| `T001` | `BUKRS ADRNR` | B3, B11 |
| `ADRC` | `NAME1 STREET STR_SUPPL1 CITY1 CITY2 POST_CODE1 REGION COUNTRY BUILDING` | B3, B11, B20, B26 |
| `ADR6` | `SMTP_ADDR` | B21, B27 |
| `ADR2` | `TELNR_LONG` where `R3_USER` = 1 (telephone) / 3 (mobile) | B22, B23, B28, B29 |
| `T005U` | `BEZEI` (logic sheet says `T005S`) | B4 |
| `T005T` | `LANDX` (logic sheet says `T005`) | B5, B20, B26 |
| `J_1BBRANCH` | `GSTIN` for `BUKRS` + `BRANCH` = `VBRK-BUPLA` | B6 |
| `T001Z` | `PAVAL` for `PARTY` = `ZCIN` | B7 |
| `VBAK` | `VGBEL` for `VBELN` = `VBRP-AUBEL` | B13 |
| `VBKD` | `BSTKD BSTDK` for `VBELN` = `VGBEL`, `POSNR` 000000 | B13 |
| `TVZBT` | `VTEXT` (SPRAS E, first `ZTAGG`) | B43 |
| `MAKT` | `MAKTX` (logic sheet says `MARA`) | B54 |
| `MARC` | `STEUC` for billed material + billing plant | B42, B58 |
| `T604N` | `TEXT1` for `SPRAS` E, `LAND1` IN | B42 |
| `TVARVC`, `USR01` | billing-type filter, default printer | — |

> ⚠ **Verify before activation** in DEV: `J_1BBRANCH-GSTIN`, `T604N` key and
> text field, `TVZBT-ZTAGG`, `KNVP-DEFPA`, `ADR2-R3_USER` values. Each read is
> in `ZCL_SD_EPACK_DATA->READ_SAP_DATA` / `READ_ADDRESSES`.

---

## 10. Activation order and transports

| Order | Objects |
|---|---|
| 1 | Domains, data elements |
| 2 | Tables `ZSD_EPACK_HDR` → `ZSD_EPACK_INV` → `ZSD_EPACK_DATA` → `ZSD_EPACK_LOG`, indexes `Z01`, lock object |
| 3 | Structures, table types |
| 4 | Message class, authorization object |
| 5 | `ZIF_SD_EPACK`, `ZCX_SD_EPACK`, `ZCL_SD_EPACK_RULES`, `ZCL_SD_EPACK_UPLOAD` (run ABAP Unit) |
| 6 | `ZCL_SD_EPACK_DATA`, `ZCL_SD_EPACK_OUTPUT`, `ZCL_SD_EPACK_STORE` |
| 7 | Smart Style, SE78 graphic, both Smart Forms |
| 8 | Program `ZSD_EPACK_PRINT`, transaction `ZSD_EPACK` |

Workbench request: everything above. Customizing request: `TVARVC`
entry (or maintain it directly in each client), roles.
