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
| `ZSD_EXPREF` | CHAR | 20 | | Exporter's ref. (IEC) override | ✔ |
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

Audit fields reuse standard data elements: `ZSD_ERNAM`/`ZSD_AENAM`/`ZSD_PRNAM` → `ERNAM`/`AENAM`/`USNAM`,
`*DAT` → `DATS`, `*ZET` → `TIMS`.

> ⚠ **SAP-VERIFY** — the maximum length of `SSTRING` in a transparent table
> is 1333 characters; tables containing `STRING` fields cannot be buffered
> (not needed here).

---

## 2. Transparent tables — new

Common technical settings: delivery class **`A`**, data class `APPL1`,
size category `0`, **no buffering**, log data changes **yes**,
table maintenance **not allowed** (maintained only via `ZSD_EPACK_UPL`).

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
| `SAME_PARTY` | `XFELD` | Sold-to = ship-to → one "CONSIGNEE" block (A15) |
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

| No. | Text |
|---|---|
| `001` | `Invalid Weight: Gross Weight cannot be less than Net Weight.` **(FS 2.1, exact text)** |
| `002` | `Package from Sl.No. &1: gross weight &2 kg is less than net weight &3 kg` |
| `003` | `Enter at least one tax invoice or a packing list number` |
| `004` | `Invoice &1 is not assigned to any packing list` |
| `005` | `Selected invoices belong to more than one packing list (&1, &2)` |
| `006` | `Packing list &1 also contains invoice &2 - include it in the selection` |
| `007` | `Packing list &1 does not exist` |
| `008` | `Invoice &1 does not exist` |
| `009` | `Invoice &1 is cancelled or is a cancellation document` |
| `010` | `Billing type &2 of invoice &1 is not allowed for export packing lists` |
| `011` | `Invoices differ in &1 (&2 / &3) - one packing list needs one value` |
| `012` | `No authorization: &1 &2, activity &3` |
| `013` | `Notify parties exist for packing list &1 - use format with notify party` |
| `014` | `Packing list &1 is locked by user &2` |
| `015` | `File &1 could not be read: &2` |
| `016` | `File is not the packing list template (column &1 heading "&2")` |
| `017` | `Row &1: mandatory field &2 is empty` |
| `018` | `Row &1: "&2" in column &3 is not a valid number` |
| `019` | `Row &1: "&2" in column &3 is not a valid date` |
| `020` | `Row &1: &2 differs from the value in row &3` |
| `021` | `Row &1: Sl.No. &2 occurs more than once` |
| `022` | `Row &1: text in column &2 is longer than &3 characters` |
| `024` | `Packing list &1 already exists - use "Overwrite"` |
| `025` | `Packing list &1 does not exist - use "Create"` |
| `026` | `Invoice &1 is already assigned to packing list &2` |
| `027` | `Row &1: first line item has no &2 (start of a package group)` |
| `028` | `Row &1: enter gross and net weight together` |
| `029` | `Sl.No. &1: quantity &2 differs from STD &3 x Articles &4` |
| `030` | `Sl.No. &1: part no. &2 is not billed in the packing list invoices` |
| `031` | `Notify party &1 is filled but notify party &2 is empty` |
| `032` | `Packing list &1 saved (revision &2, &3 lines, &4 invoices)` |
| `033` | `Test run OK: &1 lines, &2 invoices - nothing saved` |
| `034` | `Packing list &1 deleted` |
| `035` | `Errors found - nothing was saved` |
| `036` | `The file contains no line items` |
| `038` | `PDF saved: &1` |
| `039` | `Error &1 in &2` |
| `040` | `Smart Form &1 does not exist or is not active` |
| `041` | `Row &1: Sl.No. &2 is not in ascending order` |
| `042` | `No ship-to party found in invoice &1` |
| `044` | `Sl.No. &1: no HSN code in upload or material master` |
| `045` | `Invoice &1 does not belong to packing list &2` |
| `046` | `No packing list found for the selected invoices` |
| `047` | `Packing list &1 has no invoices or no line items` |
| `048` | `Sl.No. &1: no product description in upload or material master` |
| `049` | `&1 not found - field stays empty on the output` |
| `050` | `&1 rows without Sl.No. were ignored (e.g. total row)` |
| `051` | `Test run: packing list &1 can be deleted` |
| `052` | `Packing list &1 could not be saved - database error` |
| `053` | `File &1 could not be written` |
| `054` | `Enter a file name for the PDF download` |

(`023`, `037` and `043` are intentionally unused.)

---

## 6. Repository objects — new

| Object | Type | Purpose |
|---|---|---|
| `ZSD_EPACK` | Package | Container |
| `ZIF_SD_EPACK` | Interface | Types, constants |
| `ZCX_SD_EPACK` | Exception class | Errors with message list |
| `ZCL_SD_EPACK_RULES` | Class | Package groups, FS weight check, totals, formatting, parsing (**+ 26 unit tests**) |
| `ZCL_SD_EPACK_DATA` | Class | Data provider for the forms, print-time validations |
| `ZCL_SD_EPACK_UPLOAD` | Class | Excel upload, validation, save, delete |
| `ZCL_SD_EPACK_OUTPUT` | Class | Smart Form call, PDF download, message popup |
| `ZSD_EPACK_PRINT` | Executable program | Print selection screen |
| `ZSD_EPACK_UPLOAD` | Executable program | Upload selection screen |
| `ZSD_EPACK_WITH_NP` | Smart Form | Format with notify parties (see `docs/ZSD_EPACK_SmartForm_Specification.md`) |
| `ZSD_EPACK_WITHOUT_NP` | Smart Form | Format without notify parties |
| `ZSD_EPACK_STYLE` | Smart Style | Shared paragraph / character formats |
| `ZSD_EPACK_CROSS` | SE78 graphic (BMON) | Diagonal line for unused notify-party sections (A26) |

### Transaction codes (SE93)

| T-code | Type | Program | Text |
|---|---|---|---|
| **`ZSD_EPACK`** (FS) | Report transaction, screen 1000 | `ZSD_EPACK_PRINT` | Export Packing List |
| `ZSD_EPACK_UPL` (A23) | Report transaction, screen 1000 | `ZSD_EPACK_UPLOAD` | Export Packing List – Excel Upload |

---

## 7. Authorization (FS 2.3 is empty — A24)

| Object | Fields | Where |
|---|---|---|
| `S_TCODE` | `ZSD_EPACK`, `ZSD_EPACK_UPL` | Transaction start |
| **`ZSD_EPACK`** (new, SU21, class `SD`) | `ACTVT` (01 create, 02 change, 04 print, 06 delete), `BUKRS` | Upload / delete / print |
| `V_VBRK_VKO` | `VKORG`, `ACTVT` = 04 | Print — for each sales org. of the invoices |

> Placeholder until security defines the concept (CL-22). The checks are
> isolated in `ZCL_SD_EPACK_DATA->CHECK_AUTHORITY`,
> `ZCL_SD_EPACK_UPLOAD->VALIDATE` and `->DELETE`.

---

## 8. Customizing / parameters

| Object | Entry | Purpose |
|---|---|---|
| `TVARVC` select-option `ZSD_EPACK_FKART` | Export tax-invoice billing types (e.g. `I EQ ZEXP`) | Only these billing types are accepted. **Empty = all accepted** (A11). |
| `T001Z` party `ZCIN` | CIN per company code | FS logic B7 — must exist (CL-27) |
| Company address field `BUILDING` | IEC number | FS logic B11, unless overridden by `ZSD_EXPREF` (A13, CL-05) |

---

## 9. SAP objects consumed — read only

| Table | Fields | FS logic sheet row |
|---|---|---|
| `VBRK` | `VBELN FKART FKSTO SFAKN BUKRS VKORG KUNAG ZTERM INCO1 INCO2 BUPLA` | B3, B6, B43, B44 |
| `VBRP` | `VBELN POSNR MATNR WERKS AUBEL` | B13, B42 |
| `VBPA` | `VBELN POSNR PARVW KUNNR ADRNR` (`AG`, `WE`) | B20, B26 (corrected) |
| `KNA1` | `KUNNR ADRNR STCD1` | B20, B26 |
| `T001` | `BUKRS ADRNR` | B3, B11 |
| `ADRC` | `NAME1 NAME2 NAME_CO STREET STR_SUPPL1/2 CITY1 CITY2 POST_CODE1 REGION COUNTRY BUILDING` | B3, B11, B20, B26 |
| `ADR2` / `ADR3` / `ADR6` | telephone (`R3_USER` 1 / 3), fax, e-mail | B21–B23, B27–B29 |
| `T005T` / `T005U` | country / region texts | B4, B5 (corrected) |
| `J_1BBRANCH` | `GSTIN` | B6 |
| `T001Z` | `PAVAL` (party `ZCIN`) | B7 |
| `VBAK` / `VBKD` | `VGBEL VGTYP` / `BSTKD BSTDK` (POSNR 000000) | B13 |
| `TVZBT` / `TINCT` | payment-term / Incoterm texts | B43, B44 |
| `MAKT` / `MARC` / `T604N` | `MAKTX` / `STEUC` / `TEXT1` | B42, B54, B58 |
| `TVARVC`, `USR01` | parameters, default printer | — |

> ⚠ **Verify before activation** in DEV: `J_1BBRANCH-GSTIN`, `T604N` key and
> text field, `TVZBT-ZTAGG`, `ADR2-R3_USER` values, `VBAK-VGTYP`. Each read is
> confined to one method of `ZCL_SD_EPACK_DATA`.

---

## 10. Activation order and transports

| Order | Objects |
|---|---|
| 1 | Domains, data elements |
| 2 | Tables `ZSD_EPACK_HDR` → `ZSD_EPACK_INV` → `ZSD_EPACK_DATA`, index `Z01`, lock object |
| 3 | Structures, table types |
| 4 | Message class, authorization object |
| 5 | `ZIF_SD_EPACK`, `ZCX_SD_EPACK`, `ZCL_SD_EPACK_RULES` (run ABAP Unit) |
| 6 | `ZCL_SD_EPACK_DATA`, `ZCL_SD_EPACK_UPLOAD`, `ZCL_SD_EPACK_OUTPUT` |
| 7 | Smart Style, SE78 graphic, both Smart Forms |
| 8 | Programs, transaction codes |

Workbench request: everything above. Customizing request: `TVARVC`
entry (or maintain it directly in each client), roles.
