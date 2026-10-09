# Packing List: Requirement Analysis and Development Approach

| Attribute | Value |
|---|---|
| Object | WRICEF 102-B – Export Consolidated Packing List |
| Client / Project | Astral Limited / UDAY |
| Module | SD |
| Development type (per FS) | Smart Form & Z table |
| Transaction (per FS) | `ZSD_EPACK` |
| Z table (per FS) | `ZSD_EPACK_DATA` |
| Source FS | `Object ID- 102(B)Export consolidated packing list.docx` (V.01, dated 21.08.2026, prepared by Satyendra Singh, **"Approved by" blank**) |
| Document purpose | Requirement analysis and development approach only. No code was written and no repository file other than this document was created or changed. |
| Analysis date | 09.10.2026 |
| Development status | Development started on 09.10.2026 at the user's request, before the open clarifications were answered. Code in `src/`; every assumption applied is listed in `docs/Development_Assumptions.md`. |

**How to read this document.** Statements are tagged as follows:

| Tag | Meaning |
|---|---|
| **[FS]** | Stated explicitly in the FS body text. |
| **[LOGIC]** | Stated in the logic sheet (`Logic_sheet1.xlsx`, also embedded in FS §2.1). |
| **[ZTAB]** | Stated in the Z-table workbook (`Z table Format1.xlsx`, also embedded in FS §2.1). |
| **[SAMPLE]** | Seen in the scanned sample outputs (`PackingList_*.pdf`). |
| **[INFERRED]** | Inferred by the analyst. Not confirmed. Needs validation. |
| **[REC]** | Analyst recommendation. Needs approval. |
| **[SAP-VERIFY]** | A technical fact about standard SAP that is based on general SAP knowledge and **has not been verified in the Astral system**. Check it in the development system before you build on it. |

---

## 1. Executive Summary

Astral Limited prepares export packing lists by hand in Excel today. The business wants these lists printed from SAP. One packing list consolidates **several tax invoices** (SAP billing documents), often from several plants and depots, and supports **one commercial invoice**. Header data such as exporter, parties, payment terms and Incoterms comes from SAP. Shipment data and **all line-item packing data** (article ranges, STD, articles, part no., quantity, gross/net weight, HSN) are uploaded from Excel into a custom table `ZSD_EPACK_DATA`. The data is printed through one of two fixed Smart Forms: with notify parties (up to five) or without. Transaction `ZSD_EPACK` triggers the output. Its input is a selection of tax invoices. One validation rule is defined: gross weight must be ≥ net weight, otherwise the output is blocked with a defined message.

**Key findings**

1. **No technical implementation exists in the repository.** It holds only the FS, two Excel annexures (both identical to the copies embedded in the FS) and two scanned sample outputs. There is no ABAP, DDIC, Smart Form, CDS or configuration artefact.
2. **The FS is a draft.** The Messages (§2.2), Security & Authorization (§2.3) and Unit Test (§3.1) sections are empty. The "Approved by" field is blank. The table of contents has a broken bookmark ("1.2 Initiating process – Error! Bookmark not defined").
3. **The Z-table design is not buildable as specified.** Its only key fields are the packing list no. and a 255-character field that is supposed to "store multiple tax invoice[s]". There is no line-item key, so one packing list cannot hold more than one line. Header and item data are mixed in one row. Some CHAR lengths (1500) probably exceed what a database table allows.
4. **The FS, logic sheet, Z-table workbook and samples contradict each other** on Exporter Ref./IEC, Description of Goods, HSN and product description, format option labelling, and declaration field naming.
5. **The samples contain layout behaviour the FS does not describe.** Examples: package groups shown as merged cells with one weight per group, HSN section headings, package-type subtotals ("Loose Pipes", "Boxes"), a ship-to tax number, a consignee-only header variant, a signature block, and crossing out of unused rows.
6. **Neither the FS nor the workbooks specify the Excel upload program**, although every manual field depends on it. There is no transaction, template rules, validations, overwrite/revision logic or authorization.
7. Several data paths in the logic sheet name the wrong standard table or partner code. For example, `LANDX` is read from `T005` instead of `T005T`, `BEZEI` from `T005S` instead of `T005U`, `MAKTX` from `MARA` instead of `MAKT`, and `PARVW = SP/SH` is used where the stored internal codes are `AG/WE` **[SAP-VERIFY]**.

**Readiness:** **Not ready.** There are **5 critical blockers** and **38 further clarifications**: 16 High, 16 Medium, 6 Low (§11, §16). The SAP-sourced header extraction and the form page skeleton can be prototyped with low rework risk. The Z table, upload program, line-item logic and final form layout should wait for answers to the critical items.

**Recommended approach (summary):** Normalise the Z-table into header, invoice-link and item tables. Build a separate upload transaction with full validation. Use one driver program for `ZSD_EPACK` with a data-provider class that does all SELECTs. The two Smart Forms should only format the data they receive and share one style and one DDIC interface. Weight and consistency checks run before any form is called (§7, §8).

---

## 2. Repository Files Reviewed

Repository: `Nisha-lohar2/Packing_List`, branch `claude/loving-clarke-f6w9xy` (single commit `fbb3423 "Add files via upload"`, identical to `main` at analysis time).

| # | File | Type | How it was reviewed | Content summary | Relevance |
|---|---|---|---|---|---|
| F1 | `Object ID- 102(B)Export consolidated packing list.docx` | Functional Specification (Word) | Full text converted and read. Package unzipped and inspected: two embedded OLE Excel objects, header/footer XML, media. | Cover data, Overview & Scope, process flow (9 steps), 28-row field-source table, Z-table name, embedded logic sheet & Z-table sheet, GW ≥ NW validation, empty Messages/Authorization/Test sections. | **Primary requirement source** |
| F1a | `word/embeddings/Microsoft_Excel_Worksheet.xlsx` (inside F1) | Embedded Excel | Dumped cell by cell. Compared with F2. | **Identical content to F2** (sheets `Input`, `Logic`). | Same as F2 |
| F1b | `word/embeddings/Microsoft_Excel_Worksheet1.xlsx` (inside F1) | Embedded Excel | Dumped cell by cell. Compared with F3. | **Identical content to F3** (sheets `Format`, `Technical detail`, `Sheet3`). | Same as F3 |
| F1c | `word/media/image1.emf`, `image2.emf` | Icons for the embedded objects | Identified via relationships | Display icons only. No extra content. | None |
| F1d | `word/media/image3–6` | Header/footer images | Viewed | Astral logo and header graphics of the FS document template. | None (not an output layout) |
| F2 | `Logic_sheet1.xlsx` | Excel | Every cell read | `Input` sheet: T-code `ZSD_EPACK`, "Tax Invoice From/To", "Multiple selection of invoice data". `Logic` sheet: field-by-field source logic for header, line items and footer. | **Primary technical-logic source** |
| F3 | `Z table Format1.xlsx` | Excel | Every cell, merged range and total read. Totals and rules recalculated. | `Format`: upload template with 77 sample line items (matches sample PDF 90000081) and totals. `Technical detail`: proposed Z-field names, lengths, types, key fields. `Sheet3`: copy of FS field-source table **with differences** (see §6.3). | **Primary data-model source** |
| F4 | `PackingList_90000014 Tarmac UAE_BySea.pdf` | Scanned PDF, 4 pages, image only (no text layer) | Rendered and read visually, all 4 pages | Current manual output. Format **without** notify party, header shows "CONSIGNEE" only, 2 HSN groups, 90 lines, package-type subtotals, totals, declaration, signature/stamp. | Reference layout / expected output |
| F5 | `PackingList_90000081 Nabil Yemen_BySea - Rev.1.pdf` | Scanned PDF, 4 pages, image only | Rendered and read visually (pages 1, 2 and 4 in detail; page 3 is a continuation) | Current manual output. Format **with 5 notify parties**, separate Sold-to/Bill-to (Buyer) and Ship-to (Consignee) with tax number, 77 lines, merged package groups, totals. "Rev.1" in the file name. | Reference layout / expected output |

**Files referenced but not available**

| Ref | Referenced item | Where referenced | Status |
|---|---|---|---|
| M1 | "Approved Packing List format" | FS §1 Out of scope | No formal layout specification is in the repository. Only the scanned samples F4/F5 exist. It is not confirmed that these are the approved format. |
| M2 | FS §1.2 "Initiating process" | FS TOC | Section is missing ("Error! Bookmark not defined"). |
| M3 | Message list | FS §2.2 | Table is empty. |
| M4 | Security & Authorization | FS §2.3 | Section is empty. |
| M5 | Unit test conditions / results | FS §3.1 | Table is empty. |
| M6 | Commercial invoice object (probably WRICEF 102-A, by naming) | FS process steps 5 and 8. Z-table fields `ZSD_ADV_AMT`, `ZSD_CMMT`, `ZSD_INV_DECL`. | Not in the repository. Its relationship to this object is unknown **[INFERRED]**. |
| M7 | Excel upload program specification | FS §1 "Requirements", §2.1 | Not specified anywhere. |
| M8 | Existing ABAP / Smart Form / DDIC objects | — | None present. No existing implementation can be compared. |

**Other repository in this session:** `Nisha-lohar2/KPMG_Stock_Validation` was also checked out. A content search for `packing`/`EPACK` returned no matches. Its contents (Stock Validation, Order Loading, Loan Details, MDG assistant) are unrelated to this object, so it was not analysed further.

---

## 3. Business Requirement Overview

### 3.1 Business problem [FS §1]
- Users prepare the packing list manually in Excel. It contains exporter, buyer, notify party, reference numbers, material information, quantity, net weight, gross weight and other details.
- The consolidated export packing list must include the **line-level serial numbers of the corresponding commercial invoice** ("Sr.No. INV"). Customs officials use them to reconcile the packing list with the commercial invoice.
- Manual preparation is error-prone and not traceable to SAP documents **[INFERRED]**.

### 3.2 Objective [FS §1 "Requirements"]
- Generate the packing list **directly from SAP**.
- Data that is **not in SAP** is maintained by users through an **Excel upload into a Z table**.
- Data that **varies from order to order** is to be "updated manually by users in the designated fields within relevant transactional documents, such as Sales Orders and Invoices". *Note: the FS does not say which fields these are, and the logic sheet takes no data from such fields (see CL-24).*

### 3.3 Scope [FS §1]
| In scope | Out of scope |
|---|---|
| A standardised Smart Form, the same for all customers and countries | Customer-specific or country-specific dynamic fields |
| Two fixed formats: with and without notify party | Dynamic Smart Form (explicitly excluded: "There will be NO dynamic smart form provided") |
| Z table for non-SAP data, filled by Excel upload | Any format other than the approved packing list format |
| T-code `ZSD_EPACK` to generate the output | — |

### 3.4 Assumptions stated in the FS
- "Data filed correctly in excel sheet and must be validated by user." (The user is responsible for the correctness of uploaded data.)
- "Correct master data maintained in SAP system like HSN code."

---

## 4. End-to-End Business Process

### 4.1 Process flow as documented [FS §1.1]

| Step | Activity | SAP object (standard) | Source |
|---|---|---|---|
| 1 | Sales quotation | Quotation (VBAK, quotation doc. category) | [FS] |
| 2 | Sales order created with reference to the quotation | Sales order | [FS] |
| 3 | Delivery created with reference to the sales order (multiple deliveries possible; a plant-level packing slip combines delivery lines) | Outbound delivery | [FS §1] |
| 4 | Tax invoice created with reference to the delivery | Billing document (VBRK/VBRP) | [FS] |
| 5 | Commercial invoice prepared with reference to **multiple tax invoices** (one commercial invoice may cover several sales orders) | **Unknown whether this is an SAP document** | [FS] / CL-16 |
| 6 | Export packing list prepared (consolidated across all plants and depots) | **This development** | [FS] |
| 7 | "Output will be generated with reference to Tax invoice" | `ZSD_EPACK` | [FS] |
| 8 | Commercial invoice sent to bank and customer | Outside this object | [FS] |
| 9 | Two packing list formats: (a) with notify party, (b) without notify party | Two Smart Forms | [FS] |

### 4.2 Target process for this object (FS + logic sheet + inferred steps)

```
 [User prepares Excel per approved template]                       (manual, outside SAP)
          │
          ▼
 [Upload Excel → Z table ZSD_EPACK_DATA]   ◄── upload program/transaction NOT specified (CL-11)
          │  key: Packing List No. + SAP invoice number(s)
          ▼
 [T-code ZSD_EPACK] selection: Tax invoice(s) From/To + multiple selection   [LOGIC Input sheet]
          │
          ├─ Determine packing list no. for the selected invoices        (rule NOT specified – CL-02)
          ├─ Read SAP data: company code, branch, partners, quotation ref., payment terms, Incoterms, HSN
          ├─ Read Z-table header + line items
          ├─ Validate: Gross Wt ≥ Net Wt  → error "Invalid Weight: Gross Weight cannot be less than Net Weight."
          ├─ Choose form: with notify party / without notify party      (selection rule contradictory – CL-03)
          ▼
 [Smart Form output] → print / preview / (PDF?)                         (channel NOT specified – CL-19)
          │
          ▼
 [Signed & stamped hard copy → customs / bank docs]                     [SAMPLE: stamp & signature]
```

### 4.3 Triggering, inputs and outputs

| Aspect | Documented | Source | Gap |
|---|---|---|---|
| Trigger | User runs T-code `ZSD_EPACK` | [FS], [LOGIC Input] | FS step 7 ("generated with reference to Tax invoice") could also mean an output type on the billing document. CL-19. |
| Input | Tax invoice numbers: From/To range plus multiple selection | [FS] "Invoice number Multiple invoice selection", [LOGIC Input] | No other parameters (company code, sales org, PL no., format choice, print options). Mandatory/optional not defined. CL-02, CL-03. |
| Processing | Field-level logic in `Logic_sheet1.xlsx` | [LOGIC] | Multi-invoice consolidation rules are missing. CL-04. |
| Validation | GW ≥ NW, else error and no output | [FS §2.1] | Level (line/group/total) not defined. CL-10. |
| Output | Smart Form (two variants) | [FS] | Layout spec, output device, PDF, copies, reprint/revision not defined. CL-19, CL-20, CL-37. |

---

## 5. Detailed Functional Requirements

### 5.1 Requirement inventory

IDs `FR-xx` are explicit requirements from the FS and its embedded annexures. IDs `IR-xx` are requirements implied by the sample outputs or by necessity, and they are **not** in the FS text.

| ID | Requirement | Source | Type |
|---|---|---|---|
| FR-01 | Generate the export packing list from SAP via T-code `ZSD_EPACK` | FS §1.1; LOGIC Input!B4:C4 | Explicit |
| FR-02 | Input = multiple tax invoices (From/To + multiple selection) | FS §1.1; LOGIC Input!C7:F8 | Explicit |
| FR-03 | Two Smart Form formats: with and without notify party. No dynamic form. | FS §1.1 step 9 | Explicit |
| FR-04 | Notify party format holds up to 5 notify parties. Unused sections are crossed out. | FS §1.1 | Explicit |
| FR-05 | Format is standardised across customers and countries | FS §1 Scope | Explicit |
| FR-06 | Non-SAP data is uploaded from Excel into Z table `ZSD_EPACK_DATA` | FS §1, §2.1 | Explicit |
| FR-07 | Exporter name and address from company code (incl. region text, country text) | FS table #1; LOGIC B3–B5 | Explicit |
| FR-08 | GST No. from company code / business place | FS #6; LOGIC B6 | Explicit |
| FR-09 | CIN No. from company code | FS #7; LOGIC B7 | Explicit |
| FR-10 | Exporter reference no. (IEC) | FS #2; LOGIC A11/B11; ZTAB `ZSD_EXPREF` | Explicit but **conflicting** |
| FR-11 | Packing list no. and date (manual) | FS #4, #5; LOGIC B9–B10 | Explicit |
| FR-12 | SAP invoice number(s) (manual, in the Z table) | FS #3; ZTAB `VBRK_VBELN` | Explicit |
| FR-13 | Order No. & Date = customer reference and date of the quotation | FS #8; LOGIC B13 | Explicit |
| FR-14 | Buyer's reference no. & date (manual) | FS #9; LOGIC B14 | Explicit |
| FR-15 | Country of origin of goods (manual) | FS #10; LOGIC B16 | Explicit |
| FR-16 | Country of final destination = ship-to country | FS #11; LOGIC B17 | Explicit |
| FR-17 | Sold-to/Bill-to (Buyer) block: name, address, country, email, telephone, mobile | FS #13; LOGIC B20–B23 | Explicit |
| FR-18 | Ship-to (Consignee) block: name, address, country, email, telephone, mobile | FS #12; LOGIC B26–B29 | Explicit |
| FR-19 | Notify parties 1–5 (manual) | FS #14; LOGIC B31 | Explicit |
| FR-20 | Pre-carriage, place of receipt by pre-carrier, vessel/flight no., port of loading, port of discharge, place of delivery (manual) | FS #15–20; LOGIC B33–B38 | Explicit |
| FR-21 | Marks & Nos., Container No., No. of packages/bundles (manual) | FS #21, #22, #25; LOGIC B39–B41 | Explicit |
| FR-22 | Description of goods, with common HSN codes and descriptions clubbed across all billing documents | FS #26; LOGIC B42 | Explicit but **conflicting** |
| FR-23 | Payment terms text | FS #23; LOGIC B43 | Explicit |
| FR-24 | Terms of shipment = Incoterms 1 and 2 | FS #24; LOGIC B44 | Explicit |
| FR-25 | Line items: S.No, Sr.No. INV, Article No., STD, Articles, Total Articles, Part No., Product Description, Quantity (Pcs), Gross Wt. KGS, Net Wt. KGS, H.S. Code | FS #27; LOGIC B47–B58 | Explicit |
| FR-26 | Product description: from Z table **or** material master | LOGIC B54; ZTAB `Format`!G1, `Technical detail`!K5 | Explicit but **undecided** |
| FR-27 | Line-item HSN: from Z table **or** `MARC-STEUC` ("feasibility to be checked") | FS #27; LOGIC B58; ZTAB `Technical detail`!O5 | Explicit but **undecided** |
| FR-28 | Declaration in the footer (manual) | FS #28; LOGIC B61 | Explicit |
| FR-29 | Validation at output time: Gross Wt ≥ Net Wt, else error "Invalid Weight: Gross Weight cannot be less than Net Weight." and no output | FS §2.1 Additional Details | Explicit |
| FR-30 | The user validates the uploaded data; correct master data (e.g. HSN) is an assumption | FS §1.1 | Explicit (assumption) |
| FR-31 | Order-varying data maintained in designated fields of SO/invoice | FS §1 Requirements | Explicit but **undefined** |
| IR-01 | Full header block repeated on every page | SAMPLE F4/F5 all pages | Implied |
| IR-02 | "Page n of m" footer | SAMPLE | Implied |
| IR-03 | Grand total line: total articles/packages, total quantity, total gross weight, total net weight | SAMPLE F4 p4, F5 p4; ZTAB `Format`!E83:J83 | Implied |
| IR-04 | Item table split into sections, with an HSN heading row (e.g. "PVC PIPES (LEAD FREE) – HS CODE : 3917 2390") | SAMPLE | Implied |
| IR-05 | Package-type subtotal rows (e.g. "2342 <=Loose Pipes", "1073 <=Boxes") and a total text ("Total 2342 Loose Pipes + 1073 Boxes = 3415 Packages") | SAMPLE F4 p1, p4 | Implied |
| IR-06 | Package grouping: Article No., Articles, Gross Wt. and Net Wt. printed once per group of lines (merged cells) | SAMPLE F5; ZTAB `Format` merged ranges (e.g. C2:C3, E2:E3, I2:I3, J2:J3) | Implied |
| IR-07 | Signature block "Signature & Date / For Astral Limited / Authorized Signatory" | SAMPLE | Implied |
| IR-08 | Unused rows on the last page are crossed out diagonally | SAMPLE F4 p4, F5 p4 | Implied |
| IR-09 | Ship-to tax number printed ("TAX NUMBER : 0187153") | SAMPLE F5 | Implied |
| IR-10 | Contact person ("K/A:- MR. …") and FAX printed in party blocks | SAMPLE F4/F5 | Implied |
| IR-11 | Without-NP variant shows a single "CONSIGNEE" block instead of separate Sold-to and Ship-to | SAMPLE F4 | Implied (conflicts with FR-17/18) |
| IR-12 | Quantities shown as integers and weights with 3 decimals | SAMPLE | Implied |
| IR-13 | Revision of an issued packing list ("Rev.1") | SAMPLE F5 file name | Implied |
| IR-14 | Excel upload program with validation (a prerequisite for FR-06) | Necessity | Implied |
| IR-15 | Authorization, messages, test cases | FS §2.2, §2.3, §3.1 (empty) | Required but empty |

### 5.2 Header field requirements (consolidated view)

| # | Output label (sample) | Source per FS table | Source per `Sheet3` (ZTAB) | Logic sheet | Z field (Technical detail) | Consistent? |
|---|---|---|---|---|---|---|
| 1 | Exporter | Company code | Company code | `VBRK→T001→ADRC` | — | Yes |
| 2 | Exporter's Ref. (IEC) | **Company code** | **Manual (Excel)** | **`ADRC-BUILDING`** of company address | `ZSD_EXPREF` CHAR20 marked "**Not Req** / This data will be fetched from SAP System". Format col O "Exporter ReF" still present. | **No – CL-05** |
| 3 | SAP Invoice number | Manual (Excel) | Manual | — | `VBRK_VBELN` CHAR255 (key) | Partially – CL-01 |
| 4 | Packing List No. | Manual | Manual | `ZSD_PACKNO` | `ZSD_PACKNO` CHAR10 (key) | Yes |
| 5 | Date | Manual | Manual | `ZSD_PACKDT` | `ZSD_PACKDT` DATS | Yes |
| 6 | GST No. | Company code | Company code | `J_1BBRANCH-GSTIN` via `VBRK-BUPLA` | — | Yes (multi-branch – CL-30) |
| 7 | CIN No. | Company code | Company code | `T001Z` PARTY = `ZCIN` | — | Yes (config – CL-27) |
| 8 | Order No. & Date | Quotation (customer ref. & date) | Quotation | `VBAK-VGBEL → VBKD-BSTKD/BSTDK` | — | Yes (multi-SO – CL-15) |
| 9 | Buyer's Reference No. & Date | Manual | Manual | `ZSD_PARTY_REF`, `ZSD_REF_DT` | CHAR200, DATS | Yes |
| 10 | Country of Origin | Manual | Manual | `ZSD_CNTY_ORGN` | CHAR20 | Yes |
| 11 | Country of final destination | Ship-to country | Ship-to country | "Ship to party country – standard" | — | Yes |
| 12/13 | Ship-to / Sold-to | Customer master | Customer master | `KNVP` (PARVW SH/SP) → `KNA1` → `ADRC/ADR6/ADR2` | — | Yes (technical issues – CL-13) |
| 14 | Notify party 1–5 | Manual | Manual | `ZSD_NP1..5` | CHAR200 ×5 | Yes (length – CL-12) |
| 15–20 | Shipping fields | Manual | Manual | `ZSD_PREC`, `ZSD_REC_PC`, `ZSD_VSl_FLT`, `ZSD_POL`, `ZSD_POD`, `ZSD_PLD` | CHAR30 each | Yes |
| 21 | Marks & Nos. | Manual | Manual | `ZSD_MARK` | CHAR30 | Yes |
| 22 | Container No. | Manual | Manual | `ZSD_CONT` | CHAR100 | Yes (length – CL-12) |
| 23 | Payment term | Payment terms | Payment terms | `VBRK-ZTERM → TVZBT-VTEXT` | — | Yes |
| 24 | Terms of shipment | Incoterms | Incoterms | `VBRK-INCO1/INCO2` | — | Yes |
| 25 | No. of Pkgs. | Manual | Manual | `ZSD_NOPACK` | CHAR30 | Yes (multi-line – CL-12) |
| 26 | Description of Goods | **HSN code & description from material master** | **Manual (Excel)** | **`MARC-STEUC → T604N-TEXT1`, clubbed** | `ZSD_HSN_DESC` CHAR100 "Same HSN having different description" | **No – CL-06** |
| 27 | Line items | Manual (HSN from material master if feasible) | Manual | `ZSD_*` fields | see §9.3 | Partially – CL-07 |
| 28 | Declaration | Manual | Manual | **`ZSD_DECL`** | **`ZSD_PL_DECL`** CHAR1500 | Naming conflict – CL-12 |

---

## 6. Existing System and Technical Analysis

### 6.1 Existing implementation
**None in the repository.** There are no ABAP programs, classes, function modules, Smart Forms, DDIC objects, CDS views, enhancements or configuration exports. It is unknown whether `ZSD_EPACK`, `ZSD_EPACK_DATA` or the `ZSD_*` data elements already exist in the Astral system (CL-31). It is also unknown whether a commercial-invoice form (102-A?) already exists and could share objects (CL-16, CL-17).

### 6.2 Target system landscape
The repository does not state the SAP release (ECC 6.0 vs S/4HANA, and which version). The FS uses classic tables (`VBRK`, `KNVP`, `KNA1`, `T001Z`, `J_1BBRANCH`). Most of these still exist in S/4HANA **[SAP-VERIFY]**, but the release affects the Excel upload API, the clean-core options, and Business Partner versus customer master address handling. See CL-31.

### 6.3 Conflicts between repository files

| Ref | Topic | File A says | File B says | Resolution needed |
|---|---|---|---|---|
| X-01 | Exporter Ref./IEC | FS table #2: "Details in company code" | `Sheet3`!D4: "Manual update through excel sheet". ZTAB `ZSD_EXPREF`: "Not Req / fetched from SAP". LOGIC B11: `ADRC-BUILDING`. | CL-05 |
| X-02 | Description of Goods | FS #26: "HSN Code & description from material master" | `Sheet3`!D28: "Manual update through excel sheet". ZTAB `ZSD_HSN_DESC` (manual) with note "Same HSN having different description". LOGIC B42: `T604N-TEXT1`, clubbed. | CL-06 |
| X-03 | Format option labels | FS: "a. with notifier party, b. without notifier party" | FS, next sentence: "If notifier party details are maintained … user have to select **option b – Packing list with notifier party**" | CL-03 |
| X-04 | Declaration field | LOGIC B61: `ZSD_DECL` | ZTAB: `ZSD_PL_DECL` (plus `ZSD_INV_DECL` for invoice) | CL-12 |
| X-05 | Total Articles | LOGIC B52: from `ZSD_TART`. ZTAB Technical detail: `ZSD_TART` | ZTAB `Format`: no "Total Articles" column. The total (E83 = 4074) is a sum of Articles. | CL-40 |
| X-06 | SAP Invoice Number column | ZTAB `Format` col L "SAP Invoice Number" | ZTAB `Technical detail`: `VBRK_VBELN` CHAR255 "Store multiple tax invoice" as a key field | CL-01 |
| X-07 | Product description | ZTAB `Format`!G1: "(Fetch from material master)". `Technical detail`!K5: "Fetch from material master" | LOGIC B54: "`ZSD_MATDESC` OR `MAKTX`". Field `ZSD_MATDESC` exists in the Z table. | CL-07 |
| X-08 | HSN step reference | LOGIC B58: "… ( From STEP no 53)" | There is no step 53 for MATNR. Material comes from `VBRP` (B42) or Part No. (row 53 = Part No.). | CL-07 |
| X-09 | Party blocks | FS/LOGIC: separate Sold-to and Ship-to blocks | SAMPLE F4: a single "CONSIGNEE" block | CL-13 |
| X-10 | Order-varying data in SO/invoice fields | FS §1 Requirements | LOGIC: no field read from SO/invoice texts or custom fields | CL-24 |
| X-11 | Field list | FS table rows 1–28 | ZTAB adds `ZSD_ADV_AMT`, `ZSD_CMMT`, `ZSD_INV_DECL` (commercial-invoice fields) that do not appear in the packing list | CL-17 |
| X-12 | Description of goods source in FS row 26 vs row 27 | Row 26: from material master | Row 27: "We will check feasibility to fetch HSN… Otherwise fetch from Excel" | CL-06/CL-07 |

### 6.4 Technical review of the logic sheet

| Ref | Logic sheet step | Observation | Suggested correction | Status |
|---|---|---|---|---|
| T-01 | B4 "Pass BLAND in T005S get BEZEI" | The region **text** `BEZEI` is in `T005U` (key SPRAS, LAND1, BLAND). `T005S` holds regions without texts. | Read `T005U` with LAND1 + BLAND + language | [SAP-VERIFY] |
| T-02 | B5, B20, B26 "Pass LAND1 in T005 get LANDX" | The country name `LANDX` is in `T005T` (key SPRAS, LAND1), not `T005`. | Read `T005T` | [SAP-VERIFY] |
| T-03 | B54 "Pass MATNR in MARA get MAKTX" | `MAKTX` is in `MAKT` (key MATNR, SPRAS). | Read `MAKT`. Language to be confirmed. | [SAP-VERIFY] |
| T-04 | B20/B26 "KUNNR=KUNAG & PARVW=SP/SH in KNVP" | `SP`/`SH` are the English external codes. The stored internal codes are `AG` (sold-to) and `WE` (ship-to). The `KNVP` key also needs sales org, distribution channel and division. Customer-master partners are **defaults**. The partners actually used on the invoice are in `VBPA`. | Use document partners from `VBPA` (billing doc, else delivery/sales order) | [SAP-VERIFY] + CL-13 |
| T-05 | B13 "Pass VBELN (AUBEL) in VBAK get VGBEL" | `AUBEL` (sales document) is a **`VBRP`** (item) field, not `VBRK`. Several invoices → several SOs → possibly several quotations. Quotation header business data is in `VBKD` with `POSNR = 000000`. | `VBRP-AUBEL` → `VBAK-VGBEL` (with doc. category check) → `VBKD` (POSNR 000000) | [SAP-VERIFY] + CL-15 |
| T-06 | B42 "Pass MATNR in MARC get STEUC" | `MARC` is plant-specific. The plant must come from `VBRP-WERKS`. The same material in different plants may have different `STEUC`. | Key `MARC` by `VBRP-MATNR` + `VBRP-WERKS` and handle conflicts | [SAP-VERIFY] + CL-07 |
| T-07 | B42 "T604N with SPRAS=EN, LAND1=IN" | Hard-coded language and country. `T604N` key/text field names need verification. | Parameterise (constants/TVARVC) | [SAP-VERIFY] |
| T-08 | B43 "ZTERM in TVZBT get VTEXT" | The `TVZBT` key includes `ZTAGG` (day limit). The text may need `ZTAGG` = initial or the first entry. | Read with SPRAS + ZTERM (+ ZTAGG) | [SAP-VERIFY] |
| T-09 | B22/B23, B28/B29 `ADR2-R3_USER` = 1 / 3 | Values: '1' = default fixed-line, '3' = default mobile (domain values to be verified). Several numbers may exist. The sample also prints FAX (`ADR3`), which is not in the logic. | Confirm domain values. Take the default/first number. Add FAX if required (CL-14). | [SAP-VERIFY] + CL-25 |
| T-10 | B21/B27 `ADR6-SMTP_ADDR` | Several emails possible. No rule given. | Use the default flag or the first `CONSNUMBER` | CL-25 |
| T-11 | B6 GSTIN via `J_1BBRANCH` | Requires `VBRK-BUPLA` (business place) to be filled. Invoices from several branches/states may have different GSTINs. | Validate a single business place, or define a rule | [SAP-VERIFY] + CL-30 |
| T-12 | B7 `T001Z` PARTY = `ZCIN` | Custom parameter type. It must be configured. | Verify in OBY6 / `T001Z` | CL-27 |
| T-13 | B11 IEC = `ADRC-BUILDING` | Misuses an address field. It only works if business maintains the IEC there. | Confirm (CL-05) | CL-05 |
| T-14 | B3 Exporter from `VBRK-BUKRS` | `VBRK-BUKRS` exists. With multiple invoices, all must have the same company code. | Validate the same BUKRS | CL-04 |
| T-15 | B9–B61 "Pass ZSD_PACKNO" | Every Z-table read uses the packing list no., but the selection screen takes invoices. The derivation PL no. ← invoice is not defined. | Define the linkage | CL-02 |

### 6.5 Technical review of the Z-table design (`Z table Format1.xlsx` → `Technical detail`)

| Ref | Observation | Impact |
|---|---|---|
| Z-01 | The only key fields are `ZSD_PACKNO` (CHAR10) and `VBRK_VBELN` (CHAR255, "Store multiple tax invoice"). There is **no client (`MANDT`)** and **no line-item key** (e.g. `ZSD_SN`). With this key, only one row per PL no. + invoice string is possible, so the 77 to 90 lines in the samples cannot be stored. | Table cannot hold line items. **Blocker.** |
| Z-02 | Concatenating several invoice numbers into one CHAR255 field cannot be indexed, joined or validated. A search by invoice needs a `LIKE` match, and the ~23-invoice limit is implicit. | Selection by invoice (FR-02) cannot be done reliably |
| Z-03 | Header fields (PL date, parties, ports, declaration, …) and item fields are in one flat structure. The sample `Format` sheet fills only the item columns (A–K) and leaves header columns L–AK empty. It is unclear whether header values are entered once or on every row. | Redundant data, possible inconsistency, unclear upload rules |
| Z-04 | `ZSD_PL_DECL` and `ZSD_INV_DECL` are CHAR1500. The DDIC limit for CHAR fields in transparent tables is believed to be 1333 characters **[SAP-VERIFY]**. A string/long-text design may be needed. | Activation error or truncation |
| Z-05 | `ZSD_QTY`, `ZSD_GWT`, `ZSD_NWT` are QUAN 13. QUAN needs a unit reference field (not defined) and a number of decimals (not defined; samples use 3 decimals for weights). `ZSD_ADV_AMT` is CURR and needs a currency key field (not defined). | Activation errors / rounding issues |
| Z-06 | `ZSD_ART` (Articles) is NUMC6. Rows inside a merged group have **no** article value (blank). NUMC cannot tell blank from 0. There is no package-group identifier. | Grouping/merged-cell logic (IR-06) cannot be represented |
| Z-07 | Field name `ZSD_VSl_FLT` has a lowercase "l". ABAP DDIC names are upper case. | Cosmetic. Rename to `ZSD_VSL_FLT`. |
| Z-08 | `ZSD_EXPREF` is marked "Not Req" but is still listed. | CL-05 |
| Z-09 | No audit fields (created by/on, changed by/on, upload ID, revision). | Traceability and the "Rev.1" requirement (IR-13) |
| Z-10 | No package-type field (Loose Pipes / Boxes / Bundles) and no HSN section heading text. | IR-04 / IR-05 cannot be produced from the data |
| Z-11 | `ZSD_SN` and `ZSD_INV_SN` are NUMC6. The sample has Sl.No. 1..90 and Sr.No. INV 1..27. | OK, if upload rules are defined |
| Z-12 | `ZSD_ARTNO` is CHAR15. Sample values: "1 TO 60", "3781 TO 3925", "STACKER", "726020", "1069". | OK. A range text, not numeric. |

### 6.6 Facts derived from the sample data (verified by recalculation)

| Fact | Evidence | Status |
|---|---|---|
| Quantity (Pcs) = STD × Articles (of the line's package group) | All 77 lines in `Z table Format1.xlsx` `Format` satisfy it (recalculated). Tarmac sample lines checked visually, e.g. 15 × 24 = 360, 46 × 12 = 552. | [INFERRED] rule, holds for the samples. CL-41. |
| Grand totals = sums of the line/group values | `Format`!E83 = 4074 = ΣArticles, H83 = 12146 = ΣQty, I83 = 116323.8 = ΣGW, J83 = 115386.78 = ΣNW (values are constants in the sheet, not formulas) | [INFERRED]. CL-40. |
| The Z-table sample is the data of PDF 90000081 (Nabil) | Row values, merged groups and totals match `PackingList_90000081…pdf` | Fact |
| Article number ranges may have gaps | `Format` C75 "3926", then C76 "3951 TO 4023", C77 "4024", C78 "4051 TO 4124". Total articles 4074 ≠ last article no. 4124. | Fact. Total ≠ max article no. |
| Weights are given per package group, not per line | Merged I/J ranges (e.g. I2:I3, J2:J3) | Fact. Impacts FR-29 (CL-10). |
| In every sample group GW ≥ NW | Recalculated: no group with GW < NW | Fact |

---

## 7. Proposed Technical Architecture

> All object names below are **proposals** following the FS `ZSD_` naming. Final names depend on Astral naming conventions (CL-31).

### 7.1 Component view

```
                ┌──────────────────────────────────────────────────────────┐
                │   Excel template (approved, versioned)                    │
                └───────────────┬──────────────────────────────────────────┘
                                │ upload (T-code proposed: ZSD_EPACK_UPL)
                ┌───────────────▼──────────────────────────────────────────┐
                │ ZSD_EPACK_UPLOAD (report) + ZCL_SD_EPACK_UPLOAD          │
                │  parse → validate → simulate (ALV log) → save            │
                └───────────────┬──────────────────────────────────────────┘
                                │ writes (locked via EZSD_EPACK)
      ┌─────────────────────────▼────────────────────────────────────────────┐
      │ Z tables (recommended normalised design – deviates from FS, CL-01)   │
      │  ZSD_EPACK_HDR  (PL header: manual header fields, audit, revision)   │
      │  ZSD_EPACK_INV  (PL no. ↔ billing document, 1:n)                     │
      │  ZSD_EPACK_ITM  (PL no. + Sl.No.: line items, package group)         │
      │  (or ZSD_EPACK_DATA kept as the item table name if FS name required) │
      └─────────────────────────┬────────────────────────────────────────────┘
                                │ read
 T-code ZSD_EPACK ┌─────────────▼──────────────────────────────────────────────┐
 ────────────────►│ ZSD_EPACK_PRINT (driver report)                            │
                  │  ZCL_SD_EPACK_DATA   – all SELECTs (Z + SAP), consolidation │
                  │  ZCL_SD_EPACK_CHECK  – validations (weights, consistency)  │
                  │  ZCX_SD_EPACK        – exception class, msg class ZSD_EPACK│
                  └─────────────┬──────────────────────────────────────────────┘
                                │ interface structures (DDIC)
           ┌────────────────────┴────────────────────┐
           ▼                                         ▼
 ZSD_EPACK_WITH_NP (Smart Form)          ZSD_EPACK_WITHOUT_NP (Smart Form)
           └──────────── shared Smart Style ZSD_EPACK_STYLE ──┘
                                │
                     Print / Preview / PDF (CL-19)
```

### 7.2 Why this architecture

| Decision | Rationale | Dependency / limitation |
|---|---|---|
| Separate upload program from print program | The FS mentions only `ZSD_EPACK` for generation. Upload and print have different users, authorizations and frequency. Uploading must be possible (and re-runnable) before printing. | Needs FS addendum (CL-11) |
| Normalised Z tables (header / invoice link / items) | Fixes Z-01 to Z-03 and Z-06. Allows selection by invoice (FR-02), many line items, one copy of the header, revision tracking. | **Deviation from FS** `ZSD_EPACK_DATA` single table. Needs functional approval (CL-01). If the FS design must stay, at minimum add `MANDT` and `ZSD_SN` to the key and replace CHAR255 multi-invoice storage. |
| Data-provider class outside the form | Keeps SELECTs out of the Smart Forms. Both forms use one data set. Logic can be unit-tested with ABAP Unit. | Form interface must be DDIC structures/table types |
| Two Smart Forms, one style, one interface | Explicit FS requirement (FR-03, "NO dynamic smart form"). The notify-party block changes header height (samples: ~22 vs ~29 item rows per page), which is why two forms are needed. | Duplicate layout maintenance. Mitigate with a shared style and shared text modules. |
| Validations before calling the form | FR-29 requires no output on error. Checking before `SSF_FUNCTION_MODULE_NAME`/form call guarantees nothing is spooled. | Validation level must be confirmed (CL-10) |
| Lock object `EZSD_EPACK` | Prevents concurrent upload/print of the same PL | Small effort |

### 7.3 Processing flow of `ZSD_EPACK` (proposed)

1. Selection screen: billing documents (select-option, From/To + multiple), optional PL no. (proposed, CL-02), output options (preview/print/PDF – CL-19), optionally format override (CL-03).
2. Authority checks (§8 step 9).
3. Read billing headers (`VBRK`) and items (`VBRP`) for the selection. Reject cancelled/cancellation documents and wrong billing types (CL-29).
4. Determine the PL number(s) via the invoice link table. Error if no PL, several PLs, or the PL also contains invoices that were not selected (rule per CL-02).
5. Consistency checks across invoices: company code, sold-to, ship-to, payment terms, Incoterms, business place/GSTIN (CL-04, CL-30).
6. Read SAP header data (company address, GSTIN, CIN, IEC, partners, quotation references, payment term text, Incoterms).
7. Read Z header and items. Derive product description and HSN per the agreed rule (CL-07). Build package groups, HSN sections and package-type subtotals (CL-08, CL-09).
8. Business validations: GW ≥ NW (CL-10), optionally Qty = STD × Articles (CL-41), optionally reconciliation against `VBRP` (CL-18).
9. On error: show messages (application log/ALV) and stop. No form is called.
10. Determine the form: notify party maintained → with-NP form, else without-NP form (CL-03).
11. Call the form (preview/print/PDF). Optionally log output (who/when/revision – CL-37).

---

## 8. Step-by-Step Development Approach

| Step | Activity | Objects (proposed) | Pre-requisite clarifications | Why / notes |
|---|---|---|---|---|
| 0 | Close the critical clarifications and get FS sign-off, including the approved layout | — | CL-01, CL-02, CL-04, CL-08, CL-11 | Avoids rework on the data model and form |
| 1 | DDIC: domains/data elements for the manual fields. Tables `ZSD_EPACK_HDR`, `ZSD_EPACK_INV`, `ZSD_EPACK_ITM` (or a corrected `ZSD_EPACK_DATA`). Lock object `EZSD_EPACK`. Table types/structures for the form interface. | DDIC | CL-01, CL-12, CL-28 | Field lengths must handle the samples (notify party > 200 chars, 5 containers) |
| 2 | Message class `ZSD_EPACK` with all errors/warnings (incl. FS text "Invalid Weight: Gross Weight cannot be less than Net Weight.") | SE91 | CL-21 | FS §2.2 is empty. A list is proposed in §10.3. |
| 3 | Excel upload: template definition, parser, validations, simulation mode, save, re-upload/revision rules, upload log | `ZSD_EPACK_UPLOAD`, `ZCL_SD_EPACK_UPLOAD`, T-code proposed `ZSD_EPACK_UPL` | CL-11, CL-08, CL-37 | Parser API depends on the release (ECC: e.g. `TEXT_CONVERT_XLS_TO_SAP`; S/4HANA: e.g. `XCO_CP_XLSX` or `CL_FDT_XL_SPREADSHEET`) **[SAP-VERIFY]** |
| 4 | Data provider: SAP reads per the corrected logic (§9.2), consolidation across invoices, quotation refs, HSN/description clubbing | `ZCL_SD_EPACK_DATA` | CL-04, CL-05, CL-06, CL-13, CL-15, CL-30 | Can be **prototyped first** (low rework) |
| 5 | Validation class | `ZCL_SD_EPACK_CHECK` | CL-10, CL-18, CL-29, CL-41 | Unit-testable |
| 6 | Driver report + T-code `ZSD_EPACK`, selection screen, output control, error display | `ZSD_EPACK_PRINT` | CL-02, CL-03, CL-19 | |
| 7 | Smart Style + Smart Forms: page FIRST/NEXT with header windows on every page, main window table with HSN section rows, package-group rendering, subtotals, totals on last page, declaration, signature block, page n of m, NP crossing-out | `ZSD_EPACK_STYLE`, `ZSD_EPACK_WITH_NP`, `ZSD_EPACK_WITHOUT_NP`, text modules | CL-09, CL-14, CL-20 | See feasibility notes §8.1 |
| 8 | PDF download / preview / print handling | in driver | CL-19 | Standard OTF→PDF conversion **[SAP-VERIFY]** |
| 9 | Authorizations: `S_TCODE` for both T-codes. Billing-document authority (e.g. `V_VBRK_VKO` sales org, `V_VBRK_FKA` billing type **[SAP-VERIFY]**). Custom object (e.g. `ZSD_EPACK`, activity 01/02/03/06) for upload/change/delete/print. | SU21 / role | CL-22 | FS §2.3 is empty |
| 10 | Unit tests (ABAP Unit for check/derivation logic), integration and UAT per §14 | Test classes | CL-23 | |
| 11 | Transport, documentation, user guide for the Excel template | — | — | |

### 8.1 Smart Form feasibility notes (require validation)

| Requirement | Feasibility assessment | Status |
|---|---|---|
| Header repeated on every page (IR-01) | Secondary windows on pages FIRST/NEXT, main window for the item table. Standard technique. | Feasible |
| Page n of m (IR-02) | System fields `SFSY-PAGE` / `SFSY-FORMPAGES` | Feasible **[SAP-VERIFY]** |
| Merged cells for package groups (IR-06) | Smart Form tables cannot row-span. Workaround: print group values on the group's first line only and use line types without inner horizontal borders. Values cannot be vertically centred as in the Excel sample. | Feasible with deviation. **Needs business acceptance.** |
| Do not split a package group across pages | Page protection on a loop/folder per group | Feasible **[SAP-VERIFY]** |
| "Crossed out" unused notify-party sections (FR-04) | Smart Forms has no diagonal-line drawing element. Options: a fixed graphic (SE78) with a diagonal line per fixed-size NP window, or a printed text such as "NOT APPLICABLE"/"XXXX". The NP windows are fixed in size, so a graphic is realistic. | **Feasibility requires validation** |
| Diagonal cross-out of unused rows on the last page (IR-08) | The remaining space varies, so a fixed graphic cannot stretch to fit. Alternatives: a "*** END OF PACKING LIST ***" line, or fill to page end with a fixed graphic. | **Not feasible exactly as in the sample. Needs business decision.** |
| Totals only on last page | Table footer/calculation on the last loop pass | Feasible |
| Long texts (declaration 1500 chars, notify party > 200) | Pass as string/text lines (`TLINE`) or include long text | Feasible. Depends on the DDIC design (Z-04). |

---

## 9. Data Sources and Field Mapping

### 9.1 Selection / input

| Field | Proposed technical | Mandatory | Source |
|---|---|---|---|
| Tax invoice(s) | `SELECT-OPTIONS s_vbeln FOR vbrk-vbeln` (range + multiple) | Yes [INFERRED] | LOGIC Input |
| Packing list no. | `PARAMETERS p_packno` (alternative entry) | Proposed | [REC] CL-02 |
| Output mode | Preview / Print / PDF | Proposed | [REC] CL-19 |
| Format | Automatic (default) / override | Proposed | [REC] CL-03 |

### 9.2 Header – source-to-target mapping (corrected where marked)

| Target (form) | Source per logic sheet | Proposed source (corrected) | Notes / open item |
|---|---|---|---|
| Exporter name/address | `VBRK-BUKRS → T001-ADRNR → ADRC` NAME1, STREET, STR_SUPPL1, CITY1, POST_CODE1, REGION, COUNTRY | Same | Company of the first/all invoices (CL-04) |
| Exporter region text | `T005S-BEZEI` | `T005U-BEZEI` (SPRAS, LAND1, BLAND) | T-01 |
| Exporter country text | `T005-LANDX` | `T005T-LANDX` | T-02 |
| GST No. | `VBRK-BUPLA` + `BUKRS → J_1BBRANCH-GSTIN` | Same | CL-30 |
| CIN No. | `T001Z-PAVAL` (PARTY = `ZCIN`) | Same | CL-27 |
| Exporter's Ref. (IEC) | `ADRC-BUILDING` (company address) | **Undecided** | CL-05 |
| Packing list no. / date | `ZSD_PACKNO` / `ZSD_PACKDT` | Z header | CL-16 |
| Order No. & Date | `VBAK(AUBEL)-VGBEL → VBKD-BSTKD/BSTDK` | `VBRP-AUBEL → VBAK-VGBEL → VBKD (POSNR 000000)`. Distinct values concatenated. | T-05, CL-15 |
| Buyer's Ref. No. & Date | `ZSD_PARTY_REF`, `ZSD_REF_DT` | Z header | CL-34 |
| Country of origin | `ZSD_CNTY_ORGN` | Z header | |
| Country of final destination | Ship-to country | Ship-to `ADRC-COUNTRY → T005T-LANDX` | |
| Sold-to block | `KNVP` (KUNAG, PARVW SP) → `KNA1-ADRNR → ADRC`, `ADR6`, `ADR2` | `VBPA` partner `AG` of billing doc → `ADRC`/`ADR6`/`ADR2` (+`ADR3` FAX?) | T-04, CL-13, CL-14 |
| Ship-to block | Same with PARVW SH | `VBPA` partner `WE` (billing or delivery) | T-04, CL-13 |
| Ship-to tax number | — (not in logic) | e.g. `KNA1-STCD1` / BP tax number **[SAP-VERIFY]** | CL-14 |
| Notify parties 1–5 | `ZSD_NP1..5` | Z header (long text) | CL-12 |
| Pre-carriage … Place of delivery | `ZSD_PREC`, `ZSD_REC_PC`, `ZSD_VSl_FLT`, `ZSD_POL`, `ZSD_POD`, `ZSD_PLD` | Z header | |
| Marks & Nos. / Container No. / No. of Pkgs | `ZSD_MARK` / `ZSD_CONT` / `ZSD_NOPACK` | Z header (multi-line) | CL-12 |
| Description of goods | `VBRP-MATNR → MARC-STEUC → T604N-TEXT1`, clubbed by HSN | **Undecided** (SAP vs `ZSD_HSN_DESC`) | CL-06 |
| Payment term | `VBRK-ZTERM → TVZBT-VTEXT` | Same (+ZTAGG) | T-08 |
| Terms of shipment | `VBRK-INCO1`, `INCO2` | Same. In S/4HANA also check `INCO2_L`. **[SAP-VERIFY]** | Sample: "EX-WORKS" (text of INCO1 from `TINCT`?) – CL-20 |
| Declaration | `ZSD_DECL` | `ZSD_PL_DECL` (long text) | X-04 |

### 9.3 Line items – source-to-target mapping

| Form column | Z field (Technical detail) | Type/len | Alternative SAP source | Open item |
|---|---|---|---|---|
| Sl.No. | `ZSD_SN` | NUMC6 | — (or generated) | CL-43 |
| Sr.No. INV | `ZSD_INV_SN` | NUMC6 | — | CL-16 |
| Article No. | `ZSD_ARTNO` | CHAR15 | — | group-level (CL-08) |
| STD | `ZSD_STD` | NUMC6 | — | Meaning to be confirmed ("standard pack qty"?) [INFERRED] |
| Articles | `ZSD_ART` | NUMC6 | — | group-level (CL-08) |
| Total Articles | `ZSD_TART` | NUMC6 | Σ `ZSD_ART` | CL-40 |
| Part No. | `ZSD_PARTNO` | CHAR40 | = material no.? [INFERRED] | CL-07 |
| Product Description | `ZSD_MATDESC` | CHAR40 | `MAKT-MAKTX` | CL-07 |
| Quantity Pcs | `ZSD_QTY` | QUAN13 | = STD × Articles? | CL-41, CL-28 |
| Gross Wt. KGS | `ZSD_GWT` | QUAN13 | — | group-level, 3 dec. |
| Net Wt. KGS | `ZSD_NWT` | QUAN13 | — | group-level, 3 dec. |
| H.S Code | `ZSD_HSN` | CHAR17 | `MARC-STEUC` | CL-07 |
| (missing) Package group ID | — | — | — | CL-08 |
| (missing) Package type (Loose Pipes / Boxes / Bundles) | — | — | — | CL-09 |
| (missing) HSN section heading text | `ZSD_HSN_DESC` (header) | CHAR100 | `T604N-TEXT1` | CL-06, CL-09 |

### 9.4 Z-table fields not used by the packing list
`ZSD_ADV_AMT` (CURR15, "Advance amount received"), `ZSD_CMMT` (CHAR500, "Commercial terms of invoice") and `ZSD_INV_DECL` (CHAR1500, "Declaration of Invoice") look like **commercial-invoice** fields. The packing list logic and samples do not use them (CL-17).

---

## 10. Business Rules and Validations

### 10.1 Explicit rules (from the FS/annexures)

| ID | Rule | Source |
|---|---|---|
| BR-01 | Gross Weight ≥ Net Weight. Otherwise error "Invalid Weight: Gross Weight cannot be less than Net Weight." and **no form output**. | FS §2.1 |
| BR-02 | If notify-party data exists → with-notify-party format, else without. Up to 5 NPs. Unused NP sections crossed out. | FS §1.1 |
| BR-03 | Common HSN codes and descriptions are "clubbed" across all billing documents (Description of Goods) | LOGIC B42 |
| BR-04 | Telephone = `ADR2` number with R3_USER = 1. Mobile = R3_USER = 3. | LOGIC B22/B23 |
| BR-05 | Country of final destination = ship-to country | FS #11 |
| BR-06 | Order No. & Date = quotation customer reference & date | FS #8, LOGIC B13 |
| BR-07 | Product description and HSN from Z table **or** material master (decision pending) | LOGIC B54, B58 |
| BR-08 | The user is responsible for the correctness of uploaded data | FS §1.1 |

### 10.2 Recommended additional validations (require approval)

| ID | Validation | Where | Rationale | Clarification |
|---|---|---|---|---|
| BV-01 | Billing documents exist, are not cancelled/cancellation documents, and have an allowed billing type (export tax invoice) | Upload + print | Data integrity | CL-29 |
| BV-02 | All selected invoices: same company code, sold-to, ship-to, payment terms, Incoterms (or a defined rule) | Print | Single header data (FR-07 to FR-24) | CL-04 |
| BV-03 | Selected invoices belong to exactly one PL. The PL contains no unselected invoices (or warning). | Print | Linkage | CL-02 |
| BV-04 | Mandatory Z fields filled (PL no., date, invoice link, at least one line item) | Upload | | CL-11 |
| BV-05 | GW ≥ NW per package group (and total) | Upload (warning) + print (error) | FR-29 | CL-10 |
| BV-06 | Qty = STD × Articles | Upload | Holds for all sample lines | CL-41 |
| BV-07 | Part No. is a material on the selected invoices. Σ qty per material = billed quantity. | Print | Customs reconciliation purpose | CL-18 |
| BV-08 | Duplicate Sl.No. within a PL is rejected. PL no. is unique. | Upload | Data integrity | CL-11 |
| BV-09 | Field lengths / line counts fit the form windows (e.g. NP text) | Upload | Prevent truncation | CL-12 |
| BV-10 | HSN in Z table = `MARC-STEUC` (warning) | Upload/print | Consistency | CL-07 |
| BV-11 | Notify parties filled consecutively (no NP3 without NP2) | Upload | Layout | CL-03 |

### 10.3 Proposed message list (FS §2.2 is empty – to be confirmed)

| No. | Type | Proposed text | Source |
|---|---|---|---|
| 001 | E | Invalid Weight: Gross Weight cannot be less than Net Weight. | **FS §2.1 (exact text)** |
| 002 | E | No packing list data found for the selected invoices | Proposed |
| 003 | E | Selected invoices belong to more than one packing list | Proposed |
| 004 | E | Invoice &1 is cancelled or not an export tax invoice | Proposed |
| 005 | E | Selected invoices have different &1 (company code/sold-to/ship-to/…) | Proposed |
| 006 | E | No authorization for &1 | Proposed |
| 007 | E/W | Quantity ≠ STD × Articles in line &1 | Proposed |
| 008 | E | Packing list &1 is locked by user &2 | Proposed |
| 009 | E | Upload file: mandatory field &1 missing in row &2 | Proposed |
| 010 | W | HSN &1 in upload differs from material master &2 | Proposed |

---

## 11. Assumptions and Clarifications Required

Priority: **C** = Critical, **H** = High, **M** = Medium, **L** = Low. Blocking: **B** = blocks development, **P** = affects part of the development, **T** = can be resolved during testing.

| ID | FS section / file | Description of the gap / ambiguity | Assumption under consideration | Why important | Exact clarification required | Owner | Impact if assumption wrong | Prio | Block |
|---|---|---|---|---|---|---|---|---|---|
| CL-01 | FS §2.1; `Z table Format1.xlsx` › Technical detail | Z-table key = `ZSD_PACKNO` + `VBRK_VBELN` CHAR255 "store multiple tax invoice". No MANDT, no line-item key. Header and item data mixed. | Normalised design: header, invoice-link and item tables (§7) | The table cannot store line items as specified. Selection by invoice is impossible. | (1) Do you approve replacing `ZSD_EPACK_DATA` with header/invoice/item tables, or must the single table be kept? (2) If kept: may `ZSD_SN` be added to the key and invoices stored one per row? | Functional consultant (SD) + technical lead | Full redesign of DDIC, upload and data provider | C | B |
| CL-02 | FS §1.1; LOGIC Input; LOGIC B9–B61 | Selection is by tax invoices, but all Z data is keyed by PL no. No rule for PL derivation. | PL derived via the invoice link. All invoices of a PL must be selected. A PL no. parameter is offered as an alternative. | Determines the selection screen, linkage and validations | (1) Can one invoice belong to more than one PL (e.g. partial shipments, revisions)? (2) What should happen if the user selects only some invoices of a PL, or invoices from two PLs? (3) May the user also select by PL no.? (4) Which selection fields are mandatory? | Functional consultant / business | Wrong data printed or selection unusable | C | B |
| CL-03 | FS §1.1 step 9 | Option labelling contradicts itself ("a = with", then "select option b – with notifier party"). It is also unclear whether the user selects the format or the system determines it. | The system picks the with-NP form when any `ZSD_NP1..5` is filled. No user choice. | Selection screen and form determination | Confirm: (1) correct labels a/b; (2) automatic determination vs user choice; (3) if the user chooses "without" while NPs exist, error or allowed? | Functional consultant | Wrong format printed | H | P |
| CL-04 | FS §1, §1.1; LOGIC B3–B44 | One PL consolidates several invoices (several plants/depots/SOs), but all header logic reads "VBRK" as if there were one invoice. No rule if invoices differ in company code, sold-to, ship-to, payment terms, Incoterms. | All invoices must share company code, sold-to, ship-to, payment terms and Incoterms, otherwise error | Header correctness for customs | For each header field: must all invoices be identical (error if not), or which invoice wins (first/lowest number), or should values be concatenated? | Functional consultant / business | Wrong legal document content | C | B |
| CL-05 | FS table #2; Sheet3 row 2; LOGIC B11; ZTAB `ZSD_EXPREF` | Exporter Ref./IEC: company code vs manual vs `ADRC-BUILDING` vs "Not Req" | IEC read from SAP company address (`ADRC-BUILDING`) | Legal reference on the document | Confirm the single source of the IEC no. Is it maintained in the company-code address field BUILDING in all clients? Should `ZSD_EXPREF` be removed from the template? | Functional consultant / FI master data | Wrong/missing IEC | H | P |
| CL-06 | FS #26, #27; Sheet3 row 26; LOGIC B42; ZTAB `ZSD_HSN_DESC` | Description of Goods: material master/T604N vs manual. The sample shows business text ("PP LOW NOISE PIPES (PP PIPES) – HS CODE : 3917 2390"), not obviously a T604N text. "Same HSN having different description" note. | Manual `ZSD_HSN_DESC` per HSN, uploaded | Visible header content and section headings | (1) Final source: `T604N` text, uploaded text, or uploaded with T604N fallback? (2) If uploaded, one text per HSN or per PL? (3) Exact output format "<text> – HS CODE : <HSN>"? | Functional consultant / business | Wrong description of goods on export docs | H | P |
| CL-07 | FS #27; LOGIC B54, B58; ZTAB Format G1, Technical detail K5/O5 | Line HSN and product description: Z table "OR" material master. No decision and no precedence rule. Part No. = MATNR not confirmed. `MARC` needs a plant. | Part No. = material no. Description from `MAKT` (EN). HSN from `MARC-STEUC` of the invoice plant, Z value as fallback. | Data source and validation | (1) Is Part No. always the SAP material number? (2) Final source and precedence for description and HSN? (3) Language of the description? (4) What if the same material has different HSN in different plants? | Functional consultant | Wrong item descriptions/HSN | H | P |
| CL-08 | ZTAB Format merged ranges; SAMPLE F5 | Package groups (merged Article No./Articles/GW/NW across 2+ lines) are not modelled. Excel merged cells lose the value in the second row. | Add a package-group ID column. Group values are entered on the first line of the group only. | Upload design, weight validation and form rendering | (1) Confirm the concept "one package/bundle range contains several materials". (2) Approve a "Package Group No." column in the template. (3) Should groups never split across pages? (4) Is the merged-cell look mandatory, or is printing on the first line acceptable? | Functional consultant / business | Wrong totals, failing validations, layout rework | C | B |
| CL-09 | SAMPLE F4 p1/p4, F5 p1 | Item table has HSN section heading rows and package-type subtotals ("<=Loose Pipes", "<=Boxes", "Total … = 3415 Packages"). No source in the FS or Z table. | Add a package-type field per line and generate headings per HSN | Layout completeness vs the "approved format" | (1) Are HSN section headings and package-type subtotals required? (2) Where do package types come from (new upload column)? (3) Exact total-line text? | Functional consultant / business | Output rejected in UAT | H | P |
| CL-10 | FS §2.1 | GW ≥ NW: level (line / package group / grand total) not stated. Lines without weights (group members) exist. Equality allowed per FS. | Check per package group and grand total. Lines without weights are skipped. Also check at upload (warning). | Correct error triggering | (1) At which level? (2) Also at upload? (3) Is GW = NW acceptable (FS says "equal to or greater")? (4) Is zero weight allowed? | Functional consultant | False errors, or invalid documents printed | H | P |
| CL-11 | FS §1, §2.1 | **No specification for the Excel upload**: T-code, template version, header vs line rows, mandatory fields, simulate mode, overwrite/delete/re-upload, behaviour after print, who uploads | Separate T-code (`ZSD_EPACK_UPL` proposed) with test run, full validation, overwrite only with confirmation, upload log | Every manual field depends on it | Provide an upload specification: template layout (how header fields are entered), mandatory fields, duplicate handling, overwrite/versioning rules, deletion, display/maintenance transaction, users. | Functional consultant | Development cannot start for upload. Data quality risk. | C | B |
| CL-12 | ZTAB Technical detail; SAMPLE | Field lengths vs sample content: NP texts in F5 are about 230+ chars with 5 lines (CHAR200). 5 container nos. in F5 (CHAR100). No. of Pkgs is multi-line (CHAR30). Declaration CHAR1500 (possible DDIC limit). Naming `ZSD_DECL` vs `ZSD_PL_DECL`. | Use long-text/string fields with line breaks for NP, containers, packages, declaration | Truncation, activation errors | (1) Approve longer/multi-line fields. (2) How are line breaks entered in Excel (Alt+Enter)? (3) Maximum lines per NP block? (4) Confirm the declaration field name. | Functional + technical lead | Truncated legal text | H | P |
| CL-13 | LOGIC B20, B26; SAMPLE F4 | Partners read from customer master `KNVP` with `SP/SH` (internal codes AG/WE) instead of document partners. F4 shows only "CONSIGNEE". | Use document partners (`VBPA`) of the billing document | A ship-to that differs from the master default would print wrongly | (1) Document partners or customer-master defaults? (2) In the without-NP format, are Sold-to and Ship-to both printed, or one "Consignee" block (as in F4)? When? | Functional consultant | Wrong consignee on export docs | H | P |
| CL-14 | SAMPLE F4/F5 | Ship-to TAX NUMBER, contact person "K/A", FAX printed in samples. Not in FS/logic. | Not in scope unless confirmed | Layout completeness | Are tax number, contact person and fax required? From which fields (e.g. `KNA1-STCD1`, contact-person partner, `ADR3`)? | Functional consultant | UAT rejection | M | P |
| CL-15 | FS #8; LOGIC B13 | Several invoices → several SOs → several quotations. The sample prints "PROFORMA INVOICE NO.: AL/TBMT/2025-26/MAY/01R1 & 03R4-04R4 DATE: 22.05 & 21.05". The concatenation format is not defined. | Distinct quotation `BSTKD`/`BSTDK` concatenated with "&" | Header content and length | (1) Confirm the source is the quotation's customer reference (`VBKD-BSTKD/BSTDK`). (2) Concatenation/format rule for several references. (3) What if an SO has no quotation? | Functional consultant | Wrong/missing reference | H | P |
| CL-16 | FS §1.1 steps 5, 8; ZTAB `ZSD_INV_SN` | Is the commercial invoice an SAP document? Is the PL no. (samples 90000014, 90000081) a commercial-invoice/SAP number or a free manual number? Can Sr.No. INV be validated? | PL no. is manual, unique, CHAR10. Sr.No. INV is not validated. | Linkage, numbering, validation | (1) How is the commercial invoice created (SAP doc type, Excel, 102-A form)? (2) Who assigns the PL no. – SAP number range or user? (3) Is Sr.No. INV validated against commercial-invoice items? | Functional consultant / business | Missing reconciliation, duplicate numbers | H | P |
| CL-17 | ZTAB `ZSD_ADV_AMT`, `ZSD_CMMT`, `ZSD_INV_DECL` | Commercial-invoice fields in the packing list Z table | They belong to the commercial-invoice object and are excluded here | Scope and shared table design | Should these fields be part of this table (shared with 102-A) or removed? | Functional consultant | Scope creep / missing data | M | P |
| CL-18 | FS §1 (purpose: customs reconciliation) | No rule ties uploaded lines to SAP invoice items | No reconciliation (FS assumes the user validates) | Integrity of a customs document | Should the system check that each Part No. exists on the selected invoices, and that Σ quantity per material = billed quantity? Error or warning? | Functional consultant / business | Inconsistent packing list vs invoice | H | P |
| CL-19 | FS §1.1 step 7 | Output channel: T-code only, or also an output type on the billing document (NACE)? Preview, print, PDF download, email? Copies? Printer? | T-code report with preview/print/PDF download | Driver design | Confirm channel(s), PDF download/email needs, number of copies/originals, default printer. | Functional consultant | Rework of driver/output config | H | P |
| CL-20 | FS §1 Out of scope (approved format); SAMPLE | The approved layout is not in the repository. Scanned samples only. Font, logo, column widths, rows per page, decimals, date format (27-05-2025 vs 08-10-2025 vs 18.08.2025), Incoterm text ("EX-WORKS") not specified. | Samples F4/F5 are the approved layouts | Form build and UAT acceptance | Provide the approved layout (editable Excel/PDF) for both formats with dimensions, fonts, logo, date/number formats, and the Incoterm description source. Confirm the samples are the approved format. | Business / functional | Layout rework | H | P |
| CL-21 | FS §2.2 | Message table empty | Use the list in §10.3 | Error handling | Approve or extend the message list. | Functional consultant | Minor | M | T |
| CL-22 | FS §2.3 | Security & Authorization empty | Restrict by T-code, sales org/billing type, custom object for upload | Compliance | Which roles may upload, change, delete and print? Is restriction by sales org/company code/plant required? | Functional consultant / security team | Unauthorized changes to legal docs | H | P |
| CL-23 | FS §3.1 | Unit test table empty. No test data. | Test plan in §14 | Test readiness | Provide test invoices, Excel files and expected outputs for each scenario. | Functional consultant | Delayed UAT | M | T |
| CL-24 | FS §1 Requirements | "Data that varies per order to be updated manually in designated fields within SO and invoices" – no fields named | No SO/invoice fields are used | Possibly missing data source | Which data, which fields (header texts? custom fields?) in SO/invoice? Or is this sentence obsolete? | Functional consultant | Missing data | M | P |
| CL-25 | LOGIC B21–B23, B27–B29 | Several emails/phones possible. R3_USER values. Sample also shows FAX. | Default (flagged) entry, else first | Correct contact data | Confirm the selection rule for several emails/phones and the R3_USER meaning in your system. | Functional consultant | Minor content error | L | T |
| CL-26 | LOGIC B4, B5, B54, B42 | Wrong tables for texts (`T005S`, `T005`, `MARA`). Language EN and country IN hard-coded. | Corrected to `T005U`, `T005T`, `MAKT`. Language EN. | Technical correctness | Confirm the output language (always EN?) and approve the corrections. | Technical lead | Empty fields at runtime | M | T |
| CL-27 | LOGIC B7 | CIN via `T001Z` PARTY `ZCIN` | Configured in all relevant company codes | Header data | Confirm `ZCIN` is configured and maintained. | FI consultant | Empty CIN | M | T |
| CL-28 | ZTAB Technical detail | QUAN/CURR fields without unit/currency reference. Decimals not defined. Sample has 3 decimals for weights and integer quantity. | QTY 0 dec. (or 3), weights 3 dec., unit KG fixed | DDIC activation and rounding | Confirm decimals and units (always KG/PCS?). | Functional + technical | Rounding differences | M | P |
| CL-29 | FS §1.1 step 4 | Which billing types are "tax invoices"? Handling of cancelled invoices, sales orgs, domestic vs export | Allowed billing types maintained in TVARVC. Cancelled invoices rejected. | Selection correctness | List the allowed billing types/sales orgs. Confirm cancelled invoices are rejected. | Functional consultant | Wrong documents accepted | H | P |
| CL-30 | FS §1 ("combining all plants and depots"); LOGIC B6 | Invoices from several plants/depots may have different business places → different GSTINs | All invoices must have the same business place, else error | GST No. on the document | Can one PL contain invoices from different GST registrations/states? If so, which GSTIN is printed? | Functional consultant / tax | Wrong GSTIN on export doc | H | P |
| CL-31 | — | SAP release (ECC/S/4HANA, version), naming conventions, existing objects (`ZSD_EPACK`, `ZSD_EPACK_DATA`) | S/4HANA on-premise, no existing objects | API choice, clean core | Confirm the release, naming convention and whether any objects already exist. | Technical lead / Basis | Wrong technical choices | M | P |
| CL-32 | — | Volumes (invoices per PL, lines per PL, PLs per day) not stated. Samples: 77–90 lines. | Up to about 50 invoices and 1,000 lines per PL | Performance design | Confirm the expected maximum volumes. | Business | Performance issues | L | T |
| CL-33 | FS #10 | Country of origin is free text CHAR20 ("INDIA") | Free text | Consistency | Free text, or a country key with the text from `T005T`? | Functional consultant | Minor | L | T |
| CL-34 | FS #9; SAMPLE | Buyer's reference date is a separate DATS field, but the samples embed dates in text ("E-MAIL DATED: WHATSAPP DATED 14TH AUGUST, 2025"). F4 has no date. | Date optional, printed only if filled | Layout | Is the date mandatory? Print format? | Functional consultant | Minor | L | T |
| CL-35 | FS cover/TOC | FS V.01 not approved. Broken TOC entry 1.2. Empty sections. | — | Governance | Provide a signed-off FS version that includes the answers to this list. | Functional consultant / PM | Uncontrolled scope | M | P |
| CL-36 | — | Locking, audit trail and data retention for Z data | Lock object + created/changed fields | Traceability | Is a change log required? How long is Z data retained? Can it be deleted after printing? | Functional consultant | Audit gaps | M | T |
| CL-37 | SAMPLE F5 "Rev.1" | Revisions of an issued PL: re-upload, revision number on the output? | Re-upload overwrites. A revision counter is kept but not printed. | Data model and output | How are revisions handled? Should "Rev. n" be printed? Should previous versions be kept? | Functional consultant / business | Lost history or misleading output | M | P |
| CL-38 | FS §1.1 | Crossing out unused NP sections and rows: Smart Forms cannot draw dynamic diagonal lines (§8.1) | Fixed graphic for NP windows. "End of list" text instead of diagonal lines for rows. | Layout feasibility | Is the alternative acceptable for (a) NP sections and (b) unused item rows on the last page? | Business / functional | Layout rejected | M | P |
| CL-39 | ZTAB `Sheet3` vs FS table | Sheet3 (an older copy?) differs from the FS table on rows 2 and 26 | The FS table prevails, subject to CL-05/06 | Document control | Confirm which version is current and remove the outdated one. | Functional consultant | Confusion | L | T |
| CL-40 | LOGIC B52; ZTAB | Total Articles: uploaded (`ZSD_TART`) or calculated (Σ Articles)? Not in the Format template. | Calculated | Totals correctness | Confirm calculation vs upload. Is it printed per line or only as a total? | Functional consultant | Wrong totals | M | T |
| CL-41 | ZTAB sample; SAMPLE | Qty = STD × Articles holds in all samples. Meaning of "STD" is not documented. | STD = pieces per article/package. The rule is validated. | Validation | Confirm the meaning of STD and whether Qty = STD × Articles should be enforced (error/warning). | Functional consultant | Data errors not caught | M | T |
| CL-42 | ZTAB sample | Article ranges have gaps (3927–3950, 4025–4050). No rule for overlap/gap checks. | No gap/overlap check | Data quality | Should overlapping or gapped article ranges be flagged? | Functional consultant | Minor | L | T |
| CL-43 | ZTAB `ZSD_SN` | Sl.No.: uploaded or generated at print? Sorting order of lines? | Uploaded, printed in Sl.No. order | Output order | Confirm the source and sort order (by Sl.No., or by HSN section then Sl.No.?). | Functional consultant | Line order differs | M | T |

**Summary of clarifications:** 43 items in total: **Critical 5** (CL-01, 02, 04, 08, 11), **High 16** (CL-03, 05, 06, 07, 09, 10, 12, 13, 15, 16, 18, 19, 20, 22, 29, 30), **Medium 16** (CL-14, 17, 21, 23, 24, 26, 27, 28, 31, 35, 36, 37, 38, 40, 41, 43), **Low 6** (CL-25, 32, 33, 34, 39, 42). That makes 38 non-critical clarifications. CL-14 would move to High if the business confirms that the extra sample content (tax number, contact person, fax) is mandatory.

---

## 12. Development Blockers and Dependencies

### 12.1 Critical blockers

| # | Blocker | Clarification | What is blocked |
|---|---|---|---|
| B-1 | The Z-table design cannot store line items or support selection by invoice | CL-01 | DDIC, upload, data provider, form interface |
| B-2 | No rule linking selected invoices to a packing list | CL-02 | Selection screen, driver logic |
| B-3 | No consolidation rule for header data across several invoices | CL-04 | Header extraction, validations |
| B-4 | Package-group concept (merged cells, group weights) not modelled | CL-08 | Upload template, GW/NW validation, item rendering |
| B-5 | Excel upload program not specified | CL-11 | Upload development, all manual data |

### 12.2 Dependencies

| Type | Dependency | Owner |
|---|---|---|
| Configuration | `T001Z` party `ZCIN`. `J_1BBRANCH` GSTIN. Business place on billing documents. Billing types for export tax invoices. | FI/SD config |
| Master data | Company address (incl. IEC in BUILDING?). Customer addresses/emails/phones. `MARC-STEUC` (HSN) per plant. `MAKT` texts. `T604N` texts. | Master data team |
| Process | Quotation → SO → delivery → invoice document flow maintained (`VBAK-VGBEL`) | SD |
| Other objects | Commercial invoice object (102-A?), shared Z data | Functional |
| Technical | SAP release, Excel parser availability, Smart Forms (SE78 graphics for crossing out), printer/output device | Technical lead / Basis |
| Documents | Approved layout, signed-off FS, upload specification | Functional |
| Test | Test system with export invoices, sample Excel files, expected outputs | Functional / testing |

---

## 13. Requirement-to-Implementation Traceability Matrix

| Req ID | Requirement | FS source / section | Related repo file / object | Proposed implementation | Dependencies | Clarification status | Test scenario |
|---|---|---|---|---|---|---|---|
| FR-01 | Generate PL via `ZSD_EPACK` | FS §1.1 | F2 Input sheet | T-code `ZSD_EPACK` → report `ZSD_EPACK_PRINT` | CL-19 | Open (channel) | TS-01 |
| FR-02 | Multiple tax invoice selection | FS §1.1; F2 Input | F2 | `SELECT-OPTIONS` on `VBRK-VBELN` + link table lookup | CL-02, CL-29 | **Blocked (CL-02)** | TS-02, TS-03, TS-04 |
| FR-03 | Two formats, no dynamic form | FS §1.1 step 9 | F4 (without), F5 (with) | `ZSD_EPACK_WITH_NP`, `ZSD_EPACK_WITHOUT_NP`, shared style | CL-03, CL-20 | Open | TS-10, TS-11 |
| FR-04 | Up to 5 NPs, unused crossed out | FS §1.1 | F5 | Fixed NP windows, graphic/text for unused | CL-38, CL-12 | Open (feasibility) | TS-12 |
| FR-05 | Standard format for all customers | FS §1 | — | No customer-specific logic | — | Confirmed | TS-26 |
| FR-06 | Excel upload into `ZSD_EPACK_DATA` | FS §1, §2.1 | F3 | `ZSD_EPACK_UPLOAD` + Z tables | CL-01, CL-11 | **Blocked** | TS-30 to TS-38 |
| FR-07 | Exporter address | FS #1 | F2 B3–B5 | `T001`/`ADRC`/`T005U`/`T005T` | CL-26 | Confirmed (tech. correction) | TS-13 |
| FR-08 | GST No. | FS #6 | F2 B6 | `J_1BBRANCH-GSTIN` | CL-30 | Open | TS-14 |
| FR-09 | CIN | FS #7 | F2 B7 | `T001Z` `ZCIN` | CL-27 | Open (config) | TS-14 |
| FR-10 | Exporter Ref./IEC | FS #2 | F2 B11, F3 | TBD | CL-05 | **Conflict** | TS-14 |
| FR-11 | PL no. / date | FS #4, #5 | F2 B9–B10, F3 | Z header | CL-16 | Open | TS-15 |
| FR-12 | SAP invoice number(s) in Z | FS #3 | F3 `VBRK_VBELN` | Link table `ZSD_EPACK_INV` | CL-01 | **Blocked** | TS-31 |
| FR-13 | Order No. & date | FS #8 | F2 B13 | `VBRP-AUBEL`→`VBAK-VGBEL`→`VBKD` | CL-15 | Open | TS-16 |
| FR-14 | Buyer's ref. & date | FS #9 | F2 B14 | Z header | CL-34 | Mostly confirmed | TS-15 |
| FR-15 | Country of origin | FS #10 | F2 B16 | Z header | CL-33 | Mostly confirmed | TS-15 |
| FR-16 | Country of final destination | FS #11 | F2 B17 | Ship-to country text | CL-13 | Confirmed | TS-17 |
| FR-17 | Sold-to block | FS #13 | F2 B20–B23, F5 | `VBPA`(AG)→`ADRC/ADR6/ADR2` | CL-13, CL-25 | Open | TS-17 |
| FR-18 | Ship-to block | FS #12 | F2 B26–B29, F4/F5 | `VBPA`(WE)→… | CL-13, CL-14 | Open | TS-17 |
| FR-19 | Notify parties 1–5 | FS #14 | F2 B31, F3, F5 | Z header long texts | CL-12 | Open (length) | TS-12 |
| FR-20 | Shipping details | FS #15–20 | F2 B33–B38, F3 | Z header | — | Confirmed | TS-15 |
| FR-21 | Marks, containers, no. of pkgs | FS #21, #22, #25 | F2 B39–B41, F3 | Z header (multi-line) | CL-12 | Open (length) | TS-15 |
| FR-22 | Description of goods (clubbed HSN) | FS #26 | F2 B42, F3 | TBD | CL-06 | **Conflict** | TS-18 |
| FR-23 | Payment term text | FS #23 | F2 B43 | `TVZBT-VTEXT` | CL-04 | Confirmed (multi-inv. open) | TS-19 |
| FR-24 | Incoterms | FS #24 | F2 B44 | `VBRK-INCO1/2` | CL-04, CL-20 | Confirmed (text open) | TS-19 |
| FR-25 | Line-item columns | FS #27 | F2 B47–B58, F3 | Item table + form main window | CL-08, CL-43 | **Blocked (CL-08)** | TS-20, TS-21 |
| FR-26 | Product description source | FS #27 | F2 B54, F3 G1/K5 | `MAKT` or Z | CL-07 | Open | TS-22 |
| FR-27 | Line HSN source | FS #27 | F2 B58, F3 O5 | `MARC-STEUC` or Z | CL-07 | Open | TS-22 |
| FR-28 | Declaration | FS #28 | F2 B61, F3 | Z header long text | CL-12 | Open (naming/length) | TS-23 |
| FR-29 | GW ≥ NW, error, no output | FS §2.1 | F3 (groups) | `ZCL_SD_EPACK_CHECK`, msg 001 | CL-10 | Open (level) | TS-05, TS-06, TS-07 |
| FR-30 | User validates data; master data correct | FS §1.1 | — | Upload validations reduce the risk | — | Assumption | TS-35 |
| FR-31 | Order-varying data in SO/invoice fields | FS §1 | — | Cannot map: fields unknown | CL-24 | **Unmappable until clarified** | — |
| IR-01 | Header on every page | — | F4, F5 | Secondary windows on all pages | CL-20 | Open | TS-24 |
| IR-02 | Page n of m | — | F4, F5 | `SFSY` fields | — | Inferred | TS-24 |
| IR-03 | Grand totals | — | F3 row 83, F4/F5 p4 | Calculated totals on last page | CL-40 | Inferred | TS-25 |
| IR-04 | HSN section headings | — | F4, F5 | Group loop by HSN | CL-09, CL-06 | Open | TS-18 |
| IR-05 | Package-type subtotals | — | F4 | Needs package-type field | CL-09 | **Unmappable until clarified** (no data source) | TS-25 |
| IR-06 | Package groups / merged cells | — | F3, F5 | Group ID + first-line printing | CL-08 | **Blocked** | TS-21 |
| IR-07 | Signature block | — | F4, F5 | Static text in footer window | CL-20 | Inferred | TS-24 |
| IR-08 | Cross-out unused rows | — | F4, F5 | Alternative (§8.1) | CL-38 | Open (feasibility) | TS-24 |
| IR-09 | Ship-to tax number | — | F5 | `KNA1-STCD*`/BP tax no. | CL-14 | Open | TS-17 |
| IR-10 | Contact person, FAX | — | F4, F5 | Contact-person partner / `ADR3` | CL-14 | Open | TS-17 |
| IR-11 | Consignee-only header (without-NP) | — | F4 | Form variant | CL-13 | Open (conflict) | TS-11 |
| IR-12 | Number formats | — | F4, F5 | Style/output masks | CL-28 | Open | TS-25 |
| IR-13 | Revision handling | — | F5 file name | Revision counter in header | CL-37 | Open | TS-36 |
| IR-14 | Upload validations | — | — | `ZCL_SD_EPACK_UPLOAD` | CL-11 | **Blocked** | TS-30 to TS-38 |
| IR-15 | Authorizations / messages / tests | FS §2.2, §2.3, §3.1 | — | §8 step 9, §10.3, §14 | CL-21, CL-22, CL-23 | Open | TS-40 to TS-42 |

---

## 14. Testing Strategy

### 14.1 Test levels

| Level | Scope | Tooling |
|---|---|---|
| Unit | Derivation and validation logic (`ZCL_SD_EPACK_DATA`, `ZCL_SD_EPACK_CHECK`, upload parser) | ABAP Unit with test doubles for DB reads |
| Integration | Quotation → SO → delivery → invoice → upload → `ZSD_EPACK` output | Q system with real document flow |
| Functional/UAT | Output compared with the approved layout and samples F4/F5 | Business users |
| Regression | Billing output, existing SD forms, other users of `ZSD_*` objects, commercial invoice (102-A) | Regression pack |
| Performance | Maximum invoices/lines per PL | Runtime analysis (SAT), SQL trace (ST05) |

### 14.2 Test scenarios

| TS | Scenario | Type | Expected result | Defined? |
|---|---|---|---|---|
| TS-01 | Run `ZSD_EPACK` with one valid invoice that has Z data | Positive | Output generated | Yes |
| TS-02 | Several invoices of one PL (several plants/SOs) | Positive | One consolidated PL | Partly (CL-04) |
| TS-03 | Selection contains invoices of two different PLs | Negative | Error (proposed msg 003) | **No (CL-02)** |
| TS-04 | Invoice without Z data / non-existent / cancelled / wrong billing type | Negative | Error, no output | Partly (CL-29) |
| TS-05 | GW < NW in one group | Negative | Message 001, **no output** | Yes (level per CL-10) |
| TS-06 | GW = NW | Boundary | Output allowed | Yes (FS "equal to or greater") |
| TS-07 | Weights blank / zero for a group | Boundary | — | **No (CL-10)** |
| TS-08 | Invoices with different sold-to / ship-to / payment terms / Incoterms / company code | Negative | Error or rule | **No (CL-04)** |
| TS-09 | Invoices from two business places (different GSTIN) | Negative | — | **No (CL-30)** |
| TS-10 | 1, 3 and 5 notify parties | Positive | With-NP form. Unused NP sections crossed out. | Partly (CL-38) |
| TS-11 | No notify party | Positive | Without-NP form (sold-to + ship-to, or consignee only?) | **No (CL-13)** |
| TS-12 | NP text at maximum length / with line breaks | Boundary | No truncation | **No (CL-12)** |
| TS-13 | Exporter address incl. region/country text | Positive | Matches F4/F5 | Yes |
| TS-14 | GST, CIN, IEC printed | Positive | Correct values | Partly (CL-05, CL-27) |
| TS-15 | All manual header fields printed | Positive | Equal to the upload | Yes |
| TS-16 | SO without quotation; several quotations | Boundary | — | **No (CL-15)** |
| TS-17 | Party blocks: ship-to ≠ master default, several emails/phones, missing email | Positive/negative | — | Partly (CL-13, CL-25) |
| TS-18 | Two or more HSN codes; same HSN with different descriptions | Positive | Clubbed description / section headings | **No (CL-06, CL-09)** |
| TS-19 | Payment term and Incoterm texts | Positive | Matches sample | Partly (CL-20) |
| TS-20 | 1 line, 90 lines (sample), 1,000 lines | Boundary/volume | Pagination correct, header on each page | Yes |
| TS-21 | Package group with 2+ materials across a page break | Boundary | Group not split / values on first line | **No (CL-08)** |
| TS-22 | Material without `MAKT` text / without `STEUC` / different `STEUC` per plant | Negative | Fallback rule | **No (CL-07)** |
| TS-23 | Declaration at maximum length | Boundary | Fully printed | **No (CL-12)** |
| TS-24 | Page n of m, signature block, crossing out on last page | Layout | Per approved layout | **No (CL-20, CL-38)** |
| TS-25 | Totals = Σ lines (sample 90000081: 4074 / 12146 / 116323.800 / 115386.780) | Positive | Exact match | Yes (values from F3) |
| TS-26 | Different customers/countries (UAE, Yemen) | Regression | Same layout | Yes |
| TS-30 | Upload valid template (F3 sample data) | Positive | Saved, log OK | Partly (CL-11) |
| TS-31 | Upload with invoice that does not exist / is cancelled | Negative | Rejected | Partly |
| TS-32 | Upload missing mandatory fields | Negative | Rejected with row number | **No (CL-11)** |
| TS-33 | Duplicate Sl.No. / duplicate PL no. | Negative | Rejected | **No (CL-11)** |
| TS-34 | Qty ≠ STD × Articles | Negative | Error/warning | **No (CL-41)** |
| TS-35 | Non-numeric weights, wrong dates, text too long | Negative | Rejected | Partly |
| TS-36 | Re-upload of an existing PL (revision) | Positive | Per revision rule | **No (CL-37)** |
| TS-37 | Concurrent upload/print of the same PL | Negative | Lock message 008 | Yes |
| TS-38 | Wrong template version / extra columns | Negative | Rejected | Partly |
| TS-39 | Reconciliation: Part No. not on invoice; quantity ≠ billed quantity | Negative | — | **No (CL-18)** |
| TS-40 | User without `S_TCODE` / sales-org authority / upload authority | Authorization | Message 006 | **No (CL-22)** |
| TS-41 | All messages per approved list | Negative | Texts as approved | **No (CL-21)** |
| TS-42 | Preview, print, PDF download | Positive | Per channel decision | **No (CL-19)** |
| TS-43 | Volume: about 50 invoices, 1,000 lines | Performance | Runtime < agreed threshold | **No (CL-32)** |
| TS-44 | Regression: billing output, document flow, other `ZSD_*` objects | Regression | Unchanged | Yes |

**Expected results that cannot be defined until the functional team answers:** TS-03, 07, 08, 09, 11, 12, 16, 18, 21, 22, 23, 24, 32, 33, 34, 36, 39, 40, 41, 42, 43 (see the "Defined?" column).

---

## 15. Performance and Clean Core Considerations

### 15.1 Performance
- **Expected volume is small**: the samples have 77–90 lines and probably fewer than 30 invoices per PL. Performance risk is low. Still:
  - Read `VBRK`/`VBRP` by primary key with one array SELECT. Join or `FOR ALL ENTRIES` for `VBPA`, `VBAK`, `VBKD`, `MARC`, `MAKT`, `T604N`. Never SELECT inside loops or inside the Smart Form.
  - A secondary index on the invoice link table by `VBELN` (normalised design) makes invoice→PL lookup fast. With the FS CHAR255 design a `LIKE '%…%'` scan would be needed (not recommended).
  - Buffer customizing texts (`T005T`, `T005U`, `TVZBT`, `T604N`) in hashed tables.
  - Upload: parse once, validate in memory, then one `INSERT`/`MODIFY` per table within one LUW.
- Volume assumption to be confirmed (CL-32).

### 15.2 Clean core (feasibility depends on release – CL-31)
| Topic | Assessment |
|---|---|
| Smart Forms | A classic technology, still supported in S/4HANA **[SAP-VERIFY]**. Not ABAP-Cloud compliant. The FS requires Smart Forms explicitly. Keep the forms layout-only to limit lock-in. |
| Z tables / custom reports | Classic custom code. No modification of SAP objects is needed. No enhancements or user exits are required by this design. |
| Data access | In S/4HANA, consider released CDS views (e.g. `I_BillingDocument`, `I_BillingDocumentItem`, `I_SalesQuotation`) instead of direct table reads. **Availability and release status must be verified** in the target system. India-specific tables (`J_1BBRANCH`, `T001Z`) may have no released equivalent. |
| Excel parsing | Prefer a released API where available (e.g. `XCO_CP_XLSX` on suitable S/4HANA releases) over unreleased classes **[SAP-VERIFY]**. |
| Hard-coding | Language, country `IN`, partner functions, billing types and CIN party type should be constants or TVARVC entries, not literals. |

---

## 16. Development Readiness Assessment

### Classification: **NOT READY – critical clarifications are required first**

**Justification**
1. The **data model** in the FS (Z-01, Z-02) cannot store multiple line items per packing list, and cannot support selection by invoice. Every downstream object (upload, data provider, form interface) depends on it.
2. The **link between the selection (invoices) and the data key (PL no.)** is undefined. So is the **consolidation of header data across several invoices**, which is the core of a *consolidated* packing list.
3. The **Excel upload**, which is the only source for most of the content, has no specification.
4. The samples show **package grouping** that is not modelled and directly affects the only defined validation (GW ≥ NW).
5. The FS is **unapproved**, with empty Messages, Authorization and Test sections, and it has several internal contradictions (§6.3).

**Work that can start now with low rework risk** (optional, if the project accepts it):
- Prototype the SAP header extraction (exporter, GSTIN, CIN, partners, payment terms, Incoterms) in `ZCL_SD_EPACK_DATA` with the corrected tables (§6.4).
- Smart Style and page skeleton (header windows, page n of m, signature) based on F4/F5.
- Technical spikes: Excel parser on the target release. Cross-out graphic feasibility in Smart Forms.

### Questions by timing

| Category | Items |
|---|---|
| **Must be answered before development** | CL-01, CL-02, CL-04, CL-08, CL-11 (critical) and CL-03, CL-12, CL-20, CL-31 (they shape DDIC and layout) |
| **Can be answered during development** | CL-05, CL-06, CL-07, CL-09, CL-10, CL-13, CL-14, CL-15, CL-16, CL-17, CL-18, CL-19, CL-22, CL-24, CL-28, CL-29, CL-30, CL-37, CL-38 |
| **Can be deferred to testing without major rework** | CL-21, CL-23, CL-25, CL-26, CL-27, CL-32, CL-33, CL-34, CL-35*, CL-36, CL-39, CL-40, CL-41, CL-42, CL-43 (*FS sign-off is still required before transport to production) |

### Inputs required to proceed

| Item | Purpose |
|---|---|
| Signed-off FS V.02 with the answers to §11 | Baseline |
| Approved layout for both formats (editable, with dimensions/fonts/logo) | Form build |
| Upload specification + final Excel template (with package group / package type columns if approved) | Upload build |
| 3–5 real test cases: invoice numbers in DEV/QAS, filled Excel files, expected PDFs (with and without NP, multi-plant, multi-HSN) | Testing |
| Confirmation of SAP release, naming conventions, existing objects | Technical design |
| Access: DEV system with SD export document flow, `SMARTFORMS`, `SE11`, `SE78`, `SU21` | Development |
| Authorization concept / role owners | Security |

---

## 17. Recommended Next Steps

| # | Action | Owner | Output |
|---|---|---|---|
| 1 | Send §11 (at least the C and H items) to the SD functional consultant and Astral business owner. Run a 60-minute walkthrough using samples F4/F5 and `Z table Format1.xlsx`. | Technical lead | Answered clarification log |
| 2 | Decide the Z-table design (CL-01) and the package-group concept (CL-08). Update the Excel template accordingly. | Functional + technical | Approved data model + template v2 |
| 3 | Write the upload specification (CL-11), incl. revision handling (CL-37). | Functional consultant | FS addendum |
| 4 | Resolve the source conflicts X-01 to X-12 (IEC, description of goods, HSN, product description, partner blocks). | Functional consultant | FS V.02 |
| 5 | Provide the approved layout and accept the Smart Forms feasibility constraints (§8.1, CL-38). | Business | Layout sign-off |
| 6 | Complete FS §2.2 (messages), §2.3 (authorization) and §3.1 (test cases), using §10.3 and §14 as a starting point. | Functional consultant | FS V.02 |
| 7 | Confirm the release and naming. Run technical spikes (Excel parser, SE78 cross-out graphic). | Technical lead | Spike results |
| 8 | Prepare the technical specification from this document and the answers. Then build in the order DDIC → upload → data provider/checks → driver → forms. | Developer | TS + objects |
| 9 | Prepare test data (invoices + Excel) for the scenarios in §14. | Functional / test team | Test pack |

---

*End of document.*
