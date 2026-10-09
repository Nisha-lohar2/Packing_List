# Export Packing List report `ZSD_EPACK`: design, field mapping and test plan

WRICEF 102-B "Export consolidated packing list" · Astral Limited / UDAY / SD

This document covers the Excel-upload report. It reads the packing list Excel file,
fetches header and line-item data exactly as described in the FS logic sheet,
and produces **one Smart Form per packing list number**. It then saves the data
to the Z tables and writes a log row per packing list.

Sources read for this design (all in the repository):

| File | Used for |
|---|---|
| `Object ID- 102(B)Export consolidated packing list.docx` (FS V.01) | Process, scope, two form formats, T-code `ZSD_EPACK`, input "multiple tax invoices", Z table `ZSD_EPACK_DATA`, GW ≥ NW rule |
| `Logic_sheet1.xlsx` › **Logic** | Field-by-field data retrieval (rows B3–B61, quoted below) |
| `Logic_sheet1.xlsx` › Input | Selection "Tax Invoice From–To, multiple selection" |
| `Z table Format1.xlsx` › Format / Technical detail / Sheet3 | Upload layout (columns A–AK), Z field names, lengths, key fields, 77-line example |
| `PackingList_90000014 … .pdf`, `PackingList_90000081 … .pdf` | Expected output layout and totals |
| `src/*` (this development) | Reused and extended: rules class, exception, interface, output, DDIC |

No other ABAP, CDS, form or table definition exists in the repository. Nothing from SAP standard is modified.

---

## 1. Excel file structure and Logic-sheet business logic

### 1.1 Upload file (`Z table Format1.xlsx`, sheet "Format")

| Columns | Content | Level |
|---|---|---|
| A | Sl.No. | line item |
| B | Sr.No. INV (line no. of the commercial invoice) | line item |
| C | Article No. ("1 TO 60", "STACKER") — merged over package groups | line item / group |
| D | STD | line item |
| E | Articles — merged over package groups | line item / group |
| F | Part No. (= SAP material, Logic B58 "from STEP no 53") | line item |
| G | Product Description ("Fetch from material master") | line item |
| H | Quantity Pcs | line item |
| I, J | Gross / Net Wt. KGS — merged over package groups | weight group |
| K | H.S Code | line item |
| **L** | **SAP Invoice Number** (Technical detail: key field `VBRK_VBELN`, "Store multiple tax invoice") | packing list |
| **M** | **Packing list no** (key field `ZSD_PACKNO`) | packing list |
| N–AK | PL date, Exporter Ref, Buyer's ref. no./date, country of origin, notify parties 1–5, pre-carriage, place of receipt, POL, POD, place of delivery, marks, vessel, containers, no. of packages, description of goods, declaration, three commercial-invoice fields | packing list |
| (AL) | Package Type — **optional, new** (A09) | line item |
| last row | Totals (no Sl.No.) — ignored | — |

The sample sheet has 77 line items (= PDF 90000081) with totals 4074 articles,
12146 pcs, 116323.800 kg gross and 115386.780 kg net. Header columns L–AK are empty in the sample.

### 1.2 Logic sheet — what it says

* **Header data from SAP, driven by the invoice number (`VBELN`)**: exporter
  (B3–B5), GST (B6), CIN (B7), IEC (B11), order no. & date from the quotation
  (B13), sold-to and ship-to with e-mail and phones (B20–B29), country of final
  destination (B17), description of goods (B42), payment term (B43) and Incoterms (B44).
* **Header data from the Z table, driven by `ZSD_PACKNO`**: PL date, buyer's ref.,
  country of origin, notify parties, shipping fields, marks, containers,
  no. of packages and declaration (B9–B41, B61).
* **Line items from the Z table, driven by `ZSD_PACKNO`** (B47–B58).
  Product description and HSN come from the Z table **or** the material master (B54, B58).
* **Relationship.** One packing list number → several tax invoices (Technical detail key
  `ZSD_PACKNO` + `VBRK_VBELN` "multiple tax invoice") → their billing items → sales
  orders → quotations. The line items belong to the packing list, **not** to an
  individual invoice. The logic sheet has no item-to-invoice link.

---

## 2. Field mapping

`PL` = the packing list being printed. `INV` = all invoices of that PL.
"Excel col." is the upload column. "Logic" is the logic-sheet row.

### 2.1 Header — from SAP

| Output field | Logic | Source table / field | Derivation in `ZCL_SD_EPACK_DATA` |
|---|---|---|---|
| Exporter name / address | B3 | `VBRK-BUKRS` → `T001-ADRNR` → `ADRC` `NAME1 STREET STR_SUPPL1 CITY1 POST_CODE1 REGION COUNTRY` | `READ_SAP_DATA` (T001, ADRC once), `EXPORTER_LINES` |
| Exporter region text | B4 | `T005U-BEZEI` (LAND1 + BLAND = `ADRC-REGION`, SPRAS E) | **Logic says `T005S`**, which has no `BEZEI` field (§3) |
| Exporter country text | B5 | `T005T-LANDX` (SPRAS E) | **Logic says `T005`**, which has no `LANDX` field (§3) |
| GST No. | B6 | `J_1BBRANCH-GSTIN` for `VBRK-BUKRS` + `VBRK-BUPLA` | printed in the header box and under the exporter (as in the samples) |
| CIN No. | B7 | `T001Z-PAVAL` for `BUKRS`, `PARTY` = `ZCIN` | |
| Exporter's Ref. (IEC) | B11 | `ADRC-BUILDING` of the company address | printed "IEC:<value>" (sample). Excel col. O is **not** used (Technical detail: "Not Req – fetched from SAP"). Q08 |
| Order No. & Date | B13 | `VBRP-AUBEL` → `VBAK-VGBEL` → `VBKD-BSTKD`, `BSTDK` (`VBELN` = VGBEL, `POSNR` 000000) | all distinct values of all INV, "A & B DATE: d1 & d2" (Q06) |
| Country of final destination | B17 | ship-to `ADRC-COUNTRY` → `T005T-LANDX` | |
| Sold-to block | B20–B23 | `VBRK-KUNAG` + sales area → `KNVP` (`PARVW` = AG ≙ SP) → `KUNN2` → `KNA1-ADRNR` → `ADRC` `NAME1 STREET STR_SUPPL1 CITY1 CITY2 POST_CODE1 COUNTRY`, `ADR6-SMTP_ADDR`, `ADR2-TELNR_LONG` (`R3_USER` 1 = Telephone, 3 = Mobile) | `PARTNER`, `PARTY_LINES` |
| Ship-to block | B26–B29 | same with `PARVW` = WE ≙ SH | `PARTNER`, `PARTY_LINES` |
| Description of goods | B42 | `VBRP-MATNR` (+ `WERKS`) of INV → `MARC-STEUC` → `T604N-TEXT1` (SPRAS E, LAND1 IN); common HSN "clubbed" across all billing documents | `GOODS_LINES`; one line per distinct HSN, "<text> - HS CODE : 3917 2390". Excel col. AG overrides when filled (Q05) |
| Payment term | B43 | `VBRK-ZTERM` → `TVZBT-VTEXT` (SPRAS E, first ZTAGG) | "PAYMENT TERM: …" |
| Terms of shipment | B44 | `VBRK-INCO1`, `VBRK-INCO2` | "TERMS OF SHIPMENT :- EXW AHMEDABAD" (codes only, Q10) |

### 2.2 Header — from the upload (Z table)

| Output field | Logic | Excel col. | Z field |
|---|---|---|---|
| Packing list no. | B9 | M | `ZSD_PACKNO` |
| Date | B10 | N | `ZSD_PACKDT` (DD-MM-YYYY on output) |
| Buyer's Reference No. & Date | B14 | P, Q | `ZSD_PARTY_REF`, `ZSD_REF_DT` |
| Country of origin | B16 | R | `ZSD_CNTY_ORGN` |
| Notify parties 1–5 | B31 | S–W | `ZSD_NP1`…`ZSD_NP5` (multi-line) |
| Pre-carriage, place of receipt, vessel, POL, POD, place of delivery | B33–B38 | X, Y, AD, Z, AA, AB | `ZSD_PREC`, `ZSD_REC_PC`, `ZSD_VSL_FLT`, `ZSD_POL`, `ZSD_POD`, `ZSD_PLD` |
| Marks & Nos., Container No., No. of packages | B39–B41 | AC, AE, AF | `ZSD_MARK`, `ZSD_CONT`, `ZSD_NOPACK` |
| Declaration | B61 | AH | `ZSD_PL_DECL` (logic says `ZSD_DECL`; the Technical-detail name is used) |
| Description of goods (override) | — | AG | `ZSD_HSN_DESC` (Q05) |
| Invoices of the PL | — | L | `ZSD_EPACK_INV-VBELN` |
| Stored but not printed | — | AI, AJ, AK | `ZSD_ADV_AMT`, `ZSD_CMMT`, `ZSD_INV_DECL` (commercial-invoice fields) |

### 2.3 Line items

| Output column | Logic | Excel col. | Z field | Derivation |
|---|---|---|---|---|
| Sl.No. | B47 | A | `ZSD_SN` | unique and ascending per PL |
| Sr.No. INV | B48 | B | `ZSD_INV_SN` | |
| Article No. | B49 | C | `ZSD_ARTNO` | printed on the first line of a merged group |
| STD | B50 | D | `ZSD_STD` | |
| Articles | B51 | E | `ZSD_ART` | first line of an article group (A06) |
| Total Articles | B52 | — | (`ZSD_TART`) | **calculated** as Σ Articles. The template has no column for it (Q12) |
| Part No. | B53 | F | `ZSD_PARTNO` | = material number |
| Product Description | B54 | G | `ZSD_MATDESC` | else `MAKT-MAKTX` of Part No. (**logic says `MARA`**, §3) |
| Quantity Pcs | B55 | H | `ZSD_QTY` | |
| Gross / Net Wt. KGS | B56, B57 | I, J | `ZSD_GWT`, `ZSD_NWT` | first line of a weight group (A06); 3 decimals |
| H.S Code | B58 | K | `ZSD_HSN` | else `MARC-STEUC` of Part No. in the plant where it is billed on INV |
| HSN section heading, subtotal rows, totals | — | (AL) | — | from samples (A09, Q18) |

---

## 3. Deviations from the Logic sheet (technical, required)

The Logic sheet is followed row by row. In five places it names a field that
does not exist in the named table, or a display code instead of the stored
value. A literal implementation would not activate.

| Logic | Logic sheet text | Implemented | Reason |
|---|---|---|---|
| B4 | "BLAND in **T005S** get BEZEI" | `T005U-BEZEI` | `T005S` has no text field; region texts are in `T005U` |
| B5, B20, B26 | "LAND1 in **T005** get LANDX" | `T005T-LANDX` | `T005` has no text field; country texts are in `T005T` |
| B54 | "MATNR in **MARA** get MAKTX" | `MAKT-MAKTX` | Material texts are in `MAKT` |
| B20, B26 | "PARVW = **SP** / **SH** … get **KUNNR**" | `PARVW` = `AG` / `WE`, result `KUNN2`, plus `VKORG/VTWEG/SPART` of the invoice | SP/SH are the English display codes; the partner number is `KNVP-KUNN2`; the sales area is part of the KNVP key |
| B13 | "VBELN (**AUBEL**) in VBAK" | `VBRP-AUBEL` | `AUBEL` is a billing **item** field |

All five need a one-line confirmation from the technical lead (Q11). None changes the business meaning.

---

## 4. Assumptions and questions requiring functional confirmation

Each implemented assumption has an id (`Axx`) in the code. `docs/Development_Assumptions.md` lists them all.

| Q | Topic | Question | Implemented for now | Impact if different |
|---|---|---|---|---|
| Q01 | Several PLs per file | Can one Excel file hold several packing lists? Is "rows without PL no. belong to the PL above" the right rule? | Yes. Column M is carried forward (A33) | Parser split rule |
| Q02 | Invoice in two PLs | May one tax invoice belong to two packing lists (partial shipments)? | No: error for both PLs (A01) | Remove one check |
| Q03 | Duplicate rows | Should an exact duplicate row (same Sl.No., same values) be ignored, or rejected? | Ignored with a warning; same Sl.No. with different values is an error (A38) | One rule |
| Q04 | Invoices of one PL disagree | If invoices of one PL have different company code, business place, sold-to, sales area, payment terms or Incoterms — reject, or print which value? | Reject the PL (A04). No "first invoice" choice | Header logic |
| Q05 | Description of goods | Logic B42 (SAP) vs template column AG ("Same HSN having different description"). Which wins? | AG if filled, else B42 (A14) | One condition |
| Q06 | Order No. & Date | Several sales orders / quotations in one PL: print all, joined with " & "? (Sample 90000014 shows two references.) | Yes (A17) | Formatting only |
| Q07 | Several ship-to partners | Customer has several WE partners in KNVP: use the default partner? | Default (`DEFPA`), else lowest counter, with a warning (A31) | Partner selection |
| Q08 | IEC | Is the IEC maintained in company address field BUILDING in all company codes? | Yes (B11) | Data source |
| Q09 | Sample-only content | Samples show contact person "K/A", FAX, ship-to TAX NUMBER, and a single "CONSIGNEE" block. None is in the Logic sheet. Required? | Not filled. One CONSIGNEE block when sold-to = ship-to (A15) | Add 3 reads |
| Q10 | Incoterms | B44 prints codes (INCO1, INCO2). The sample prints "EX-WORKS" (a text). Print the `TINCT` description? | Codes only | One read |
| Q11 | Technical deviations | Confirm the five corrections in §3 | As in §3 | — |
| Q12 | Total Articles | B52 reads `ZSD_TART`, but the template has no column. Is calculating Σ Articles correct? | Calculated (A08) | — |
| Q13 | Weight check level | FS: GW ≥ NW. Per package group, or per line? (Weights are given per merged group.) | Per weight group, at upload and at output (A07) | Rule granularity |
| Q14 | Authorization | FS §2.3 is empty. Approve object `ZSD_EPACK` (ACTVT 01/04, BUKRS) + `V_VBRK_VKO`? | As stated (A24) | Roles |
| Q15 | When to save | Save the uploaded data after **preview** as well, or only after print / PDF? | After any successful output, incl. preview. Never in "Validate only" (A39) | One condition |
| Q16 | "Z table created for log" | Is the log the FS Z table itself (`ZSD_EPACK_DATA` + header + invoice tables), a separate run log, or both? | **Both**: data → `ZSD_EPACK_HDR/INV/DATA`, run log → `ZSD_EPACK_LOG` (A40) | Drop one table |
| Q17 | Re-upload | Uploading an existing PL replaces it and increases the revision. Is history needed? | Replace, revision + 1 (A22) | Versioning table |
| Q18 | Package type | Sample 90000014 shows "Loose Pipes / Boxes" subtotals. Is the new optional column AL acceptable? | Yes (A09) | Template |
| Q19 | Field lengths | Notify parties, containers and declaration texts in the samples exceed the FS lengths. Approve longer fields? | Yes (A19) | DDIC |
| Q20 | Layout | Are the two sample PDFs the approved layout? Is crossing out with a graphic acceptable? | Yes (A26–A28) | Form |
| Q21 | Release | SAP release / Excel API (`CL_FDT_XL_SPREADSHEET`) | S/4HANA ≥ 7.54 (A21, A29) | Parser |
| Q22 | Billing types | Which billing types are export tax invoices (TVARVC `ZSD_EPACK_FKART`)? | Empty variable = all accepted (A11) | Customizing |
| Q23 | Missing master data | Missing GSTIN / CIN / IEC / description / HSN: warning (print with an empty field), or error? | Warning (A36) | Severity |

---

## 5. Data flow

```
 ZSD_EPACK selection screen
   (•) Upload Excel file  ( ) Saved packing lists (reprint by invoice / PL no.)
   format: automatic / with NP / without NP   output: preview / print / PDF folder
   [ ] Validate only
          │
 1  ZCL_SD_EPACK_UPLOAD
   read_file (GUI_UPLOAD) → CL_FDT_XL_SPREADSHEET → rows
   check template (A1 "Sl.No.", M1 "Packing list no")
   split rows by column M (carry forward)   ──► tt_packing_list (sorted by PL no.)
   per PL: header (N–AK, conflicts = error), invoices (L, de-duplicated),
           line items (A–K, AL; merged-cell groups; duplicate rows),
           FS weight rule, quantity plausibility        → messages per PL
   cross-check: invoice in two PLs → error in both
          │
 2  ZCL_SD_EPACK_DATA->PREPARE  (bulk)
   READ_SAP_DATA: union of all invoices → VBRK, VBRP, TVARVC, KNVP, KNA1, T001,
     T001Z, J_1BBRANCH, ADRC, ADR6, ADR2, T005T, T005U, VBAK, VBKD, TVZBT, MAKT,
     MARC, T604N — each ONCE (FOR ALL ENTRIES), into HASHED / SORTED tables
   per PL (BUILD_ONE, internal tables only):
     invoices exist / not cancelled / billing type  → error
     invoices agree on header fields (Q04)          → error
     authorization (BUKRS 04, V_VBRK_VKO 04)       → error
     GW ≥ NW (FS 2.1)                               → error, no form
     partners via KNVP (B20/B26)                    → error if missing
     header + text blocks + items (B3 … B61)        → ts_result (status S/W/E)
          │
 3  ZSD_EPACK_PRINT->OUTPUT_FORMS
   LOOP over results (one per PL), skip status E
     ZCL_SD_EPACK_OUTPUT->SEND: SSF_FUNCTION_MODULE_NAME + form call with
     IS_HEADER / IT_TEXT / IT_ITEM of THIS PL only → preview / spool / PDF file
     error → logged on this PL, loop continues
          │
 4  ZCL_SD_EPACK_STORE  (skipped in "Validate only")
   SAVE_PACKING_LISTS: PLs with output = X → ZSD_EPACK_HDR / _INV / _DATA
     (lock, revision, array DELETE + INSERT, rollback on DB error)
   WRITE_LOG: one ZSD_EPACK_LOG row per PL + one for file-level errors
   FINISH: COMMIT WORK, release locks
          │
   Result list (ALV): per PL status, form, invoices, items, output, saved, messages
```

**Grouping guarantees.** `tt_packing_list` and `tt_result` are sorted tables
keyed by the packing list number. Every entry carries only its own header,
invoices, items and messages. The form is called with the data of one entry at a
time. Billing items are collected per PL only from that PL's invoices.

**Performance.** The Excel is parsed once, with no DB access. Data retrieval is **19 SELECTs per run**
(14 in `READ_SAP_DATA`, 5 in `READ_ADDRESSES`), independent of the number of packing lists, invoices
or items. Add one read of existing headers before saving and one read of the default printer.
There is no SELECT in a loop.
Lookups use hashed / sorted keys. The writes are 2 DELETE, 1 MODIFY, 2 INSERT and 1 INSERT (log), all
array operations. AUTHORITY-CHECK results are buffered per value.

---

## 6. Test plan

`UT` = automated ABAP Unit (`ZCL_SD_EPACK_UPLOAD` 20 tests, `ZCL_SD_EPACK_RULES`
26 tests). `IT` = integration test in QAS with real invoices. Expected results
marked **(Qxx)** depend on the answer to that question.

| # | Scenario | Test data | Expected result | Type |
|---|---|---|---|---|
| T01 | One invoice, one PL | 1 PL, 1 invoice, 2 lines | 1 PL with 1 invoice, 2 items; 1 form; saved; log status S | UT `one_invoice_one_pl` + IT |
| T02 | Several invoices, same PL | invoices in one cell "A, B" and in another row | 3 invoices on the PL; header from SAP identical for all; 1 form | UT `several_invoices_one_pl` + IT |
| T03 | Several PLs in one file | PL 90000014 rows + PL 90000081 rows | 2 results, **2 separate forms** (2 previews / 2 spool requests / 2 PDF files), 2 log rows | UT `several_pls_in_one_file` + IT |
| T04 | Items belong to their own PL only | parts PART-A* under PL A, PART-B* under PL B | form A shows only PART-A*, form B only PART-B* | UT `items_never_cross_pls` + IT |
| T05 | Several line items / merged groups | full `Example_90000081` sheet (77 lines) | 77 rows; Articles / weights on the first line of each group; totals 4074 / 12146 / 116323.800 / 115386.780; text "(Total 4074 Bundles)" | UT rules tests + IT, compare with PDF 90000081 |
| T06 | Two package types, two HSN | Tarmac data (sample 90000014) | headings "PP LOW NOISE PIPES … 3917 2390", "… FITTINGS … 3917 4000"; "2342 <=Loose Pipes", "1073 <=Boxes"; "(Total 2342 Loose Pipes + 1073 Boxes = 3415 Packages)" | UT `print_rows_heading_subtotal` + IT |
| T07 | Duplicate invoice numbers | same invoice in 2 rows of a PL | counted once, info 056, no error | UT `duplicate_invoice_once` |
| T08 | Same invoice in two PLs | invoice under PL A and PL B | error 058 on both, no form for either **(Q02)** | UT `invoice_in_two_pls_rejected` |
| T09 | Invalid invoice token | "INV-X" | error 018, no form for that PL | UT `invalid_invoice_rejected` |
| T10 | Invoice not in SAP | 0099999999 | error 008, no form; other PLs printed | IT |
| T11 | Cancelled invoice | cancelled billing doc | error 009 | IT |
| T12 | Missing invoice / PL date | empty L or N | error 017 | UT `missing_invoice_rejected`, `missing_pl_date_rejected` |
| T13 | PL without line items | header row only | error 036, no form; other PL printed | UT `pl_without_items_rejected` + IT |
| T14 | Rows above the first PL no. | data row before column M is filled | file-level error 055; the rows are ignored | UT `rows_before_first_pl` |
| T15 | No PL number at all | column M empty | error 059, nothing printed, log row | UT `no_pl_number_in_file` + IT |
| T16 | Wrong / empty file | other layout; heading only | error 016 / 036 | UT `wrong_template_rejected`, `empty_file_rejected` |
| T17 | Header conflict inside a PL | POL differs between rows of a PL | error 020; different POL in **different** PLs is fine | UT `header_conflict_in_pl`, `header_differs_between_pls` |
| T18 | Duplicate rows | identical row twice / same Sl.No. with other values | warning 057, counted once / error 021 **(Q03)** | UT `exact_duplicate_row_ignored`, `conflicting_sl_no_rejected` |
| T19 | GW < NW | one group 100 < 101 | message 001 (exact FS text) + 002, **no form** for this PL, others printed | UT `weight_error_found_in_upload`, rules tests + IT |
| T20 | Invoices of one PL disagree | 2 invoices with different payment terms | error 011, no form **(Q04)** | IT |
| T21 | No KNVP ship-to / several ship-to | customer without WE / with 2 WE | error 042 / warning 060, default partner used **(Q07)** | IT |
| T22 | Header data per Logic sheet | 1 PL against SE16 values | exporter = T001/ADRC; GST = J_1BBRANCH; CIN = T001Z ZCIN; IEC = ADRC-BUILDING; order no. = quotation VBKD-BSTKD/BSTDK; parties = KNVP→KNA1→ADRC/ADR6/ADR2; goods = VBRP→MARC→T604N clubbed; payment = TVZBT; Incoterms = INCO1 INCO2 | IT |
| T23 | Item fallback to material master | G and K empty | description = MAKT, HSN = MARC-STEUC of the billing plant; warnings 048 / 044 if missing | IT |
| T24 | Part not on any invoice | Part No. not billed | warning 030; HSN from upload only | IT |
| T25 | Format selection | 0 NP / 1–5 NP / "without" forced with NP | without-NP form / with-NP form with unused windows crossed out / error 013 | IT |
| T26 | Form missing / form error | deactivate one form | error 040 / 039 logged for that PL; the loop continues with the next PL | IT |
| T27 | Preview cancelled | cancel in the preview dialog | warning 064, status W, not saved | IT |
| T28 | PDF per PL | PDF + folder, 2 PLs | `PackingList_90000014.pdf`, `PackingList_90000081.pdf` | IT |
| T29 | Save and log | after T03 | `ZSD_EPACK_HDR` 2 rows (revision 0), `_INV` / `_DATA` complete as uploaded; `ZSD_EPACK_LOG` 2 rows (S) with invoices, counts, form | IT |
| T30 | Re-upload | T03 file again | revision 1, data replaced, new log rows **(Q17)** | IT |
| T31 | Validate only | any file | result list, no form, no DB change, no log | IT |
| T32 | Reprint saved | option "Saved packing lists" by invoice | PL found through its invoice, printed with all its invoices (info 006), print stamp set for print / PDF | IT |
| T33 | Authorization | user without `ZSD_EPACK` 04 / `V_VBRK_VKO` / 01 | error 012, no form / form printed but not saved (warning 012) **(Q14)** | IT |
| T34 | Volume | 20 PLs × 100 lines, 100 invoices | runtime dominated by form output; SQL trace shows 19 data SELECTs | IT (ST05) |
| T35 | Regression | billing output, VF03, other `ZSD_*` objects | unchanged — no SAP object modified | IT |

---

## 7. Created and modified development objects

| Object | Type | Status in this delivery |
|---|---|---|
| `ZSD_EPACK_PRINT` (T-code `ZSD_EPACK`) | Report | **Reworked**: Excel upload → one form per PL → save + log |
| `ZCL_SD_EPACK_UPLOAD` + test class | Class | **Reworked**: parser with no DB access, several PLs per file, 20 unit tests |
| `ZCL_SD_EPACK_DATA` | Class | **Reworked**: bulk reads for all PLs, Logic-sheet derivations (KNVP partners, Incoterm codes, strict B11/B42) |
| `ZCL_SD_EPACK_STORE` | Class | **New**: save to Z tables, log table |
| `ZCL_SD_EPACK_OUTPUT` | Class | Extended: result list, output result for preview |
| `ZIF_SD_EPACK` | Interface | Extended: PL list / result types, status, source |
| `ZCL_SD_EPACK_RULES` + tests, `ZCX_SD_EPACK` | Class | Reused unchanged |
| `ZSD_EPACK_LOG`, `ZSD_EPACK_STATUS` | Table, data element | **New** |
| `ZSD_EPACK_HDR`, `ZSD_EPACK_INV`, `ZSD_EPACK_DATA`, structures, lock object | DDIC | Unchanged design |
| Message class `ZSD_EPACK` | Messages | 055–067 new, 004 / 006 / 011 / 042 reworded, upload-only messages withdrawn |
| `ZSD_EPACK_UPLOAD` / T-code `ZSD_EPACK_UPL` | Report | **Withdrawn**, replaced by `ZSD_EPACK` |
| `ZSD_EPACK_WITH_NP`, `ZSD_EPACK_WITHOUT_NP`, style, graphic | Smart Forms | Specification updated (one call per PL, Logic-sheet block content) |

Full DDIC details: `src/DDIC_and_Repository_Objects.md`. Form build instructions:
`docs/ZSD_EPACK_SmartForm_Specification.md`. Upload template:
`templates/ZSD_EPACK_Upload_Template_v2.xlsx`.
