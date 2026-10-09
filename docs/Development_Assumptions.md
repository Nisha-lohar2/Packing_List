# WRICEF 102-B Export Packing List — development status and assumptions

**Version 2.** This version covers the Excel-upload report `ZSD_EPACK`: one form per
packing list, then saving and logging. The assumptions below are referenced in the source code as `Axx`.
The open questions that decide them are in `Packing_List_Report_Design.md` §4 (`Qxx`).
The requirement analysis with all 43 original clarifications is in
`Packing_List_Requirement_Analysis_and_Development_Approach.md`.

## 1. Status

| Deliverable | File | Status |
|---|---|---|
| Report (T-code `ZSD_EPACK`) | `src/ZSD_EPACK_PRINT.prog.abap` | Code complete |
| Excel parser + 20 unit tests | `src/ZCL_SD_EPACK_UPLOAD.clas.abap`, `.testclasses.abap` | Code complete |
| Data retrieval per Logic sheet (bulk) | `src/ZCL_SD_EPACK_DATA.clas.abap` | Code complete |
| Business rules + 26 unit tests | `src/ZCL_SD_EPACK_RULES.clas.abap`, `.testclasses.abap` | Code complete (unchanged) |
| Form output + result list | `src/ZCL_SD_EPACK_OUTPUT.clas.abap` | Code complete |
| Save to Z tables + log | `src/ZCL_SD_EPACK_STORE.clas.abap` | Code complete |
| Types / exception | `src/ZIF_SD_EPACK.intf.abap`, `src/ZCX_SD_EPACK.clas.abap` | Code complete |
| DDIC incl. log table, messages, auth | `src/DDIC_and_Repository_Objects.md` | Specification |
| Smart Forms | `docs/ZSD_EPACK_SmartForm_Specification.md` | Build specification |
| Design, mapping, test plan | `docs/Packing_List_Report_Design.md` | Done |
| Upload template + example | `templates/ZSD_EPACK_Upload_Template_v2.xlsx` | Done |

**Not verified on a system.** No SAP system was available. The code has not been
compiled or run. A static check confirmed balanced blocks, no DB statement inside a
loop, and no untyped inline declarations. The upload rules were cross-checked offline against the
example sheet: 77 lines, totals 4074 / 12146 / 116323.800 / 115386.780, matching PDF 90000081.

## 2. Assumptions applied

| ID | Assumption | Basis | Question |
|---|---|---|---|
| A01 | An invoice may belong to only one packing list in a file. If it appears under two, both packing lists get an error. | FS silent | Q02 |
| A04 | All invoices of one packing list must agree on company code, business place, sold-to, sales area, payment terms and Incoterms. Otherwise the packing list is rejected. No "first invoice" is chosen. | Logic sheet reads one VBRK; the user asked not to pick an arbitrary invoice | Q04 |
| A05 | The format is chosen automatically: "with notify party" when any notify party is uploaded. A manual choice of "without" while notify parties exist is an error. | FS 1.1 (labels a/b contradict) | — |
| A06 | Merged Excel cells mean package groups: an empty Articles cell, or empty Gross + Net cells, belong to the group above. | Template merged ranges, samples | Q13 |
| A07 | GW ≥ NW is checked per weight group, at upload and again before output. An error means no form for that packing list. | FS 2.1 | Q13 |
| A08 | Total Articles is calculated (Σ Articles). Quantity ≠ STD × Articles gives a warning only. | Template has no `ZSD_TART` column | Q12 |
| A09 | Optional column AL "Package Type" drives the subtotal rows and the total text. | Sample 90000014 | Q18 |
| A10 | Notify parties must be filled from 1 upwards. | Form layout | — |
| A11 | TVARVC `ZSD_EPACK_FKART` limits the billing types. If it is empty, all types are accepted. Cancelled invoices are always rejected. | FS silent | Q22 |
| A14 | Description of goods: Excel column AG if filled, else Logic B42 (VBRP→MARC→T604N, clubbed). | Logic B42 vs Technical detail `ZSD_HSN_DESC` | Q05 |
| A15 | One "CONSIGNEE" block when sold-to = ship-to (sample 90000014). Contact person, FAX and tax number are **not** printed (not in the Logic sheet). | Samples vs Logic sheet | Q09 |
| A17 | Order No. & Date: all distinct quotation references and dates of the packing list's sales orders, joined with " & ". If an order has no quotation, there is a warning and no value. | Logic B13, sample 90000014 | Q06 |
| A19 | Multi-line text fields are longer than in the FS (samples do not fit). | Samples | Q19 |
| A20 | English texts. Dates DD-MM-YYYY. Weights with 3 decimals. HSN "3917 2390". | Samples | — |
| A21 | `.xlsx` only, first sheet, read with `CL_FDT_XL_SPREADSHEET`. | Target release | Q21 |
| A22 | Re-uploading a saved packing list replaces it and increases the revision. Creation data is kept. | FS silent; "Rev.1" sample | Q17 |
| A24 | Authorization object `ZSD_EPACK` (04 print, 01 save, per company code) + `V_VBRK_VKO` 04. | FS 2.3 empty | Q14 |
| A25 | Output: preview, print, or PDF into a folder (one file per packing list). The default printer comes from the user master. | FS silent | — |
| A26–A28 | Layout as the sample PDFs. Unused notify parties are crossed out with a graphic. Group values print on the first line. | FS + samples | Q20 |
| A29 | S/4HANA, ABAP ≥ 7.54 (same as the other Astral / UDAY developments). | Repository | Q21 |
| A30 | Several telephone / mobile / e-mail entries: the default entry, else the first one found. | Logic B21–B23 | — |
| A31 | Partners via KNVP (Logic B20 / B26) with the sales area of the invoice. Several partners: the default (DEFPA), else lowest counter, with a warning. No partner is an error. | Logic sheet | Q07, Q11 |
| A33 | Several packing lists per file. A row without "Packing list no" belongs to the packing list above. | User request "multiple PL numbers" | Q01 |
| A36 | Missing GSTIN / CIN / IEC / description / HSN gives a warning and an empty field. | FS silent | Q23 |
| A38 | An exact duplicate row is ignored with a warning. The same Sl.No. with different values is an error. | User request "avoid duplicates" | Q03 |
| A39 | Data is saved after a successful output, including preview. Nothing is saved or logged in "Validate only". | "at last add the data in Z table" | Q15 |
| A40 | Uploaded data → `ZSD_EPACK_HDR / INV / DATA` (FS table). Run log → new `ZSD_EPACK_LOG`, one row per packing list. | "Z table that was created for log" is ambiguous | Q16 |

### Confirmed by the repository (no longer assumptions)

* Part No. = SAP material number: Logic B58 "MARC get STEUC (From STEP no 53)", and row 53 is Part No.
* IEC = company address field BUILDING: Logic B11. The Technical detail marks `ZSD_EXPREF` "Not Req – fetched from SAP".
* Terms of shipment = INCO1 & INCO2: Logic B44.

### Withdrawn from version 1

The separate upload transaction (`ZSD_EPACK_UPL`, old A23). The IEC override from the Excel
(old A13). Document partners from VBPA (old A31; the Logic sheet requires KNVP). Contact person, FAX
and tax number (old A15 / A16). Incoterm text from TINCT (old A18). One packing list per file (old A33).
