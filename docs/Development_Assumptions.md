# WRICEF 102-B Export Packing List — development status and assumptions

The FS (`Object ID- 102(B)Export consolidated packing list.docx`, V.01) is
**not approved** and leaves 43 clarifications open (see
`Packing_List_Requirement_Analysis_and_Development_Approach.md` §11).
Development started anyway, at the user's request. Every gap was filled
with the assumption listed below. Each one is referenced in the source code as `Axx`.
Wherever possible an assumption is isolated in one method or constant, so a
different answer from the functional team is a local change.

## 1. What has been built

| Deliverable | File | Status |
|---|---|---|
| DDIC + repository object specification | `src/DDIC_and_Repository_Objects.md` | Ready to create in SE11/SE91/SU21/SE93 |
| Types / constants | `src/ZIF_SD_EPACK.intf.abap` | Code complete |
| Exception class | `src/ZCX_SD_EPACK.clas.abap` | Code complete |
| Business rules + 26 ABAP Unit tests | `src/ZCL_SD_EPACK_RULES.clas.abap`, `.testclasses.abap` | Code complete |
| Data provider (print) | `src/ZCL_SD_EPACK_DATA.clas.abap` | Code complete |
| Excel upload | `src/ZCL_SD_EPACK_UPLOAD.clas.abap` | Code complete |
| Form output / message popup | `src/ZCL_SD_EPACK_OUTPUT.clas.abap` | Code complete |
| Print report (T-code `ZSD_EPACK`) | `src/ZSD_EPACK_PRINT.prog.abap` | Code complete |
| Upload report (T-code `ZSD_EPACK_UPL`) | `src/ZSD_EPACK_UPLOAD.prog.abap` | Code complete |
| Smart Forms (2) + style + graphic | `docs/ZSD_EPACK_SmartForm_Specification.md` | Build specification (Smart Forms are created in the `SMARTFORMS` transaction) |
| Upload template v2 + example | `templates/ZSD_EPACK_Upload_Template_v2.xlsx` | Ready |

**Not verified:** none of the code has been compiled or run. No SAP system was
available. Before activation, check the table and field names marked
**SAP-VERIFY** in the code and in the DDIC document. The upload rules were
cross-checked against the example sheet with an off-system simulation. It
reproduces the totals of sample 90000081 exactly (77 lines,
4074 / 12146 / 116323.800 / 115386.780, no errors, no warnings).

## 2. Assumptions applied

| ID | Area | Assumption applied | FS / source gap | Where in code | Clarification |
|---|---|---|---|---|---|
| A01 | Data model | One tax invoice belongs to exactly one packing list. The upload rejects an invoice that is already linked to another one. | Not stated | `ZCL_SD_EPACK_UPLOAD->VALIDATE` (msg 026) | CL-02 |
| A02 | Selection | A packing list is printed as a unit. If invoices are entered, every invoice of that packing list must be in the selection. | Not stated | `ZCL_SD_EPACK_DATA->RESOLVE_PACKNO` (msg 006) | CL-02 |
| A03 | Selection | Invoices entered one by one must be linked to a packing list (error 004). Unlinked invoices that only fall inside a From–To range are ignored. A **Packing List No.** field was added as an alternative input. | FS gives only "Tax invoice From/To, multiple selection" | `RESOLVE_PACKNO`, `ZSD_EPACK_PRINT` | CL-02 |
| A04 | Consolidation | All invoices of one packing list must have the same company code, sold-to, ship-to, business place (GSTIN), payment terms and Incoterms, otherwise there is an error. Header data is then taken from the first invoice. | FS logic reads "VBRK" as if there were one invoice | `CHECK_CONSISTENCY`, `READ_PARTNERS` (msg 011) | CL-04, CL-30 |
| A05 | Format | The format is chosen **automatically**: "with notify party" if any notify party is uploaded, otherwise "without". The user may override it. Choosing "without" while notify parties exist is an error. | FS a/b labels contradict each other | `DETERMINE_FORM` (msg 013) | CL-03 |
| A06 | Package groups | Merged Excel cells mean a package group. An empty *Articles* cell belongs to the article group above, and empty *Gross + Net Wt.* cells belong to the weight group above. The two groupings are independent. Values are stored on the group's first line only. | Not described; derived from the samples | `ZCL_SD_EPACK_RULES->BUILD_ITEMS` | CL-08 |
| A07 | Weight check | FS rule GW ≥ NW is checked **per weight group**, which implies the grand total. Equal is allowed. Checked at upload (save refused) **and** again at print (no output, FS message 001). Gross and net must be entered together (msg 028). | Level not stated | `CHECK_WEIGHTS` | CL-10 |
| A08 | Quantities / totals | Quantity = STD × Articles (true for all sample lines) is reported as a **warning** only. *Total Articles* (`ZSD_TART`) is **calculated** as the sum of Articles, not uploaded. | Not stated; FS template has no Total Articles column | `CHECK_QUANTITIES`, `CALC_TOTALS` | CL-40, CL-41 |
| A09 | Package type | New optional template column **AL "Package Type"**, carried forward to the following lines. Subtotal rows ("2342 <=Loose Pipes") appear only when more than one type is used. Total text: "Total N <type>", or "Total a T1 + b T2 = N Packages", or "Total N Packages" when no type is given. | Seen in the samples, absent from the FS | `BUILD_PRINT_ITEMS`, `TOTAL_TEXT` | CL-09 |
| A10 | Notify parties | Notify parties must be filled from 1 upwards, with no gaps, so that crossed-out windows are always the last ones. | Not stated | `MAP_HEADER` (msg 031) | CL-03 |
| A11 | Invoice types | Allowed billing types come from TVARVC `ZSD_EPACK_FKART`. **If the variable is empty, all types are accepted.** Cancelled invoices and cancellation documents are always rejected. | Not stated | `READ_INVOICES`, `VALIDATE` (msgs 009, 010) | CL-29 |
| A12 | Material data | *Part No.* = SAP material number. An uploaded description or HSN **wins**. If it is empty, `MAKT` (EN) or `MARC-STEUC` of the billing plant is used. A part no. not found on the invoices gives a warning only. | FS: "Z table **or** material master" | `RESOLVE_ITEMS`, `VALIDATE` (msgs 030, 044, 048) | CL-07, CL-18 |
| A13 | IEC | Exporter's Ref. = uploaded `ZSD_EXPREF` if filled, otherwise company address field `ADRC-BUILDING` (logic sheet B11). Printed as "IEC:<no>". | Four conflicting sources | `BUILD_HEADER` | CL-05 |
| A14 | Description of goods | Uploaded text (one line per HSN) if filled. Otherwise one line per HSN, "<T604N text> - HS CODE : <hsn>" (clubbed across invoices). The section heading of each HSN is the goods line that contains that HSN. | FS vs Sheet3 conflict | `RESOLVE_ITEMS`, `HEADING_FOR_HSN`, `BUILD_TEXTS` | CL-06 |
| A15 | Party blocks | Party address layout copied from sample 90000081. Contact person "K/A:-" comes from `ADRC-NAME_CO`. When sold-to = ship-to, the "without NP" form prints one **CONSIGNEE** block, as in sample 90000014. Otherwise it prints separate Sold-to and Ship-to blocks (FS). | FS vs sample conflict | `PARTY_LINES`, `SAME_PARTY`, form §5 | CL-13, CL-14 |
| A16 | Ship-to tax no. | "TAX NUMBER :" comes from `KNA1-STCD1` and is printed for the ship-to only, when filled. | Only in sample | `PARTY_LINES` | CL-14 |
| A17 | Order No. & Date | The customer reference and date of the **quotation** the sales order was created from (`VBAK-VGTYP = 'B'`), else of the sales order itself. Several are joined with " & ", e.g. "AL/…/01 & AL/…/02 DATE: 18.08.2025 & 21.08.2025". | Several SOs not covered | `ORDER_REFERENCE` | CL-15 |
| A18 | Payment / Incoterms | Payment term text = `TVZBT-VTEXT` of the first day-limit entry. Terms of shipment = `TINCT` description in upper case + `INCO2`. | Text source not stated | `PAYMENT_LINES` | CL-20 |
| A19 | Field lengths | Notify parties, containers and description of goods are widened to SSTRING 1333. Marks and no. of packages are widened to 255. Declarations are STRING. The samples do not fit the FS lengths, and CHAR 1500 is too long for a table. Other fields keep their FS length; longer values are rejected (msg 022). | FS lengths too short | DDIC §1, `HEADER_MAP` | CL-12 |
| A20 | Formats | Language English. Dates DD-MM-YYYY (order date DD.MM.YYYY, as in the sample). Weights with 3 decimals. Quantity without decimals when whole. Country and region names in upper case. HSN printed as "3917 2390". | Not stated | `ZCL_SD_EPACK_RULES` format methods | CL-20, CL-26 |
| A21 | Excel reading | `.xlsx` only, read with `CL_FDT_XL_SPREADSHEET` from the **first sheet**. Header data may stand in any row. Rows without Sl.No. (e.g. the total row) are ignored. Weights are rounded to 3 decimals. Sl.No., STD and Articles must be whole numbers ≤ 999999. Sl.No. must be unique and ascending. | Upload not specified | `READ_WORKSHEET`, `MAP_*` | CL-11 |
| A22 | Revisions / audit | "Overwrite" **replaces** the packing list completely. The revision number increases by 1 and old versions are not kept. Creator data is kept, and changer and printer (who and when) are recorded. The revision is in the form interface but **not printed**. | Not stated; "Rev.1" in a sample file name | `SAVE`, `MARK_PRINTED` | CL-36, CL-37 |
| A23 | Upload transaction | A separate T-code **`ZSD_EPACK_UPL`** with Test run (default on), Create, Overwrite and Delete. | FS names only `ZSD_EPACK` | `ZSD_EPACK_UPLOAD` | CL-11 |
| A24 | Authorization | New object `ZSD_EPACK` (ACTVT 01/02/04/06 + BUKRS) for upload, delete and print, plus `V_VBRK_VKO` (ACTVT 04) per sales org. at print. | FS §2.3 empty | `CHECK_AUTHORITY`, `VALIDATE`, `DELETE` | CL-22 |
| A25 | Output | Preview (default), print, or PDF download. The device defaults to the user's printer, else LOCL. Print and PDF record a print stamp; preview does not. No e-mail. | Not stated | `ZCL_SD_EPACK_OUTPUT`, `ZSD_EPACK_PRINT` | CL-19 |
| A26 | Crossing out | Unused notify-party windows show a fixed SE78 graphic `ZSD_EPACK_CROSS` (diagonal line). The empty rows below the last item on the last page are **not** crossed out (not possible with a dynamic height); "*** END OF PACKING LIST ***" is printed instead. | FS + sample | Form spec §4, §6 | CL-38 |
| A27 | Merged-cell look | Group values print on the first line of the group, and the inner borders are left open. They are not vertically centred. Groups may break across pages (the sample does too). | Smart Forms limitation | Form spec §6 | CL-08 |
| A28 | Layout | The two sample PDFs are the "approved Packing List format". Smart Forms (not Adobe Forms), as the FS requires. Declaration and signature appear on the last page only, as in the samples. "For Astral Limited" is fixed text. | Approved layout not in the repository | Form spec | CL-20 |
| A29 | System | S/4HANA on-premise, ABAP 7.54 or higher. This matches the other Astral/UDAY developments in `KPMG_Stock_Validation`, which use `+=` and CDS views. Classic tables are read because the FS logic and the India tables (`J_1BBRANCH`, `T001Z`) have no released CDS equivalent. | Release not stated | all classes | CL-31 |
| A30 | Communication data | Telephone = `ADR2` entry with R3_USER 1, mobile = R3_USER 3 (logic sheet), otherwise the default or first number of that kind. E-mail = default or first `ADR6` entry. FAX from `ADR3` (seen in the sample) is added. | Several numbers possible | `READ_ADDRESSES`, `COMM_VALUE` | CL-25 |
| A31 | Partners | Sold-to and ship-to are read from `VBPA` of the billing document: the header partner, otherwise the first item partner. Address from `VBPA-ADRNR`, else `KNA1-ADRNR`. A missing ship-to is an error (msg 042). | Logic sheet uses KNVP defaults | `READ_PARTNERS` | CL-13 |
| A32 | Re-validation | At print, invoices are checked again (cancelled since upload → error) and the weight rule is checked again. | Not stated | `GET_PRINT_DATA` | — |
| A33 | Packing list no. | Free manual number, CHAR 10, with ALPHA conversion (numeric numbers are stored with leading zeros and displayed without). One packing list per Excel file. | FS: manual | Domain `ZSD_PACKNO`, `MAP_HEADER` | CL-16 |
| A34 | Commercial-invoice fields | `ZSD_ADV_AMT`, `ZSD_CMMT` and `ZSD_INV_DECL` are uploaded and stored, but not printed on the packing list. The advance-amount currency is left empty. | In FS template, not in FS table | DDIC §2.1 | CL-17 |
| A35 | Item order | Items print in Sl.No. order. An HSN section heading is printed whenever the HSN changes from the previous line; lines are not re-sorted by HSN. | Not stated | `BUILD_PRINT_ITEMS` | CL-43 |
| A36 | Missing master data | Missing GSTIN, CIN, IEC, address, description or HSN gives a **warning** and an empty field, not an error, so that the document can still be previewed. | Not stated | `BUILD_HEADER`, `RESOLVE_ITEMS` (msgs 044, 048, 049) | CL-27 |
| A37 | Order-varying SO / invoice fields | The FS sentence about "designated fields within … Sales Orders and Invoices" is **not implemented**, because no field is named. All manual data comes from the upload. | Undefined | — | CL-24 |

## 3. Before go-live, these assumptions need explicit sign-off

A01–A05, A06/A07 (package groups and the weight check level), A12–A14 (data
sources), A19 (field lengths), A23/A24 (upload transaction and authorizations),
A26–A28 (layout).
