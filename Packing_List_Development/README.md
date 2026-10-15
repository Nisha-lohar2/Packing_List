# Packing List development — WRICEF 102-B "Export consolidated packing list"

Everything needed to build the export packing list in SAP. The FS and its
annexures stay in the repository root, unchanged:
`Object ID- 102(B)Export consolidated packing list.docx`, `Logic_sheet1.xlsx`,
`Z table Format1.xlsx` and the two sample PDFs.

## Current version: `v3/` (revised requirement, two reports)

`v3/README.md` — report 1 `ZSD_EPACK_EXTRACT` (sales orders → editable ALV → Excel in the Z-table format),
report 2 `ZSD_EPACK_UPLOAD` (upload → validation → `ZSD_EPACK_DATA` → Smart Form), DDIC, mappings,
open questions, test cases, activation steps. The folders below are earlier versions.

## Contents

| Folder / file | What it is |
|---|---|
| `docs/Packing_List_Report_Design.md` | **Start here.** Excel and Logic-sheet summary, field mapping, deviations, open questions, data flow, test plan, object list |
| `docs/Development_Assumptions.md` | Every assumption applied (`Axx` in the code) |
| `docs/Packing_List_Requirement_Analysis_and_Development_Approach.md` | Full requirement analysis of the FS, with 43 clarifications |
| `src/DDIC_and_Repository_Objects.md` | **Complete list of Z objects** (§5), data elements, log table `ZSD_EPACK_DATA` (the only Z table, all invoices in one STRING field), message class |
| `src/Adobe_Form_Interface_Structures.md` | Adobe Form `ZSD_EPACK_AF` (the only form): interface `ZSD_EPACK_AF_IF`, header structure `ZSD_S_EPACK_AF_HDR`, item structure `ZSD_S_EPACK_AF_ITM`, table type `ZSD_TT_EPACK_AF_ITM`, layout notes |
| `src/ZSD_EPACK_PRINT.prog.abap` | Report, started with SE38 / SA38 (procedural, no custom classes) |
| `src/ZSD_EPACK_PRINT_TOP.prog.abap` | Include: types, constants, global data |
| `src/ZSD_EPACK_PRINT_SEL.prog.abap` | Include: selection screen and its events |
| `src/ZSD_EPACK_PRINT_F01.prog.abap` | Include: main routines — Excel upload and split by packing list, Logic-sheet data retrieval (bulk), one form per packing list, write to log table `ZSD_EPACK_DATA`, result list |
| `src/ZSD_EPACK_PRINT_F02.prog.abap` | Include: helper routines — package groups, FS weight rule, totals, printed rows, formatting, parsing |
| `templates/ZSD_EPACK_Upload_Template_v2.xlsx` | Upload template with instructions and a worked example (sample 90000081) |
| `test_data/Sample_data_ABAP.xlsx` | Business sample upload file (packing list 90000088) |
| `test_data/Sample_data_ABAP_Expected_Results.md` | Expected result of that file, and data points to check |

## Build order in SAP

1. Package, data elements, table `ZSD_EPACK_DATA`, Adobe structures and table type, as in `src/DDIC_and_Repository_Objects.md` §9.
2. Form interface `ZSD_EPACK_AF_IF` and Adobe Form `ZSD_EPACK_AF` (SFP), as in `src/Adobe_Form_Interface_Structures.md`.
3. Message class `ZSD_EPACK`, then program `ZSD_EPACK_PRINT` with its includes `_TOP`, `_SEL`, `_F01`, `_F02` (SE38).
4. Test with `test_data/Sample_data_ABAP.xlsx` and the test plan in `docs/Packing_List_Report_Design.md` §6.

**Status:** code complete but **not yet compiled or run in an SAP system**.
Check the items marked *SAP-VERIFY* before activation. Open functional
questions: `docs/Packing_List_Report_Design.md` §4.
