# Packing List development — WRICEF 102-B "Export consolidated packing list"

Everything needed to build the export packing list in SAP. The FS and its
annexures stay in the repository root, unchanged:
`Object ID- 102(B)Export consolidated packing list.docx`, `Logic_sheet1.xlsx`,
`Z table Format1.xlsx` and the two sample PDFs.

## Contents

| Folder / file | What it is |
|---|---|
| `docs/Packing_List_Report_Design.md` | **Start here.** Excel and Logic-sheet summary, field mapping, deviations, open questions, data flow, test plan, object list |
| `docs/Development_Assumptions.md` | Every assumption applied (`Axx` in the code) |
| `docs/Packing_List_Requirement_Analysis_and_Development_Approach.md` | Full requirement analysis of the FS, with 43 clarifications |
| `docs/ZSD_EPACK_SmartForm_Specification.md` | Build instructions for the two Smart Forms, the style and the graphic |
| `src/DDIC_and_Repository_Objects.md` | Log table `ZSD_EPACK_DATA` (the only Z table, FS layout, all invoices in one STRING field), form structures, lock object, message class, authorization object, transaction |
| `src/Adobe_Form_Interface_Structures.md` | Adobe Form interface `ZSD_EPACK_AF_IF`: header structure `ZSD_S_EPACK_AF_HDR`, item structure `ZSD_S_EPACK_AF_ITM`, table type `ZSD_TT_EPACK_AF_ITM` (layout of sample 90000014) |
| `src/ZSD_EPACK_PRINT.prog.abap` | Report behind transaction `ZSD_EPACK` (procedural, no custom classes) |
| `src/ZSD_EPACK_PRINT_TOP.prog.abap` | Include: types, constants, global data |
| `src/ZSD_EPACK_PRINT_SEL.prog.abap` | Include: selection screen and its events |
| `src/ZSD_EPACK_PRINT_F01.prog.abap` | Include: main routines — Excel upload and split by packing list, Logic-sheet data retrieval (bulk), one form per packing list, write to log table `ZSD_EPACK_DATA`, result list |
| `src/ZSD_EPACK_PRINT_F02.prog.abap` | Include: helper routines — package groups, FS weight rule, totals, printed rows, formatting, parsing |
| `templates/ZSD_EPACK_Upload_Template_v2.xlsx` | Upload template with instructions and a worked example (sample 90000081) |
| `test_data/Sample_data_ABAP.xlsx` | Business sample upload file (packing list 90000088) |
| `test_data/Sample_data_ABAP_Expected_Results.md` | Expected result of that file, and data points to check |

## Build order in SAP

1. DDIC objects, lock object, message class and authorization object, as in `src/DDIC_and_Repository_Objects.md` §10.
2. Smart Style, SE78 graphic and both Smart Forms, as in `docs/ZSD_EPACK_SmartForm_Specification.md`.
3. Program `ZSD_EPACK_PRINT` with its includes `_TOP`, `_SEL`, `_F01`, `_F02` (SE38), then transaction `ZSD_EPACK` (SE93).
4. Test with `test_data/Sample_data_ABAP.xlsx` and the test plan in `docs/Packing_List_Report_Design.md` §6.

**Status:** code complete but **not yet compiled or run in an SAP system**.
Check the items marked *SAP-VERIFY* before activation. Open functional
questions: `docs/Packing_List_Report_Design.md` §4.
