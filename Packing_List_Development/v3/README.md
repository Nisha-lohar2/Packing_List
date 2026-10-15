# Version 3 — two reports (revised requirement)

Revised FS + revised `Logic_sheet1.xlsx` (copy: `input/Logic_sheet1_revised.xlsx`).
**Report 1** reads the SAP data of sales orders, shows it in an editable ALV and downloads an Excel file in
the Z-table format. **Report 2** uploads that file, validates it, saves it to `ZSD_EPACK_DATA` and prints
the Smart Form.

**Status:** written and checked offline (balanced blocks, every PERFORM has a FORM, no SELECT in a loop).
**Not yet compiled or run in S2A** — check the *SAP-VERIFY* points.

## Files (file change summary)

| File | New / changed | Content |
|---|---|---|
| `ZSD_EPACK_COMMON.prog.abap` | new | Include: Excel layout (46 columns of sheet "Z table"), conversions Excel ↔ field, address helpers, message popup, F4 |
| `ZSD_EPACK_EXTRACT.prog.abap` | new | Report 1 — sales orders → editable ALV → Excel download |
| `ZSD_EPACK_UPLOAD.prog.abap` | new | Report 2 — upload → validation → save → Smart Form |
| `DDIC_and_SmartForm.md` | new | Data elements, table `ZSD_EPACK_DATA`, lock object, Smart Form structures / interface / layout |
| `Field_Mapping.md` | new | Excel column ↔ table field ↔ SAP logic ↔ Smart Form |
| `Validation_Mapping.md` | new | Validations V01 – V25 with source, level, message |
| `Analysis_and_Open_Points.md` | new | File analysis, approach, objects, assumptions A01 – A20, open questions Q01 – Q13 |
| `Test_Cases.md` | new | Test cases report 1 and 2 |
| `input/Logic_sheet1_revised.xlsx` | new | Revised logic sheet (source) |

Unchanged: `v2/`, `src/`, `src_step1/`, `SAP_Objects_To_Create/`, `docs/` (earlier invoice / Adobe Form versions).
Note: `SAP_Objects_To_Create/02_Table_ZSD_EPACK_DATA.md` describes the **old** table; the valid definition is
`v3/DDIC_and_SmartForm.md` §3.

## Activation steps

1. **Package / transport** for the objects below.
2. **Data elements** `ZDE_EPACK_*` (SE11) — `DDIC_and_SmartForm.md` §2. Activate.
3. **Table** `ZSD_EPACK_DATA` (SE11) — §3: fields, key `MANDT, ZSD_PACKNO, VBAK_VBELN, ZSD_SN`,
   reference fields (`ZSD_QTY` → `ZSD_MEINS`, `ZSD_GWT`/`ZSD_NWT` → `ZSD_GEWEI`), technical settings,
   enhancement category. Activate; SE14 if needed. *SAP-VERIFY* table width (§3 note).
4. **Lock object** `EZSD_EPACK` (SE11) — §4, parameters `MANDT`, `ZSD_PACKNO`. Activate (generates
   `ENQUEUE_/DEQUEUE_EZSD_EPACK`).
5. **Structures** `ZSD_S_EPACK_SF_HDR`, `ZSD_S_EPACK_SF_ITM` (reference fields for the QUAN components),
   **table type** `ZSD_TT_EPACK_SF_ITM` — §5. Activate.
6. **Smart Forms** `ZSD_EPACK_SF` and `ZSD_EPACK_SF_NP` (SMARTFORMS) — interface §6
   (`IS_HEADER`, `IT_ITEMS`), layout as the sample PDFs. Activate (generates the FMs).
7. **Include** `ZSD_EPACK_COMMON` (SE38, type *I – Include program*). Paste, activate together with step 8/9.
8. **Report 1** `ZSD_EPACK_EXTRACT` (SE38, type 1):
   * paste the source; text symbol `B01`, selection text `S_VBELN` (Dictionary reference) — see end of source;
   * **GUI status `ZEPACK_STD`**: SE41 → *Copy status* from program `SAPLKKBL`, status `STANDARD_FULLSCREEN`
     to program `ZSD_EPACK_EXTRACT`, status `ZEPACK_STD`; add an application-toolbar button with function
     code `ZDOWNLOAD`, text "Download Excel", icon `ICON_EXPORT`; activate;
   * activate the program.
9. **Report 2** `ZSD_EPACK_UPLOAD` (SE38, type 1): paste, text symbols `B01`–`B03` and selection texts
   (end of source), activate.
10. **T-codes** (SE93, report transaction, start with selection screen): `ZSD_EPACK` → `ZSD_EPACK_EXTRACT`;
    `ZSD_EPACK_UPL` → `ZSD_EPACK_UPLOAD` (name to confirm, Q02). GUI support SAP GUI for Windows.
11. **SAP-VERIFY** in S2A: `T001W-J_1BBRANCH` and `J_1BBRANCH-GSTIN` exist and are filled (A10);
    `T001Z` party `ZCIN`; `ADR2-R3_USER` values 1 / 3; LIPS batch-split handling (Q08);
    `CL_SALV_TABLE->TO_XML` XLSX available (SAP_BASIS ≥ 7.40).
12. **Test** with `Test_Cases.md`; ATC check; SQL trace.
13. Confirm the open questions in `Analysis_and_Open_Points.md` §5.
