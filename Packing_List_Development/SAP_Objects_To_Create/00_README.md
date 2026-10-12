# WRICEF 102-B Export Packing List — Z objects to create

Target: SAP system **S2A**, one **local package** (proposal: `$ZSD_EPACK`, no transport).
Create the objects in the order of the files 01 → 09. Activate after each step.

| File | Object(s) | Transaction |
|---|---|---|
| `01_Data_Elements.md` | Data elements `ZSD_*` (only the missing ones) | SE11 |
| `02_Table_ZSD_EPACK_DATA.md` | Table `ZSD_EPACK_DATA` (the only Z table — log) | SE11 |
| `03_Structure_ZSD_S_EPACK_AF_HDR.md` | Structure — Adobe Form header | SE11 |
| `04_Structure_ZSD_S_EPACK_AF_ITM.md` | Structure — Adobe Form item row | SE11 |
| `05_TableType_ZSD_TT_EPACK_AF_ITM.md` | Table type | SE11 |
| `06_Form_Interface_ZSD_EPACK_AF_IF.md` | Adobe Form interface | SFP |
| `07_Adobe_Form_ZSD_EPACK_AF.md` | Adobe Form (layout) | SFP |
| `08_Message_Class_ZSD_EPACK.md` / `.csv` | Message class and its messages | SE91 |
| `09_Program_ZSD_EPACK_PRINT/` | Report + 4 includes + text elements | SE38 |
| `Object_List.csv` | All objects as a checklist | — |

## Rules applied

* **Only Z objects** are created. Standard objects (`MANDT`, `MEINS`, `GEWEI`, `WAERS`, `XFELD`,
  domain `VBELN` …) are only referenced.
* **No domains.** Data elements use predefined types; `ZSD_PACKNO` uses the standard domain
  `VBELN` (CHAR 10 with conversion `ALPHA`), so no Z domain is needed.
* **Reuse first.** Before creating a data element, search in SE11 (F4 on "Data type", or
  table `DD04L`) for an existing `Z*` data element with the **same type, length and decimals**.
  If one exists, do not create the new one — enter the existing name in the table / structure
  instead (and keep the field name). `01_Data_Elements.md` has a "reuse check" column.
* **Package, message class:** create only if they do not exist yet. If a message class for
  SD packing lists already exists, the messages of file 08 can be added there instead — then
  change `MESSAGE-ID zsd_epack` / `(zsd_epack)` / `gc_msgid` in the program to that name.
* **Not created for now:** transaction code, lock object, custom authorization object,
  TVARVC entry, Smart Forms.

## After creation

1. Activate the program; fix anything marked `SAP-VERIFY` in the source comments
   (`J_1BBRANCH-GSTIN`, `T604N`, `TVZBT-ZTAGG`, `KNVP-DEFPA`, `ADR2-R3_USER`).
2. Run `ZSD_EPACK_PRINT` with `test_data/Sample_data_ABAP.xlsx` → expected result in
   `test_data/Sample_data_ABAP_Expected_Results.md`.

The code has **not** been compiled in an SAP system yet.
