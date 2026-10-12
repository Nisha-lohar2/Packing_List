# 09 — Program `ZSD_EPACK_PRINT` (SE38)

| Object | SE38 attributes | Source file |
|---|---|---|
| `ZSD_EPACK_PRINT` | Type **1 Executable program**, title "Export consolidated packing list", status Customer production program, application S (SD) | `ZSD_EPACK_PRINT.abap` |
| `ZSD_EPACK_PRINT_TOP` | Type **I Include** | `ZSD_EPACK_PRINT_TOP.abap` |
| `ZSD_EPACK_PRINT_SEL` | Type **I Include** | `ZSD_EPACK_PRINT_SEL.abap` |
| `ZSD_EPACK_PRINT_F01` | Type **I Include** | `ZSD_EPACK_PRINT_F01.abap` |
| `ZSD_EPACK_PRINT_F02` | Type **I Include** | `ZSD_EPACK_PRINT_F02.abap` |

Order: create the main program, paste its source, then double-click each `INCLUDE` line to create
the include and paste its source. Activate all together (they reference files 01–08).

### Text elements (Goto → Text elements)

| Type | Key | Text |
|---|---|---|
| Text symbol | `B01` | Excel file |
| Selection text | `P_FILE` | Excel file (.xlsx) |

Selection screen: only the Excel file (mandatory). Output: print preview of each packing list on
device `LP01` (constant `GC_DEST` in `ZSD_EPACK_PRINT_TOP`); the user prints from the preview.

Run with SE38 / SA38 (no transaction code for now).
