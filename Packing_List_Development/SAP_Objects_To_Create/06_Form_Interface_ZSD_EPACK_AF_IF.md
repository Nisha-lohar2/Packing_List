# 06 — Form interface `ZSD_EPACK_AF_IF` (SFP → Interface)

**Description:** Export packing list — form interface
**Interface type:** ABAP Dictionary-based interface

### Import parameters

| Parameter | Type assignment | Type | Optional | Pass by value |
|---|---|---|:---:|:---:|
| `IS_HEADER` | TYPE | `ZSD_S_EPACK_AF_HDR` | | ✔ |
| `IT_ITEM` | TYPE | `ZSD_TT_EPACK_AF_ITM` | | ✔ |

Export parameters, exceptions, global data, initialization: none.
Keep the standard system parameters (`/1BCDWB/DOCPARAMS`, `/1BCDWB/FORMOUTPUT`).
