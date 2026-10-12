# 04 — Structure `ZSD_S_EPACK_AF_ITM` (SE11)

**Short description:** Export packing list — Adobe Form item row
**Enhancement category:** Can be enhanced (character-type or numeric)

One entry = one printed row of the item table: `ROW_TYPE` `H` section heading, `I` line item,
`S` package-type subtotal. Details: `src/Adobe_Form_Interface_Structures.md` §3.

| Component | Typing method | Component type / built-in type | Length | Dec | Description / column |
|---|---|---|---:|---:|---|
| `ROW_TYPE` | Predefined type | `CHAR` | 1 | 0 | — (row layout) |
| `ROW_TEXT` | Predefined type | `CHAR` | 132 | 0 | heading text (`H`) / subtotal label (`S`) |
| `SL_NO` | Predefined type | `CHAR` | 6 | 0 | Sl.No. |
| `INV_SN` | Predefined type | `CHAR` | 6 | 0 | Sr.No. INV |
| `ARTICLE_NO` | Predefined type | `CHAR` | 15 | 0 | Article No. |
| `STD` | Predefined type | `CHAR` | 6 | 0 | STD |
| `ARTICLES` | Predefined type | `CHAR` | 10 | 0 | Articles |
| `PART_NO` | Predefined type | `CHAR` | 40 | 0 | Part No. |
| `PROD_DESC` | Predefined type | `CHAR` | 40 | 0 | Product Description |
| `QUANTITY` | Predefined type | `CHAR` | 17 | 0 | Quantity Pcs |
| `GROSS_WT` | Predefined type | `CHAR` | 17 | 0 | Gross Wt. KGS |
| `NET_WT` | Predefined type | `CHAR` | 17 | 0 | Net Wt. KGS |
| `HSN` | Predefined type | `CHAR` | 10 | 0 | H.S Code |
| `PKG_TYPE` | Predefined type | `CHAR` | 20 | 0 | — |
| `ART_GRP` | Predefined type | `NUMC` | 4 | 0 | — |
| `ART_FIRST` | Data element | `XFELD` | | | — |
| `ART_LAST` | Data element | `XFELD` | | | — |
| `ART_ROWS` | Predefined type | `INT2` | 5 | | — |
| `WT_GRP` | Predefined type | `NUMC` | 4 | 0 | — |
| `WT_FIRST` | Data element | `XFELD` | | | — |
| `WT_LAST` | Data element | `XFELD` | | | — |
| `WT_ROWS` | Predefined type | `INT2` | 5 | | — |

**Total:** 22 components.
