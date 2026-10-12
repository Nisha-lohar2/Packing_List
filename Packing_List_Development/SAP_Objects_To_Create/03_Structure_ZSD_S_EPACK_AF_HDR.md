# 03 — Structure `ZSD_S_EPACK_AF_HDR` (SE11)

**Short description:** Export packing list — Adobe Form header
**Enhancement category:** Can be enhanced (character-type or numeric)

Values are passed ready to print (character fields). Multi-line blocks are `STRING` with line
breaks. Example values and the source of every field: `src/Adobe_Form_Interface_Structures.md` §2.

| Component | Typing method | Component type / built-in type | Length | Dec | Description / label |
|---|---|---|---:|---:|---|
| `PACKNO` | Data element | `ZSD_VBELN` | | | Packing List No. |
| `PACKDT` | Predefined type | `CHAR` | 10 | 0 | Date |
| `FORMAT` | Predefined type | `CHAR` | 1 | 0 | — (controls layout) |
| `NP_COUNT` | Predefined type | `INT1` | 3 | | — |
| `EXP_NAME` | Predefined type | `CHAR` | 80 | 0 | EXPORTER : (1st line) |
| `EXP_ADDRESS` | Predefined type | `STRING` | | | address lines |
| `EXP_GSTIN` | Predefined type | `CHAR` | 18 | 0 | GST NO. (exporter box) and GST No. (right box) |
| `EXP_IEC` | Predefined type | `CHAR` | 20 | 0 | Exporter's Ref. → `IEC:` + value |
| `EXP_CIN` | Predefined type | `CHAR` | 21 | 0 | CIN No. |
| `SIGN_FOR` | Predefined type | `CHAR` | 80 | 0 | Signature & Date → `For` + value |
| `SAME_PARTY` | Data element | `XFELD` | | | — |
| `CNS_NAME` | Predefined type | `CHAR` | 80 | 0 | CONSIGNEE : (1st line) |
| `CNS_ADDRESS` | Predefined type | `STRING` | | | address lines |
| `CNS_CONTACT` | Predefined type | `CHAR` | 40 | 0 | `K/A:-` |
| `CNS_TEL` | Predefined type | `CHAR` | 30 | 0 | `TEL.:-` |
| `CNS_FAX` | Predefined type | `CHAR` | 30 | 0 | `FAX :` |
| `CNS_MOBILE` | Predefined type | `CHAR` | 30 | 0 | `MO :` |
| `CNS_EMAIL` | Predefined type | `CHAR` | 241 | 0 | `E-MAIL:` |
| `BUY_NAME` | Predefined type | `CHAR` | 80 | 0 | BUYER : (1st line) |
| `BUY_ADDRESS` | Predefined type | `STRING` | | | address lines |
| `BUY_TEL` | Predefined type | `CHAR` | 30 | 0 | `TEL.:-` |
| `BUY_MOBILE` | Predefined type | `CHAR` | 30 | 0 | `MO :` |
| `BUY_EMAIL` | Predefined type | `CHAR` | 241 | 0 | `E-MAIL:` |
| `NP1` | Predefined type | `STRING` | | | Notify party 1 |
| `NP2` | Predefined type | `STRING` | | | Notify party 2 |
| `NP3` | Predefined type | `STRING` | | | Notify party 3 |
| `NP4` | Predefined type | `STRING` | | | Notify party 4 |
| `NP5` | Predefined type | `STRING` | | | Notify party 5 |
| `ORDER_REF` | Predefined type | `STRING` | | | Order No. & Date |
| `BUYER_REF` | Predefined type | `STRING` | | | Buyer's Reference No. & Date |
| `CNTY_ORGN` | Predefined type | `CHAR` | 20 | 0 | Country of Origin of Goods |
| `CNTY_DEST` | Predefined type | `CHAR` | 50 | 0 | Country of final Destination |
| `PRE_CARR` | Predefined type | `CHAR` | 30 | 0 | Pre-Carriage |
| `REC_PLACE` | Predefined type | `CHAR` | 30 | 0 | Place of Receipt by Pre-carrier |
| `VESSEL` | Predefined type | `CHAR` | 30 | 0 | Vessel / Flight No. |
| `POL` | Predefined type | `CHAR` | 30 | 0 | Port of Loading |
| `POD` | Predefined type | `CHAR` | 30 | 0 | Port of Discharge |
| `PLD` | Predefined type | `CHAR` | 30 | 0 | Place of Delivery |
| `PAY_TERMS` | Predefined type | `STRING` | | | Terms of Payment |
| `SHIP_TERMS` | Predefined type | `CHAR` | 60 | 0 | `TERMS OF SHIPMENT :-` + value |
| `MARKS` | Predefined type | `STRING` | | | Marks & Nos. |
| `CONTAINERS` | Predefined type | `STRING` | | | Container No. |
| `NO_PKGS` | Predefined type | `STRING` | | | No.of Pkgs. |
| `GOODS_DESC` | Predefined type | `STRING` | | | Description of Goods |
| `TOT_ART` | Predefined type | `CHAR` | 10 | 0 | Articles |
| `TOT_TEXT` | Predefined type | `CHAR` | 132 | 0 | Part No. + Product Description |
| `TOT_QTY` | Predefined type | `CHAR` | 17 | 0 | Quantity Pcs |
| `TOT_GWT` | Predefined type | `CHAR` | 17 | 0 | Gross Wt. KGS |
| `TOT_NWT` | Predefined type | `CHAR` | 17 | 0 | Net Wt. KGS |
| `DECLARATION` | Predefined type | `STRING` | | | Declaration : |

**Total:** 50 components. Only `ZSD_VBELN` (existing) and the standard `XFELD` are data elements.
