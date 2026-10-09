"! <p class="shorttext synchronized">Export Packing List - types and constants</p>
"!
"! WRICEF 102-B "Export consolidated packing list" (Astral / UDAY / SD).
"!
"! Central type pool and constants. Types that cross the Smart Form
"! interface or are persisted are DDIC objects (ZSD_EPACK_*, ZSD_S_/TT_*);
"! everything that only lives inside the ABAP classes is declared here.
"!
"! Every value marked "ASSUMPTION Axx" below is a decision taken without
"! a confirmed requirement. See docs/Development_Assumptions.md.
INTERFACE zif_sd_epack
  PUBLIC.

  CONSTANTS gc_msgid TYPE symsgid VALUE 'ZSD_EPACK'.

  "! Output / text language. ASSUMPTION A20: always English.
  CONSTANTS gc_langu TYPE spras VALUE 'E'.

  "! Country for HSN (control code) texts in T604N - FS logic sheet B42
  CONSTANTS gc_hsn_country TYPE land1 VALUE 'IN'.

  "! T001Z parameter type holding the CIN - FS logic sheet B7
  CONSTANTS gc_cin_party TYPE party VALUE 'ZCIN'.

  "! TVARVC select-option with the billing types accepted as export tax
  "! invoices. ASSUMPTION A11: if the variable is not maintained, every
  "! billing type is accepted.
  CONSTANTS gc_tvarv_fkart TYPE rvari_vnam VALUE 'ZSD_EPACK_FKART'.

  "! Custom authorization object (fields ACTVT, BUKRS) - ASSUMPTION A24
  CONSTANTS gc_auth_object TYPE xuobject VALUE 'ZSD_EPACK'.

  "! Partner functions - INTERNAL codes. The FS logic sheet writes the
  "! English external codes SP / SH; the database stores AG / WE.
  CONSTANTS:
    BEGIN OF gc_parvw,
      sold_to TYPE parvw VALUE 'AG',
      ship_to TYPE parvw VALUE 'WE',
    END OF gc_parvw.

  CONSTANTS:
    BEGIN OF gc_actvt,
      create  TYPE activ_auth VALUE '01',
      change  TYPE activ_auth VALUE '02',
      display TYPE activ_auth VALUE '03',
      print   TYPE activ_auth VALUE '04',
      delete  TYPE activ_auth VALUE '06',
    END OF gc_actvt.

  "! Units for the QUAN fields. FS columns are "Pcs" and "KGS".
  CONSTANTS:
    BEGIN OF gc_uom,
      quantity TYPE meins VALUE 'ST',
      weight   TYPE gewei VALUE 'KG',
    END OF gc_uom.

  "! Smart Forms - FS: "Two smart form formats ... with notifier party
  "! and without notifier party"
  CONSTANTS:
    BEGIN OF gc_form,
      with_np    TYPE tdsfname VALUE 'ZSD_EPACK_WITH_NP',
      without_np TYPE tdsfname VALUE 'ZSD_EPACK_WITHOUT_NP',
    END OF gc_form.

  "! Format choice on the selection screen
  CONSTANTS:
    BEGIN OF gc_format,
      auto       TYPE char1 VALUE 'A',
      with_np    TYPE char1 VALUE 'W',
      without_np TYPE char1 VALUE 'O',
    END OF gc_format.

  CONSTANTS:
    BEGIN OF gc_output,
      preview TYPE char1 VALUE 'V',
      print   TYPE char1 VALUE 'P',
      pdf     TYPE char1 VALUE 'F',
    END OF gc_output.

  CONSTANTS:
    BEGIN OF gc_upload_mode,
      create    TYPE char1 VALUE 'C',
      overwrite TYPE char1 VALUE 'O',
      delete    TYPE char1 VALUE 'D',
    END OF gc_upload_mode.

  "! Row types of the printed item table (ZSD_S_EPACK_PRT_ITM-ROW_TYPE)
  CONSTANTS:
    BEGIN OF gc_row_type,
      "! HSN section heading, e.g. "PVC PIPES (LEAD FREE) - HS CODE : 3917 2390"
      heading  TYPE zsd_epack_row_type VALUE 'H',
      item     TYPE zsd_epack_row_type VALUE 'I',
      "! Package type subtotal, e.g. "2342 <=Loose Pipes"
      subtotal TYPE zsd_epack_row_type VALUE 'S',
    END OF gc_row_type.

  "! Text blocks passed to the forms (ZSD_S_EPACK_PRT_TXT-BLOCK)
  CONSTANTS:
    BEGIN OF gc_block,
      exporter    TYPE zsd_epack_block VALUE 'EXPORTER',
      sold_to     TYPE zsd_epack_block VALUE 'SOLDTO',
      ship_to     TYPE zsd_epack_block VALUE 'SHIPTO',
      np1         TYPE zsd_epack_block VALUE 'NP1',
      np2         TYPE zsd_epack_block VALUE 'NP2',
      np3         TYPE zsd_epack_block VALUE 'NP3',
      np4         TYPE zsd_epack_block VALUE 'NP4',
      np5         TYPE zsd_epack_block VALUE 'NP5',
      order_ref   TYPE zsd_epack_block VALUE 'ORDERREF',
      buyer_ref   TYPE zsd_epack_block VALUE 'BUYERREF',
      payment     TYPE zsd_epack_block VALUE 'PAYMENT',
      marks       TYPE zsd_epack_block VALUE 'MARKS',
      container   TYPE zsd_epack_block VALUE 'CONTAINER',
      packages    TYPE zsd_epack_block VALUE 'PACKAGES',
      goods       TYPE zsd_epack_block VALUE 'GOODS',
      declaration TYPE zsd_epack_block VALUE 'DECL',
    END OF gc_block.

  "! Maximum number of notify parties - FS: "up to five notifier parties"
  CONSTANTS gc_max_np TYPE i VALUE 5.

  TYPES tr_vbeln TYPE RANGE OF vbeln_vf.
  TYPES tt_vbeln TYPE SORTED TABLE OF vbeln_vf WITH UNIQUE KEY table_line.
  TYPES tt_item  TYPE STANDARD TABLE OF zsd_epack_data WITH DEFAULT KEY.
  TYPES tt_text_line TYPE STANDARD TABLE OF string WITH EMPTY KEY.

  "! One packing list as stored in the three Z tables
  TYPES:
    BEGIN OF ts_packing_list,
      header   TYPE zsd_epack_hdr,
      invoices TYPE tt_vbeln,
      items    TYPE tt_item,
    END OF ts_packing_list.

  "! Everything handed to the Smart Form
  TYPES:
    BEGIN OF ts_print_data,
      form   TYPE tdsfname,
      header TYPE zsd_s_epack_prt_hdr,
      texts  TYPE zsd_tt_epack_prt_txt,
      items  TYPE zsd_tt_epack_prt_itm,
    END OF ts_print_data.

  "! One Excel line item before grouping. The *_filled flags keep the
  "! difference between "cell empty" (merged cell) and "cell = 0",
  "! which the NUMC / QUAN target fields cannot express.
  TYPES:
    BEGIN OF ts_upload_line,
      row        TYPE i,
      sn         TYPE zsd_sn,
      inv_sn     TYPE zsd_inv_sn,
      artno      TYPE zsd_artno,
      std        TYPE zsd_std,
      art        TYPE zsd_art,
      art_filled TYPE abap_bool,
      partno     TYPE zsd_partno,
      matdesc    TYPE zsd_matdesc,
      qty        TYPE zsd_qty,
      gwt        TYPE zsd_gwt,
      gwt_filled TYPE abap_bool,
      nwt        TYPE zsd_nwt,
      nwt_filled TYPE abap_bool,
      hsn        TYPE zsd_hsn,
      pkg_type   TYPE zsd_pkg_type,
    END OF ts_upload_line,
    tt_upload_line TYPE STANDARD TABLE OF ts_upload_line WITH EMPTY KEY.

  "! Totals of one packing list
  TYPES:
    BEGIN OF ts_pkg_total,
      pkg_type TYPE zsd_pkg_type,
      articles TYPE i,
    END OF ts_pkg_total,
    tt_pkg_total TYPE STANDARD TABLE OF ts_pkg_total WITH EMPTY KEY,
    BEGIN OF ts_totals,
      articles  TYPE i,
      quantity  TYPE zsd_qty,
      gross     TYPE zsd_gwt,
      net       TYPE zsd_nwt,
      "! In order of first appearance
      pkg_types TYPE tt_pkg_total,
    END OF ts_totals.

  "! HSN code -> section heading text
  TYPES:
    BEGIN OF ts_hsn_heading,
      hsn  TYPE zsd_hsn,
      text TYPE string,
    END OF ts_hsn_heading,
    tt_hsn_heading TYPE STANDARD TABLE OF ts_hsn_heading WITH EMPTY KEY.

ENDINTERFACE.
