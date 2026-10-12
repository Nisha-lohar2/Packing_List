*&---------------------------------------------------------------------*
*& Include  ZSD_EPACK_PRINT_TOP
*&---------------------------------------------------------------------*
*& WRICEF 102-B "Export consolidated packing list" - global types,
*& constants and data. Procedural version: no custom classes or
*& interfaces. Assumption references (Axx) point to
*& docs/Development_Assumptions.md.
*&---------------------------------------------------------------------*

*----------------------------------------------------------------------*
* Constants
*----------------------------------------------------------------------*
CONSTANTS:
  gc_msgid       TYPE symsgid    VALUE 'ZSD_EPACK',
  " Output / text language (A20)
  gc_langu       TYPE spras      VALUE 'E',
  " Country for HSN texts in T604N - logic sheet B42
  gc_hsn_country TYPE land1      VALUE 'IN',
  " T001Z parameter type holding the CIN - logic sheet B7
  gc_cin_party   TYPE party      VALUE 'ZCIN',
  " Partner functions - INTERNAL codes of SP / SH (logic B20 / B26)
  gc_parvw_sold  TYPE parvw      VALUE 'AG',
  gc_parvw_ship  TYPE parvw      VALUE 'WE',
  " Display / print authorization for billing documents (V_VBRK_VKO)
  gc_actvt_print TYPE activ_auth VALUE '04',
  gc_uom_qty     TYPE meins      VALUE 'ST',
  gc_uom_weight  TYPE gewei      VALUE 'KG',
  " Adobe Form (SFP) - one form; the notify-party section is shown only
  " when notify parties are maintained (IS_HEADER-FORMAT / NP_COUNT)
  gc_form_name   TYPE fpname     VALUE 'ZSD_EPACK_AF',
  gc_fmt_with_np TYPE char1      VALUE 'N',
  gc_fmt_wout_np TYPE char1      VALUE 'O',
  " Output device for the form (print preview, printable from there)
  gc_dest        TYPE rspopname  VALUE 'LP01',
  " Processing status per packing list (result list)
  gc_st_success  TYPE char1      VALUE 'S',
  gc_st_warning  TYPE char1      VALUE 'W',
  gc_st_error    TYPE char1      VALUE 'E',
  " Printed row types
  gc_row_heading TYPE char1      VALUE 'H',
  gc_row_item    TYPE char1      VALUE 'I',
  gc_row_subtot  TYPE char1      VALUE 'S',
  " Notify parties - FS: "up to five"
  gc_max_np      TYPE i          VALUE 5.

*----------------------------------------------------------------------*
* Types - packing list data
*----------------------------------------------------------------------*
TYPES:
  ty_packno_range TYPE RANGE OF zsd_vbeln,
  ty_vbeln_tab   TYPE SORTED TABLE OF vbeln_vf WITH UNIQUE KEY table_line,
  ty_text_tab    TYPE STANDARD TABLE OF string WITH EMPTY KEY.

TYPES:
  " Header data of one packing list - FS "Technical detail" fields
  " (logic sheet "Pass ZSD_PACKNO get ...")
  BEGIN OF ty_header,
    zsd_packno    TYPE zsd_vbeln,
    zsd_packdt    TYPE zsd_date,
    zsd_expref    TYPE zsd_uid,
    zsd_party_ref TYPE zref,
    zsd_ref_dt    TYPE zsd_date,
    zsd_cnty_orgn TYPE zsd_uid,
    zsd_np1       TYPE zsd_np,
    zsd_np2       TYPE zsd_np,
    zsd_np3       TYPE zsd_np,
    zsd_np4       TYPE zsd_np,
    zsd_np5       TYPE zsd_np,
    zsd_prec      TYPE zchar30,
    zsd_rec_pc    TYPE zchar30,
    zsd_pol       TYPE zchar30,
    zsd_pod       TYPE zchar30,
    zsd_pld       TYPE zchar30,
    zsd_mark      TYPE zsd_mark,
    zsd_vsl_flt   TYPE zchar30,
    zsd_cont      TYPE zsd_cont,
    zsd_nopack    TYPE zsd_nopack,
    zsd_hsn_desc  TYPE zsd_hsn_desc,
    zsd_pl_decl   TYPE z_string,
    zsd_adv_amt   TYPE zde_amt1,
    zsd_cmmt      TYPE zsd_cmmt,
    zsd_inv_decl  TYPE z_string,
  END OF ty_header,

  " Line item of one packing list - FS "Technical detail" item fields
  " plus the package groups needed for printing (A06, A09; not stored)
  BEGIN OF ty_item,
    zsd_packno   TYPE zsd_vbeln,
    zsd_sn       TYPE zserno,
    zsd_inv_sn   TYPE zserno,
    zsd_artno    TYPE zchar15,
    zsd_std      TYPE zsqno,
    zsd_art      TYPE zsqno,
    zsd_partno   TYPE zsd_debezei40,
    zsd_matdesc  TYPE zsd_debezei40,
    zsd_qty      TYPE zqty,
    zsd_gwt      TYPE zde_grwt,
    zsd_nwt      TYPE zde_net_wt,
    zsd_hsn      TYPE zna_segment,
    zsd_art_grp  TYPE i,
    zsd_wt_grp   TYPE i,
    zsd_pkg_type TYPE char20,
  END OF ty_item,
  ty_item_tab TYPE STANDARD TABLE OF ty_item WITH DEFAULT KEY.

TYPES:
  " One packing list: header, its invoices, its items and the messages
  " raised for exactly this packing list. Header and items are always
  " kept together under the packing list number.
  BEGIN OF ty_packing_list,
    packno   TYPE zsd_vbeln,
    header   TYPE ty_header,
    invoices TYPE ty_vbeln_tab,
    items    TYPE ty_item_tab,
    rows     TYPE string,          " Excel rows of this packing list
    messages TYPE bapiret2_t,
  END OF ty_packing_list,
  ty_packing_list_tab TYPE SORTED TABLE OF ty_packing_list WITH UNIQUE KEY packno.

TYPES:
  " Everything handed to the Adobe Form (interface ZSD_EPACK_AF_IF)
  BEGIN OF ty_print_data,
    form   TYPE fpname,
    header TYPE zsd_s_epack_af_hdr,
    items  TYPE zsd_tt_epack_af_itm,
  END OF ty_print_data,

  " Processing result of one packing list (one form, one log row)
  BEGIN OF ty_result,
    packno     TYPE zsd_vbeln,
    bukrs      TYPE bukrs,
    status     TYPE char1,
    print      TYPE ty_print_data,
    messages   TYPE bapiret2_t,
    inv_count  TYPE i,
    item_count TYPE i,
    output     TYPE abap_bool,     " form produced
    saved      TYPE abap_bool,     " written to the Z tables
  END OF ty_result,
  ty_result_tab TYPE SORTED TABLE OF ty_result WITH UNIQUE KEY packno.

TYPES:
  " One Excel line item before grouping. The *_filled flags keep the
  " difference between "cell empty" (merged / repeated) and "cell = 0".
  BEGIN OF ty_upload_line,
    row        TYPE i,
    sn         TYPE zserno,
    inv_sn     TYPE zserno,
    artno      TYPE zchar15,
    std        TYPE zsqno,
    art        TYPE zsqno,
    art_filled TYPE abap_bool,
    partno     TYPE zsd_debezei40,
    matdesc    TYPE zsd_debezei40,
    qty        TYPE zqty,
    gwt        TYPE zde_grwt,
    gwt_filled TYPE abap_bool,
    nwt        TYPE zde_net_wt,
    nwt_filled TYPE abap_bool,
    hsn        TYPE zna_segment,
    pkg_type   TYPE char20,
  END OF ty_upload_line,
  ty_upload_line_tab TYPE STANDARD TABLE OF ty_upload_line WITH EMPTY KEY.

TYPES:
  BEGIN OF ty_pkg_total,
    pkg_type TYPE char20,
    articles TYPE i,
  END OF ty_pkg_total,
  ty_pkg_total_tab TYPE STANDARD TABLE OF ty_pkg_total WITH EMPTY KEY,

  " Number of printed rows of an article / weight group
  BEGIN OF ty_pkg_count,
    grp  TYPE i,
    rows TYPE i,
  END OF ty_pkg_count,

  BEGIN OF ty_totals,
    articles  TYPE i,
    quantity  TYPE zqty,
    gross     TYPE zde_grwt,
    net       TYPE zde_net_wt,
    pkg_types TYPE ty_pkg_total_tab,   " in order of first appearance
  END OF ty_totals,

  BEGIN OF ty_hsn_heading,
    hsn  TYPE zna_segment,
    text TYPE string,
  END OF ty_hsn_heading,
  ty_hsn_heading_tab TYPE STANDARD TABLE OF ty_hsn_heading WITH EMPTY KEY.

*----------------------------------------------------------------------*
* Types - Excel reading
*----------------------------------------------------------------------*
TYPES:
  ty_cell_tab TYPE STANDARD TABLE OF string WITH EMPTY KEY,
  BEGIN OF ty_row,
    row   TYPE i,
    cells TYPE ty_cell_tab,
  END OF ty_row,
  ty_row_tab TYPE STANDARD TABLE OF ty_row WITH EMPTY KEY,

  BEGIN OF ty_pl_rows,
    packno TYPE zsd_vbeln,
    rows   TYPE ty_row_tab,
  END OF ty_pl_rows,
  ty_pl_rows_tab TYPE STANDARD TABLE OF ty_pl_rows WITH EMPTY KEY,

  " Header column -> header field (ty_header), found by heading
  BEGIN OF ty_map,
    col    TYPE i,                 " 0 = not in the file
    regex  TYPE string,            " on the normalized heading
    field  TYPE fieldname,
    kind   TYPE char1,             " C char / T multi-line / D date / A amount
    maxlen TYPE i,                 " 0 = unlimited
    label  TYPE string,
  END OF ty_map,
  ty_map_tab TYPE STANDARD TABLE OF ty_map WITH EMPTY KEY,

  " Item / key columns, found by heading (column order is free - A41)
  BEGIN OF ty_col,
    sn       TYPE i,
    inv_sn   TYPE i,
    artno    TYPE i,
    std      TYPE i,
    art      TYPE i,
    partno   TYPE i,
    matdesc  TYPE i,
    qty      TYPE i,
    gwt      TYPE i,
    nwt      TYPE i,
    hsn      TYPE i,
    invoice  TYPE i,
    packno   TYPE i,
    pkg_type TYPE i,
  END OF ty_col.

*----------------------------------------------------------------------*
* Types - SAP data buffers (filled once per run, logic sheet rows)
*----------------------------------------------------------------------*
TYPES:
  BEGIN OF ty_vbrk,
    vbeln TYPE vbrk-vbeln,
    fkart TYPE vbrk-fkart,
    fksto TYPE vbrk-fksto,
    sfakn TYPE vbrk-sfakn,
    bukrs TYPE vbrk-bukrs,
    vkorg TYPE vbrk-vkorg,
    vtweg TYPE vbrk-vtweg,
    spart TYPE vbrk-spart,
    kunag TYPE vbrk-kunag,
    zterm TYPE vbrk-zterm,
    inco1 TYPE vbrk-inco1,
    inco2 TYPE vbrk-inco2,
    bupla TYPE vbrk-bupla,
  END OF ty_vbrk,
  ty_vbrk_hash TYPE HASHED TABLE OF ty_vbrk WITH UNIQUE KEY vbeln,
  ty_vbrk_list TYPE SORTED TABLE OF ty_vbrk WITH UNIQUE KEY vbeln,

  BEGIN OF ty_vbrp,
    vbeln TYPE vbrp-vbeln,
    posnr TYPE vbrp-posnr,
    matnr TYPE vbrp-matnr,
    werks TYPE vbrp-werks,
    aubel TYPE vbrp-aubel,
  END OF ty_vbrp,
  ty_vbrp_tab TYPE SORTED TABLE OF ty_vbrp WITH UNIQUE KEY vbeln posnr,

  BEGIN OF ty_knvp,
    kunnr TYPE knvp-kunnr,
    vkorg TYPE knvp-vkorg,
    vtweg TYPE knvp-vtweg,
    spart TYPE knvp-spart,
    parvw TYPE knvp-parvw,
    parza TYPE knvp-parza,
    kunn2 TYPE knvp-kunn2,
    defpa TYPE knvp-defpa,
  END OF ty_knvp,
  ty_knvp_tab TYPE SORTED TABLE OF ty_knvp WITH UNIQUE KEY kunnr vkorg vtweg spart parvw parza,

  BEGIN OF ty_id_adrnr,
    id    TYPE char10,
    adrnr TYPE adrnr,
  END OF ty_id_adrnr,
  ty_id_adrnr_tab TYPE HASHED TABLE OF ty_id_adrnr WITH UNIQUE KEY id,

  BEGIN OF ty_adrc,
    addrnumber TYPE adrc-addrnumber,
    date_from  TYPE adrc-date_from,
    name1      TYPE adrc-name1,
    street     TYPE adrc-street,
    str_suppl1 TYPE adrc-str_suppl1,
    city1      TYPE adrc-city1,
    city2      TYPE adrc-city2,
    post_code1 TYPE adrc-post_code1,
    region     TYPE adrc-region,
    country    TYPE adrc-country,
    building   TYPE adrc-building,
  END OF ty_adrc,
  ty_adrc_tab TYPE HASHED TABLE OF ty_adrc WITH UNIQUE KEY addrnumber,

  BEGIN OF ty_adr6,
    addrnumber TYPE adr6-addrnumber,
    flgdefault TYPE adr6-flgdefault,
    consnumber TYPE adr6-consnumber,
    smtp_addr  TYPE adr6-smtp_addr,
  END OF ty_adr6,
  ty_adr6_tab TYPE SORTED TABLE OF ty_adr6 WITH NON-UNIQUE KEY addrnumber,

  BEGIN OF ty_adr2,
    addrnumber TYPE adr2-addrnumber,
    r3_user    TYPE adr2-r3_user,
    flgdefault TYPE adr2-flgdefault,
    consnumber TYPE adr2-consnumber,
    telnr_long TYPE adr2-telnr_long,
  END OF ty_adr2,
  ty_adr2_tab TYPE SORTED TABLE OF ty_adr2 WITH NON-UNIQUE KEY addrnumber r3_user,

  BEGIN OF ty_t005t,
    land1 TYPE t005t-land1,
    landx TYPE t005t-landx,
  END OF ty_t005t,
  ty_t005t_tab TYPE HASHED TABLE OF ty_t005t WITH UNIQUE KEY land1,

  BEGIN OF ty_t005u,
    land1 TYPE t005u-land1,
    bland TYPE t005u-bland,
    bezei TYPE t005u-bezei,
  END OF ty_t005u,
  ty_t005u_tab TYPE HASHED TABLE OF ty_t005u WITH UNIQUE KEY land1 bland,

  BEGIN OF ty_branch,
    bukrs  TYPE j_1bbranch-bukrs,
    branch TYPE j_1bbranch-branch,
    gstin  TYPE j_1bbranch-gstin,
  END OF ty_branch,
  ty_branch_tab TYPE HASHED TABLE OF ty_branch WITH UNIQUE KEY bukrs branch,

  BEGIN OF ty_t001z,
    bukrs TYPE t001z-bukrs,
    paval TYPE t001z-paval,
  END OF ty_t001z,
  ty_t001z_tab TYPE HASHED TABLE OF ty_t001z WITH UNIQUE KEY bukrs,

  BEGIN OF ty_vbak,
    vbeln TYPE vbak-vbeln,
    vgbel TYPE vbak-vgbel,
  END OF ty_vbak,
  ty_vbak_tab TYPE HASHED TABLE OF ty_vbak WITH UNIQUE KEY vbeln,

  BEGIN OF ty_vbkd,
    vbeln TYPE vbkd-vbeln,
    bstkd TYPE vbkd-bstkd,
    bstdk TYPE vbkd-bstdk,
  END OF ty_vbkd,
  ty_vbkd_tab TYPE HASHED TABLE OF ty_vbkd WITH UNIQUE KEY vbeln,

  BEGIN OF ty_tvzbt,
    zterm TYPE tvzbt-zterm,
    ztagg TYPE tvzbt-ztagg,
    vtext TYPE tvzbt-vtext,
  END OF ty_tvzbt,
  ty_tvzbt_tab TYPE SORTED TABLE OF ty_tvzbt WITH UNIQUE KEY zterm ztagg,

  BEGIN OF ty_makt,
    matnr TYPE makt-matnr,
    maktx TYPE makt-maktx,
  END OF ty_makt,
  ty_makt_tab TYPE HASHED TABLE OF ty_makt WITH UNIQUE KEY matnr,

  BEGIN OF ty_marc,
    matnr TYPE marc-matnr,
    werks TYPE marc-werks,
    steuc TYPE marc-steuc,
  END OF ty_marc,
  ty_marc_tab TYPE HASHED TABLE OF ty_marc WITH UNIQUE KEY matnr werks,

  BEGIN OF ty_t604n,
    steuc TYPE t604n-steuc,
    text1 TYPE t604n-text1,
  END OF ty_t604n,
  ty_t604n_tab TYPE HASHED TABLE OF ty_t604n WITH UNIQUE KEY steuc,

  BEGIN OF ty_partno_matnr,
    partno TYPE zsd_debezei40,
    matnr  TYPE matnr,
  END OF ty_partno_matnr,
  ty_partno_matnr_tab TYPE HASHED TABLE OF ty_partno_matnr WITH UNIQUE KEY partno,

  BEGIN OF ty_auth,
    object TYPE xuobject,
    value  TYPE char10,
    ok     TYPE abap_bool,
  END OF ty_auth,
  ty_auth_tab TYPE HASHED TABLE OF ty_auth WITH UNIQUE KEY object value,

  " Resolved partner (logic B20 / B26)
  BEGIN OF ty_partner,
    kunnr TYPE kunnr,
    adrnr TYPE adrnr,
  END OF ty_partner,

  " Printed address block of a partner (logic B20-B23 / B26-B29)
  BEGIN OF ty_party_data,
    name    TYPE string,
    address TYPE string,           " lines separated by newline
    tel     TYPE string,
    mobile  TYPE string,
    email   TYPE string,
  END OF ty_party_data.

*----------------------------------------------------------------------*
* Global data
*----------------------------------------------------------------------*
" Run data
DATA gt_packing_lists TYPE ty_packing_list_tab.
DATA gt_file_messages TYPE bapiret2_t.
DATA gt_result        TYPE ty_result_tab.

" Excel column map of the uploaded file
DATA gs_col TYPE ty_col.
DATA gt_map TYPE ty_map_tab.

" SAP data buffers - each table read ONCE per run (no SELECT in loops)
DATA gt_vbrk   TYPE ty_vbrk_hash.
DATA gt_vbrp   TYPE ty_vbrp_tab.
DATA gt_knvp   TYPE ty_knvp_tab.
DATA gt_kna1   TYPE ty_id_adrnr_tab.
DATA gt_t001   TYPE ty_id_adrnr_tab.
DATA gt_adrc   TYPE ty_adrc_tab.
DATA gt_adr6   TYPE ty_adr6_tab.
DATA gt_adr2   TYPE ty_adr2_tab.
DATA gt_t005t  TYPE ty_t005t_tab.
DATA gt_t005u  TYPE ty_t005u_tab.
DATA gt_branch TYPE ty_branch_tab.
DATA gt_t001z  TYPE ty_t001z_tab.
DATA gt_vbak   TYPE ty_vbak_tab.
DATA gt_vbkd   TYPE ty_vbkd_tab.
DATA gt_tvzbt  TYPE ty_tvzbt_tab.
DATA gt_makt   TYPE ty_makt_tab.
DATA gt_marc   TYPE ty_marc_tab.
DATA gt_t604n  TYPE ty_t604n_tab.
DATA gt_matnr  TYPE ty_partno_matnr_tab.
DATA gt_auth   TYPE ty_auth_tab.
