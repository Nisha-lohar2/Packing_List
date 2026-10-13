*&---------------------------------------------------------------------*
*& Report ZSD_EPACK_REPORT
*&---------------------------------------------------------------------*
*& WRICEF 102-B  Export consolidated packing list - data report
*&
*& Input : sales orders (S_VBELN) and the packing list Excel file
*&         (layout of Sample_data_ABAP.xlsx, columns A..AK, row 1 =
*&         headings). The Excel file is the ZSD_PACKNO data source of the
*&         FS logic sheet ("Pass ZSD_PACKNO get ...").
*& Output: one ALV row per packing list line item with all header,
*&         sold-to, ship-to, notifier, shipment, goods, line item and
*&         footer fields of the logic sheet (Logic_sheet1.xlsx).
*&
*& Relationship (validated, see Assumptions_v2.md):
*&   sales order VBAK-VBELN = VBRP-AUBEL  ->  billing documents VBRK
*&   packing list (Excel) column B "SAP Invoice Number" = VBRK-VBELN
*&   -> only packing lists whose invoices belong to the entered sales
*&      orders are reported.
*&
*& All database reads are bulk reads (JOIN / FOR ALL ENTRIES) done once
*& before the output is built. No SELECT and no SORT inside a loop.
*&---------------------------------------------------------------------*
REPORT zsd_epack_report.

*----------------------------------------------------------------------*
* Constants
*----------------------------------------------------------------------*
CONSTANTS:
  gc_langu   TYPE spras VALUE 'E',          " T604N "SPRAS = EN", texts
  gc_land_in TYPE land1 VALUE 'IN',         " T604N "LAND1 = IN"
  gc_cin     TYPE t001z-party VALUE 'ZCIN',       " T001Z "PARTY = ZCIN"
  " Logic sheet "PARVW = SP / SH" = sold-to / ship-to party. Stored in
  " KNVP with the internal codes AG / WE (A4).
  gc_parvw_sp TYPE parvw VALUE 'AG',
  gc_parvw_sh TYPE parvw VALUE 'WE',
  gc_tel      TYPE adr2-r3_user VALUE '1',   " ADR2-R3_USER telephone
  gc_mob      TYPE adr2-r3_user VALUE '3'.   " ADR2-R3_USER mobile

*----------------------------------------------------------------------*
* Types - Excel (ZSD_PACKNO data), columns A..AK, all as read
*----------------------------------------------------------------------*
TYPES:
  BEGIN OF ty_xl,
    row       TYPE i,
    packno    TYPE string,   " A  Packing list Number      ZSD_PACKNO
    invoices  TYPE string,   " B  SAP Invoice Number(s)
    sn        TYPE string,   " C  Sl.No.                   ZSD_SN
    inv_sn    TYPE string,   " D  Sr.No. INV               ZSD_INV_SN
    artno     TYPE string,   " E  Article No.              ZSD_ARTNO
    std       TYPE string,   " F  STD                      ZSD_STD
    art       TYPE string,   " G  Articles                 ZSD_ART
    partno    TYPE string,   " H  Part No.                 ZSD_PARTNO
    matdesc   TYPE string,   " I  Product Description      ZSD_MATDESC
    qty       TYPE string,   " J  Quantity Pcs             ZSD_QTY
    gwt       TYPE string,   " K  Gross Wt. KGS            ZSD_GWT
    nwt       TYPE string,   " L  Net Wt. KGS              ZSD_NWT
    hsn       TYPE string,   " M  H.S Code                 ZSD_HSN
    packdt    TYPE string,   " N  Packing list date        ZSD_PACKDT
    expref    TYPE string,   " O  Exporter Ref (not used - B11 from ADRC)
    party_ref TYPE string,   " P  Buyer's Reference No.    ZSD_PARTY_REF
    ref_dt    TYPE string,   " Q  Buyer's Reference date   ZSD_REF_DT
    cnty_orgn TYPE string,   " R  Country of origin        ZSD_CNTY_ORGN
    np1       TYPE string,   " S  Notifier Party1          ZSD_NP1
    np2       TYPE string,   " T  Notifier Party2          ZSD_NP2
    np3       TYPE string,   " U  Notifier Party3          ZSD_NP3
    np4       TYPE string,   " V  Notifier Party4          ZSD_NP4
    np5       TYPE string,   " W  Notifier Party5          ZSD_NP5
    prec      TYPE string,   " X  Pre-Carriage             ZSD_PREC
    rec_pc    TYPE string,   " Y  Place of Receipt         ZSD_REC_PC
    pol       TYPE string,   " Z  Port of Loading          ZSD_POL
    pod       TYPE string,   " AA Port of Discharge        ZSD_POD
    pld       TYPE string,   " AB Place of Delivery        ZSD_PLD
    mark      TYPE string,   " AC Marks & Nos.             ZSD_MARK
    vsl_flt   TYPE string,   " AD Vessel / Flight No.      ZSD_VSL_FLT
    cont      TYPE string,   " AE Container No             ZSD_CONT
    nopack    TYPE string,   " AF No. of Packages          ZSD_NOPACK
    goods     TYPE string,   " AG Description of goods (not used - B42)
    decl      TYPE string,   " AH Declaration of PL        ZSD_DECL
  END OF ty_xl,
  ty_xl_tab TYPE STANDARD TABLE OF ty_xl WITH EMPTY KEY.

*----------------------------------------------------------------------*
* Types - final output (sequence of the logic sheet)
*----------------------------------------------------------------------*
TYPES:
  BEGIN OF ty_out,
    " Reference
    sales_orders  TYPE string,             " entered sales orders of the PL
    invoices      TYPE string,             " billing documents of the PL
    " Header - Exporter (B3-B7)
    exp_name1     TYPE adrc-name1,
    exp_street    TYPE adrc-street,
    exp_str_supl1 TYPE adrc-str_suppl1,
    exp_city1     TYPE adrc-city1,
    exp_post_code TYPE adrc-post_code1,
    exp_region    TYPE adrc-region,
    exp_region_tx TYPE t005u-bezei,
    exp_country   TYPE land1,
    exp_country_tx TYPE t005t-landx,
    gstin         TYPE j_1bbranch-gstin,
    cin           TYPE t001z-paval,
    " Header (B9-B17)
    packno        TYPE string,
    packdt        TYPE d,
    iec           TYPE adrc-building,
    order_no      TYPE string,
    order_date    TYPE string,
    buyer_ref     TYPE string,
    buyer_ref_dt  TYPE d,
    cnty_orgn     TYPE string,
    cnty_dest     TYPE t005t-landx,
    " Sold-to / bill-to party (B20-B23)
    sp_kunnr      TYPE kunnr,
    sp_name1      TYPE adrc-name1,
    sp_street     TYPE adrc-street,
    sp_str_supl1  TYPE adrc-str_suppl1,
    sp_city1      TYPE adrc-city1,
    sp_city2      TYPE adrc-city2,
    sp_post_code  TYPE adrc-post_code1,
    sp_country    TYPE land1,
    sp_country_tx TYPE t005t-landx,
    sp_email      TYPE adr6-smtp_addr,
    sp_telephone  TYPE adr2-telnr_long,
    sp_mobile     TYPE adr2-telnr_long,
    " Ship-to party (B26-B29)
    sh_kunnr      TYPE kunnr,
    sh_name1      TYPE adrc-name1,
    sh_street     TYPE adrc-street,
    sh_str_supl1  TYPE adrc-str_suppl1,
    sh_city1      TYPE adrc-city1,
    sh_city2      TYPE adrc-city2,
    sh_post_code  TYPE adrc-post_code1,
    sh_country    TYPE land1,
    sh_country_tx TYPE t005t-landx,
    sh_email      TYPE adr6-smtp_addr,
    sh_telephone  TYPE adr2-telnr_long,
    sh_mobile     TYPE adr2-telnr_long,
    " Notifier parties (B31)
    np1           TYPE string,
    np2           TYPE string,
    np3           TYPE string,
    np4           TYPE string,
    np5           TYPE string,
    " Shipment (B33-B41)
    pre_carriage  TYPE string,
    place_receipt TYPE string,
    vessel_flight TYPE string,
    port_loading  TYPE string,
    port_disch    TYPE string,
    place_deliv   TYPE string,
    marks_nos     TYPE string,
    container_no  TYPE string,
    no_packages   TYPE string,
    " Description of goods (B42), payment term (B43), shipment (B44)
    goods_desc    TYPE string,
    zterm         TYPE vbrk-zterm,
    payment_term  TYPE tvzbt-vtext,
    inco1         TYPE vbrk-inco1,
    inco2         TYPE vbrk-inco2,
    " Line item (B47-B58)
    sn            TYPE i,
    inv_sn        TYPE i,
    article_no    TYPE string,
    std           TYPE i,
    articles      TYPE i,
    total_art     TYPE i,
    part_no       TYPE string,
    prod_desc     TYPE string,
    quantity      TYPE p LENGTH 13 DECIMALS 3,
    gross_wt      TYPE p LENGTH 13 DECIMALS 3,
    net_wt        TYPE p LENGTH 13 DECIMALS 3,
    hsn           TYPE string,
    " Footer (B61)
    declaration   TYPE string,
  END OF ty_out,
  ty_out_tab TYPE STANDARD TABLE OF ty_out WITH EMPTY KEY.

*----------------------------------------------------------------------*
* Types - SAP data buffers
*----------------------------------------------------------------------*
TYPES:
  " Billing item of an entered sales order with its billing header
  BEGIN OF ty_bill,
    vbeln TYPE vbrk-vbeln,
    posnr TYPE vbrp-posnr,
    matnr TYPE vbrp-matnr,
    werks TYPE vbrp-werks,
    aubel TYPE vbrp-aubel,
    bukrs TYPE vbrk-bukrs,
    bupla TYPE vbrk-bupla,
    kunag TYPE vbrk-kunag,
    vkorg TYPE vbrk-vkorg,
    vtweg TYPE vbrk-vtweg,
    spart TYPE vbrk-spart,
    zterm TYPE vbrk-zterm,
    inco1 TYPE vbrk-inco1,
    inco2 TYPE vbrk-inco2,
  END OF ty_bill,
  ty_bill_tab TYPE SORTED TABLE OF ty_bill WITH UNIQUE KEY vbeln posnr,

  ty_vbeln_tab TYPE SORTED TABLE OF vbeln_vf WITH UNIQUE KEY table_line,

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

  BEGIN OF ty_t001,
    bukrs TYPE t001-bukrs,
    adrnr TYPE t001-adrnr,
  END OF ty_t001,
  ty_t001_tab TYPE HASHED TABLE OF ty_t001 WITH UNIQUE KEY bukrs,

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
  ty_adrc_std  TYPE STANDARD TABLE OF ty_adrc WITH EMPTY KEY,
  ty_adrc_tab  TYPE HASHED TABLE OF ty_adrc WITH UNIQUE KEY addrnumber,

  BEGIN OF ty_adr6,
    addrnumber TYPE adr6-addrnumber,
    flgdefault TYPE adr6-flgdefault,
    consnumber TYPE adr6-consnumber,
    smtp_addr  TYPE adr6-smtp_addr,
  END OF ty_adr6,
  ty_adr6_tab TYPE STANDARD TABLE OF ty_adr6 WITH EMPTY KEY,

  BEGIN OF ty_adr2,
    addrnumber TYPE adr2-addrnumber,
    r3_user    TYPE adr2-r3_user,
    flgdefault TYPE adr2-flgdefault,
    consnumber TYPE adr2-consnumber,
    telnr_long TYPE adr2-telnr_long,
  END OF ty_adr2,
  ty_adr2_tab TYPE STANDARD TABLE OF ty_adr2 WITH EMPTY KEY,

  BEGIN OF ty_t005u,
    land1 TYPE t005u-land1,
    bland TYPE t005u-bland,
    bezei TYPE t005u-bezei,
  END OF ty_t005u,
  ty_t005u_tab TYPE HASHED TABLE OF ty_t005u WITH UNIQUE KEY land1 bland,

  BEGIN OF ty_t005t,
    land1 TYPE t005t-land1,
    landx TYPE t005t-landx,
  END OF ty_t005t,
  ty_t005t_tab TYPE HASHED TABLE OF ty_t005t WITH UNIQUE KEY land1,

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

  BEGIN OF ty_knvp,
    kunnr TYPE knvp-kunnr,
    vkorg TYPE knvp-vkorg,
    vtweg TYPE knvp-vtweg,
    spart TYPE knvp-spart,
    parvw TYPE knvp-parvw,
    parza TYPE knvp-parza,
    kunn2 TYPE knvp-kunn2,
    adrnr TYPE kna1-adrnr,
  END OF ty_knvp,
  ty_knvp_tab TYPE SORTED TABLE OF ty_knvp
              WITH NON-UNIQUE KEY kunnr vkorg vtweg spart parvw parza,

  BEGIN OF ty_tvzbt,
    zterm TYPE tvzbt-zterm,
    ztagg TYPE tvzbt-ztagg,
    vtext TYPE tvzbt-vtext,
  END OF ty_tvzbt,
  ty_tvzbt_tab TYPE SORTED TABLE OF ty_tvzbt WITH UNIQUE KEY zterm ztagg,

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

  BEGIN OF ty_makt,
    matnr TYPE makt-matnr,
    maktx TYPE makt-maktx,
  END OF ty_makt,
  ty_makt_tab TYPE HASHED TABLE OF ty_makt WITH UNIQUE KEY matnr,

  BEGIN OF ty_partno,
    partno TYPE string,
    matnr  TYPE matnr,
  END OF ty_partno,
  ty_partno_tab TYPE HASHED TABLE OF ty_partno WITH UNIQUE KEY partno,

  " Address block of a partner
  BEGIN OF ty_party,
    kunnr TYPE kunnr,
    adrnr TYPE adrnr,
  END OF ty_party,

  " Message log
  BEGIN OF ty_log,
    type   TYPE symsgty,
    packno TYPE string,
    text   TYPE string,
  END OF ty_log,
  ty_log_tab TYPE STANDARD TABLE OF ty_log WITH EMPTY KEY,

  ty_string_tab TYPE STANDARD TABLE OF string WITH EMPTY KEY.

*----------------------------------------------------------------------*
* Global data
*----------------------------------------------------------------------*
DATA gv_vbeln TYPE vbak-vbeln.                       " for SELECT-OPTIONS

DATA gt_xl     TYPE ty_xl_tab.                       " Excel rows
DATA gt_out    TYPE ty_out_tab.                      " final output
DATA gt_log    TYPE ty_log_tab.                      " messages

DATA gt_vbak   TYPE ty_vbak_tab.
DATA gt_bill   TYPE ty_bill_tab.
DATA gt_inv    TYPE ty_vbeln_tab.                    " billing docs of the SOs
DATA gt_vbkd   TYPE ty_vbkd_tab.
DATA gt_t001   TYPE ty_t001_tab.
DATA gt_adrc   TYPE ty_adrc_tab.
DATA gt_adr6   TYPE ty_adr6_tab.
DATA gt_adr2   TYPE ty_adr2_tab.
DATA gt_t005u  TYPE ty_t005u_tab.
DATA gt_t005t  TYPE ty_t005t_tab.
DATA gt_branch TYPE ty_branch_tab.
DATA gt_t001z  TYPE ty_t001z_tab.
DATA gt_knvp   TYPE ty_knvp_tab.
DATA gt_tvzbt  TYPE ty_tvzbt_tab.
DATA gt_marc   TYPE ty_marc_tab.
DATA gt_t604n  TYPE ty_t604n_tab.
DATA gt_makt   TYPE ty_makt_tab.
DATA gt_partno TYPE ty_partno_tab.

*----------------------------------------------------------------------*
* Selection screen
*----------------------------------------------------------------------*
SELECTION-SCREEN BEGIN OF BLOCK b1 WITH FRAME TITLE TEXT-b01.
SELECT-OPTIONS s_vbeln FOR gv_vbeln OBLIGATORY.
PARAMETERS     p_file  TYPE rlgrap-filename LOWER CASE OBLIGATORY.
SELECTION-SCREEN END OF BLOCK b1.

*----------------------------------------------------------------------*
AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_file.
*----------------------------------------------------------------------*
  PERFORM f4_file CHANGING p_file.

*----------------------------------------------------------------------*
START-OF-SELECTION.
*----------------------------------------------------------------------*
  PERFORM read_excel.            " file -> GT_XL (ZSD_PACKNO data)
  PERFORM read_sales_orders.     " VBAK, VBRP/VBRK of the sales orders
  PERFORM read_master_data.      " all other tables, bulk
  PERFORM build_output.          " GT_OUT, one row per PL line item
  PERFORM display.

*&---------------------------------------------------------------------*
*& 1. Excel file -> GT_XL
*&---------------------------------------------------------------------*
FORM read_excel.

  DATA lt_bin     TYPE solix_tab.
  DATA lv_size    TYPE i.
  DATA lv_xstring TYPE xstring.
  DATA lv_file    TYPE string.
  DATA lo_excel   TYPE REF TO cl_fdt_xl_spreadsheet.
  DATA lt_sheets  TYPE if_fdt_doc_spreadsheet=>t_worksheet_names.
  DATA lv_sheet   TYPE string.
  DATA lr_data    TYPE REF TO data.
  DATA ls_xl      TYPE ty_xl.
  DATA lv_row     TYPE i.
  DATA lv_packno  TYPE string.

  FIELD-SYMBOLS <lt_excel> TYPE STANDARD TABLE.
  FIELD-SYMBOLS <ls_excel> TYPE any.

  lv_file = p_file.

  " File in binary form
  cl_gui_frontend_services=>gui_upload(
    EXPORTING
      filename        = lv_file
      filetype        = 'BIN'
    IMPORTING
      filelength      = lv_size
    CHANGING
      data_tab        = lt_bin
    EXCEPTIONS
      file_open_error = 1
      file_read_error = 2
      OTHERS          = 3 ).
  IF sy-subrc <> 0 OR lv_size = 0.
    MESSAGE |File { lv_file } could not be read| TYPE 'S' DISPLAY LIKE 'E'.
    LEAVE LIST-PROCESSING.
  ENDIF.
  lv_xstring = cl_bcs_convert=>solix_to_xstring( it_solix = lt_bin
                                                 iv_size  = lv_size ).

  " First worksheet -> internal table, as it is
  TRY.
      CREATE OBJECT lo_excel
        EXPORTING
          document_name = lv_file
          xdocument     = lv_xstring.
    CATCH cx_fdt_excel_core.
      MESSAGE 'The file is not a valid Excel (.xlsx) file' TYPE 'S' DISPLAY LIKE 'E'.
      LEAVE LIST-PROCESSING.
  ENDTRY.

  lo_excel->if_fdt_doc_spreadsheet~get_worksheet_names(
    IMPORTING worksheet_names = lt_sheets ).
  READ TABLE lt_sheets INTO lv_sheet INDEX 1.
  IF sy-subrc <> 0.
    MESSAGE 'The Excel file has no worksheet' TYPE 'S' DISPLAY LIKE 'E'.
    LEAVE LIST-PROCESSING.
  ENDIF.

  lr_data = lo_excel->if_fdt_doc_spreadsheet~get_itab_from_worksheet( lv_sheet ).
  ASSIGN lr_data->* TO <lt_excel>.
  IF <lt_excel> IS NOT ASSIGNED OR lines( <lt_excel> ) < 2.
    MESSAGE 'The worksheet contains no data below the heading row' TYPE 'S' DISPLAY LIKE 'E'.
    LEAVE LIST-PROCESSING.
  ENDIF.

  " Row 1 = headings
  DELETE <lt_excel> INDEX 1.

  LOOP AT <lt_excel> ASSIGNING <ls_excel>.
    lv_row = sy-tabix + 1.                         " Excel row number
    CLEAR ls_xl.
    ls_xl-row = lv_row.
    PERFORM cell USING <ls_excel>  1 CHANGING ls_xl-packno.
    PERFORM cell USING <ls_excel>  2 CHANGING ls_xl-invoices.
    PERFORM cell USING <ls_excel>  3 CHANGING ls_xl-sn.
    PERFORM cell USING <ls_excel>  4 CHANGING ls_xl-inv_sn.
    PERFORM cell USING <ls_excel>  5 CHANGING ls_xl-artno.
    PERFORM cell USING <ls_excel>  6 CHANGING ls_xl-std.
    PERFORM cell USING <ls_excel>  7 CHANGING ls_xl-art.
    PERFORM cell USING <ls_excel>  8 CHANGING ls_xl-partno.
    PERFORM cell USING <ls_excel>  9 CHANGING ls_xl-matdesc.
    PERFORM cell USING <ls_excel> 10 CHANGING ls_xl-qty.
    PERFORM cell USING <ls_excel> 11 CHANGING ls_xl-gwt.
    PERFORM cell USING <ls_excel> 12 CHANGING ls_xl-nwt.
    PERFORM cell USING <ls_excel> 13 CHANGING ls_xl-hsn.
    PERFORM cell USING <ls_excel> 14 CHANGING ls_xl-packdt.
    PERFORM cell USING <ls_excel> 15 CHANGING ls_xl-expref.
    PERFORM cell USING <ls_excel> 16 CHANGING ls_xl-party_ref.
    PERFORM cell USING <ls_excel> 17 CHANGING ls_xl-ref_dt.
    PERFORM cell USING <ls_excel> 18 CHANGING ls_xl-cnty_orgn.
    PERFORM cell USING <ls_excel> 19 CHANGING ls_xl-np1.
    PERFORM cell USING <ls_excel> 20 CHANGING ls_xl-np2.
    PERFORM cell USING <ls_excel> 21 CHANGING ls_xl-np3.
    PERFORM cell USING <ls_excel> 22 CHANGING ls_xl-np4.
    PERFORM cell USING <ls_excel> 23 CHANGING ls_xl-np5.
    PERFORM cell USING <ls_excel> 24 CHANGING ls_xl-prec.
    PERFORM cell USING <ls_excel> 25 CHANGING ls_xl-rec_pc.
    PERFORM cell USING <ls_excel> 26 CHANGING ls_xl-pol.
    PERFORM cell USING <ls_excel> 27 CHANGING ls_xl-pod.
    PERFORM cell USING <ls_excel> 28 CHANGING ls_xl-pld.
    PERFORM cell USING <ls_excel> 29 CHANGING ls_xl-mark.
    PERFORM cell USING <ls_excel> 30 CHANGING ls_xl-vsl_flt.
    PERFORM cell USING <ls_excel> 31 CHANGING ls_xl-cont.
    PERFORM cell USING <ls_excel> 32 CHANGING ls_xl-nopack.
    PERFORM cell USING <ls_excel> 33 CHANGING ls_xl-goods.
    PERFORM cell USING <ls_excel> 34 CHANGING ls_xl-decl.

    IF ls_xl IS INITIAL OR ( ls_xl-packno IS INITIAL AND ls_xl-sn IS INITIAL ).
      CONTINUE.                                     " empty row
    ENDIF.

    " A row without packing list number belongs to the packing list above
    IF ls_xl-packno IS INITIAL.
      ls_xl-packno = lv_packno.
    ENDIF.
    lv_packno = ls_xl-packno.

    IF ls_xl-packno IS INITIAL.
      APPEND VALUE #( type = 'W' text = |Row { lv_row }: no packing list number - row ignored| ) TO gt_log.
      CONTINUE.
    ENDIF.

    APPEND ls_xl TO gt_xl.
  ENDLOOP.

  IF gt_xl IS INITIAL.
    MESSAGE 'The Excel file contains no packing list data' TYPE 'S' DISPLAY LIKE 'E'.
    LEAVE LIST-PROCESSING.
  ENDIF.

ENDFORM.

*&---------------------------------------------------------------------*
*& Cell n of an Excel row as text. Placeholder texts of the template
*& ("Fetch from logic", "fetch from material master") mean "no value".
*&---------------------------------------------------------------------*
FORM cell USING    is_row  TYPE any
                   iv_col  TYPE i
          CHANGING cv_text TYPE string.

  FIELD-SYMBOLS <lv_cell> TYPE any.

  CLEAR cv_text.
  ASSIGN COMPONENT iv_col OF STRUCTURE is_row TO <lv_cell>.
  IF sy-subrc <> 0.
    RETURN.
  ENDIF.

  cv_text = <lv_cell>.
  IF condense( cv_text ) IS INITIAL.
    CLEAR cv_text.
    RETURN.
  ENDIF.

  IF to_upper( condense( cv_text ) ) CP 'FETCH FROM*'.
    CLEAR cv_text.
  ENDIF.

ENDFORM.

*&---------------------------------------------------------------------*
*& 2. Sales orders -> billing documents (VBRP-AUBEL = sales order)
*&---------------------------------------------------------------------*
FORM read_sales_orders.

  DATA lv_text   TYPE string.
  DATA lt_billed TYPE SORTED TABLE OF vbak-vbeln WITH UNIQUE KEY table_line.

  " Sales orders exist? VGBEL = quotation (logic B13)
  SELECT vbeln, vgbel
    FROM vbak
    WHERE vbeln IN @s_vbeln
    INTO TABLE @gt_vbak.
  IF gt_vbak IS INITIAL.
    MESSAGE 'No sales order found for the selection' TYPE 'S' DISPLAY LIKE 'E'.
    LEAVE LIST-PROCESSING.
  ENDIF.

  " Single values entered that do not exist
  LOOP AT s_vbeln INTO DATA(ls_sel) WHERE sign = 'I' AND option = 'EQ'.
    IF NOT line_exists( gt_vbak[ vbeln = ls_sel-low ] ).
      APPEND VALUE #( type = 'E' text = |Sales order { ls_sel-low ALPHA = OUT } does not exist| ) TO gt_log.
    ENDIF.
  ENDLOOP.

  " Billing items of the sales orders with their billing header.
  " Cancelled billing documents and cancellation documents are excluded (A2).
  SELECT r~vbeln, r~posnr, r~matnr, r~werks, r~aubel,
         k~bukrs, k~bupla, k~kunag, k~vkorg, k~vtweg, k~spart,
         k~zterm, k~inco1, k~inco2
    FROM vbrp AS r
    INNER JOIN vbrk AS k ON k~vbeln = r~vbeln
    WHERE r~aubel IN @s_vbeln
      AND k~fksto = @space
      AND k~sfakn = @space
    INTO TABLE @gt_bill.
  IF gt_bill IS INITIAL.
    MESSAGE 'No billing document found for the entered sales orders' TYPE 'S' DISPLAY LIKE 'E'.
    LEAVE LIST-PROCESSING.
  ENDIF.

  LOOP AT gt_bill INTO DATA(ls_bill).
    INSERT ls_bill-vbeln INTO TABLE gt_inv.
    INSERT ls_bill-aubel INTO TABLE lt_billed.
  ENDLOOP.

  " Sales orders without billing document
  LOOP AT gt_vbak INTO DATA(ls_vbak).
    IF NOT line_exists( lt_billed[ table_line = ls_vbak-vbeln ] ).
      lv_text = |Sales order { ls_vbak-vbeln ALPHA = OUT } has no billing document|.
      APPEND VALUE #( type = 'W' text = lv_text ) TO gt_log.
    ENDIF.
  ENDLOOP.

ENDFORM.

*&---------------------------------------------------------------------*
*& 3. All master data, each table read once (FOR ALL ENTRIES / JOIN)
*&---------------------------------------------------------------------*
FORM read_master_data.

  TYPES: BEGIN OF ty_area,
           kunnr TYPE kunnr,
           vkorg TYPE vkorg,
           vtweg TYPE vtweg,
           spart TYPE spart,
         END OF ty_area,
         BEGIN OF ty_branch_key,
           bukrs  TYPE bukrs,
           branch TYPE j_1bbranch-branch,
         END OF ty_branch_key,
         BEGIN OF ty_region,
           land1 TYPE land1,
           bland TYPE adrc-region,
         END OF ty_region,
         BEGIN OF ty_matwerks,
           matnr TYPE matnr,
           werks TYPE werks_d,
         END OF ty_matwerks.

  DATA lt_bukrs    TYPE SORTED TABLE OF bukrs WITH UNIQUE KEY table_line.
  DATA lt_area     TYPE SORTED TABLE OF ty_area WITH UNIQUE KEY kunnr vkorg vtweg spart.
  DATA lt_branch   TYPE SORTED TABLE OF ty_branch_key WITH UNIQUE KEY bukrs branch.
  DATA lt_zterm    TYPE SORTED TABLE OF vbrk-zterm WITH UNIQUE KEY table_line.
  DATA lt_vgbel    TYPE SORTED TABLE OF vbak-vgbel WITH UNIQUE KEY table_line.
  DATA lt_adrnr    TYPE SORTED TABLE OF adrnr WITH UNIQUE KEY table_line.
  DATA lt_land     TYPE SORTED TABLE OF land1 WITH UNIQUE KEY table_line.
  DATA lt_region   TYPE SORTED TABLE OF ty_region WITH UNIQUE KEY land1 bland.
  DATA lt_matwerks TYPE SORTED TABLE OF ty_matwerks WITH UNIQUE KEY matnr werks.
  DATA lt_matnr    TYPE SORTED TABLE OF matnr WITH UNIQUE KEY table_line.
  DATA lt_steuc    TYPE SORTED TABLE OF marc-steuc WITH UNIQUE KEY table_line.
  DATA lt_adrc     TYPE ty_adrc_std.
  DATA lv_matnr    TYPE matnr.

  " Keys from the billing documents
  LOOP AT gt_bill INTO DATA(ls_bill).
    INSERT ls_bill-bukrs INTO TABLE lt_bukrs.
    INSERT VALUE #( kunnr = ls_bill-kunag vkorg = ls_bill-vkorg
                    vtweg = ls_bill-vtweg spart = ls_bill-spart ) INTO TABLE lt_area.
    INSERT VALUE #( bukrs = ls_bill-bukrs branch = ls_bill-bupla ) INTO TABLE lt_branch.
    INSERT ls_bill-zterm INTO TABLE lt_zterm.
    IF ls_bill-matnr IS NOT INITIAL.
      INSERT VALUE #( matnr = ls_bill-matnr werks = ls_bill-werks ) INTO TABLE lt_matwerks.
      INSERT ls_bill-matnr INTO TABLE lt_matnr.
    ENDIF.
  ENDLOOP.

  LOOP AT gt_vbak INTO DATA(ls_vbak) WHERE vgbel IS NOT INITIAL.
    INSERT ls_vbak-vgbel INTO TABLE lt_vgbel.
  ENDLOOP.

  " Part No. of the Excel = material number (logic B58 "from STEP no 53")
  LOOP AT gt_xl INTO DATA(ls_xl) WHERE partno IS NOT INITIAL.
    IF line_exists( gt_partno[ partno = ls_xl-partno ] ).
      CONTINUE.
    ENDIF.
    CLEAR lv_matnr.
    CALL FUNCTION 'CONVERSION_EXIT_MATN1_INPUT'
      EXPORTING
        input        = ls_xl-partno
      IMPORTING
        output       = lv_matnr
      EXCEPTIONS
        length_error = 1
        OTHERS       = 2.
    IF sy-subrc <> 0.
      CLEAR lv_matnr.
    ENDIF.
    INSERT VALUE #( partno = ls_xl-partno matnr = lv_matnr ) INTO TABLE gt_partno.
    IF lv_matnr IS NOT INITIAL.
      INSERT lv_matnr INTO TABLE lt_matnr.
    ENDIF.
  ENDLOOP.

  " Order no. & date (B13): quotation VBKD header data
  IF lt_vgbel IS NOT INITIAL.
    SELECT vbeln, bstkd, bstdk
      FROM vbkd
      FOR ALL ENTRIES IN @lt_vgbel
      WHERE vbeln = @lt_vgbel-table_line
        AND posnr = '000000'
      INTO TABLE @gt_vbkd.
  ENDIF.

  " Exporter (B3, B11): T001-ADRNR
  SELECT bukrs, adrnr
    FROM t001
    FOR ALL ENTRIES IN @lt_bukrs
    WHERE bukrs = @lt_bukrs-table_line
    INTO TABLE @gt_t001.
  LOOP AT gt_t001 INTO DATA(ls_t001) WHERE adrnr IS NOT INITIAL.
    INSERT ls_t001-adrnr INTO TABLE lt_adrnr.
  ENDLOOP.

  " GST No. (B6)
  SELECT bukrs, branch, gstin
    FROM j_1bbranch
    FOR ALL ENTRIES IN @lt_branch
    WHERE bukrs  = @lt_branch-bukrs
      AND branch = @lt_branch-branch
    INTO TABLE @gt_branch.

  " CIN (B7)
  SELECT bukrs, paval
    FROM t001z
    FOR ALL ENTRIES IN @lt_bukrs
    WHERE bukrs = @lt_bukrs-table_line
      AND party = @gc_cin
    INTO TABLE @gt_t001z.

  " Sold-to / ship-to (B20, B26): KNVP partner -> KNA1-ADRNR
  SELECT p~kunnr, p~vkorg, p~vtweg, p~spart, p~parvw, p~parza, p~kunn2, c~adrnr
    FROM knvp AS p
    INNER JOIN kna1 AS c ON c~kunnr = p~kunn2
    FOR ALL ENTRIES IN @lt_area
    WHERE p~kunnr = @lt_area-kunnr
      AND p~vkorg = @lt_area-vkorg
      AND p~vtweg = @lt_area-vtweg
      AND p~spart = @lt_area-spart
      AND p~parvw IN ( @gc_parvw_sp, @gc_parvw_sh )
    INTO TABLE @gt_knvp.
  LOOP AT gt_knvp INTO DATA(ls_knvp) WHERE adrnr IS NOT INITIAL.
    INSERT ls_knvp-adrnr INTO TABLE lt_adrnr.
  ENDLOOP.

  " Addresses (ADRC), e-mail (ADR6), telephone / mobile (ADR2)
  IF lt_adrnr IS NOT INITIAL.
    SELECT addrnumber, date_from, name1, street, str_suppl1, city1, city2,
           post_code1, region, country, building
      FROM adrc
      FOR ALL ENTRIES IN @lt_adrnr
      WHERE addrnumber = @lt_adrnr-table_line
        AND date_from <= @sy-datum
        AND nation     = @space
      INTO TABLE @lt_adrc.
    " Latest valid version per address
    SORT lt_adrc BY addrnumber date_from DESCENDING.
    DELETE ADJACENT DUPLICATES FROM lt_adrc COMPARING addrnumber.
    gt_adrc = lt_adrc.

    SELECT addrnumber, flgdefault, consnumber, smtp_addr
      FROM adr6
      FOR ALL ENTRIES IN @lt_adrnr
      WHERE addrnumber = @lt_adrnr-table_line
      INTO TABLE @gt_adr6.
    " Default entry first, then lowest sequence number
    SORT gt_adr6 BY addrnumber flgdefault DESCENDING consnumber.

    SELECT addrnumber, r3_user, flgdefault, consnumber, telnr_long
      FROM adr2
      FOR ALL ENTRIES IN @lt_adrnr
      WHERE addrnumber = @lt_adrnr-table_line
        AND r3_user IN ( @gc_tel, @gc_mob )
      INTO TABLE @gt_adr2.
    SORT gt_adr2 BY addrnumber r3_user flgdefault DESCENDING consnumber.
  ENDIF.

  " Region and country texts (B4, B5, B17, B20, B26)
  LOOP AT gt_adrc INTO DATA(ls_adrc).
    IF ls_adrc-country IS NOT INITIAL.
      INSERT ls_adrc-country INTO TABLE lt_land.
      IF ls_adrc-region IS NOT INITIAL.
        INSERT VALUE #( land1 = ls_adrc-country bland = ls_adrc-region ) INTO TABLE lt_region.
      ENDIF.
    ENDIF.
  ENDLOOP.
  IF lt_region IS NOT INITIAL.
    " Region text BEZEI is in T005U (logic sheet: T005S) - A3
    SELECT land1, bland, bezei
      FROM t005u
      FOR ALL ENTRIES IN @lt_region
      WHERE spras = @gc_langu
        AND land1 = @lt_region-land1
        AND bland = @lt_region-bland
      INTO TABLE @gt_t005u.
  ENDIF.
  IF lt_land IS NOT INITIAL.
    " Country text LANDX is in T005T (logic sheet: T005) - A3
    SELECT land1, landx
      FROM t005t
      FOR ALL ENTRIES IN @lt_land
      WHERE spras = @gc_langu
        AND land1 = @lt_land-table_line
      INTO TABLE @gt_t005t.
  ENDIF.

  " Payment term text (B43)
  IF lt_zterm IS NOT INITIAL.
    SELECT zterm, ztagg, vtext
      FROM tvzbt
      FOR ALL ENTRIES IN @lt_zterm
      WHERE spras = @gc_langu
        AND zterm = @lt_zterm-table_line
      INTO TABLE @gt_tvzbt.
  ENDIF.

  " HSN of the billed materials (B42, B58)
  IF lt_matwerks IS NOT INITIAL.
    SELECT matnr, werks, steuc
      FROM marc
      FOR ALL ENTRIES IN @lt_matwerks
      WHERE matnr = @lt_matwerks-matnr
        AND werks = @lt_matwerks-werks
      INTO TABLE @gt_marc.
  ENDIF.
  LOOP AT gt_marc INTO DATA(ls_marc) WHERE steuc IS NOT INITIAL.
    INSERT ls_marc-steuc INTO TABLE lt_steuc.
  ENDLOOP.
  IF lt_steuc IS NOT INITIAL.
    SELECT steuc, text1
      FROM t604n
      FOR ALL ENTRIES IN @lt_steuc
      WHERE spras = @gc_langu
        AND land1 = @gc_land_in
        AND steuc = @lt_steuc-table_line
      INTO TABLE @gt_t604n.
  ENDIF.

  " Material description (B54) - MAKTX is in MAKT (logic sheet: MARA) - A3
  IF lt_matnr IS NOT INITIAL.
    SELECT matnr, maktx
      FROM makt
      FOR ALL ENTRIES IN @lt_matnr
      WHERE matnr = @lt_matnr-table_line
        AND spras = @gc_langu
      INTO TABLE @gt_makt.
  ENDIF.

ENDFORM.

*&---------------------------------------------------------------------*
*& 4. Final table: per packing list, header data from SAP and from the
*&    Excel (ZSD_PACKNO), one output row per line item
*&---------------------------------------------------------------------*
FORM build_output.

  DATA lt_packno   TYPE ty_string_tab.
  DATA ls_head     TYPE ty_xl.
  DATA lt_inv      TYPE ty_vbeln_tab.
  DATA lt_inv_ok   TYPE ty_vbeln_tab.
  DATA lt_items    TYPE ty_xl_tab.
  DATA ls_out      TYPE ty_out.
  DATA ls_line     TYPE ty_out.
  DATA lv_text     TYPE string.

  " Packing lists in the order of the file
  LOOP AT gt_xl INTO DATA(ls_xl).
    IF NOT line_exists( lt_packno[ table_line = ls_xl-packno ] ).
      APPEND ls_xl-packno TO lt_packno.
    ENDIF.
  ENDLOOP.

  LOOP AT lt_packno INTO DATA(lv_packno).

    " ZSD_PACKNO data of this packing list: header values = first filled
    " value of the packing list rows, items = rows with Sl.No.
    PERFORM pl_data USING lv_packno CHANGING ls_head lt_inv lt_items.

    " Invoices of the packing list that belong to the entered sales orders
    CLEAR lt_inv_ok.
    LOOP AT lt_inv INTO DATA(lv_inv).
      IF line_exists( gt_inv[ table_line = lv_inv ] ).
        INSERT lv_inv INTO TABLE lt_inv_ok.
      ELSE.
        lv_text = |Invoice { lv_inv ALPHA = OUT } does not belong to the entered sales orders (or is cancelled) - ignored|.
        APPEND VALUE #( type = 'W' packno = lv_packno text = lv_text ) TO gt_log.
      ENDIF.
    ENDLOOP.

    IF lt_inv_ok IS INITIAL.
      " Packing list of other sales orders - not reported
      CONTINUE.
    ENDIF.

    IF lt_items IS INITIAL.
      APPEND VALUE #( type = 'E' packno = lv_packno text = 'No line items (Sl.No.) in the file' ) TO gt_log.
      CONTINUE.
    ENDIF.

    PERFORM pl_header USING lv_packno ls_head lt_inv_ok CHANGING ls_out.

    LOOP AT lt_items INTO DATA(ls_item).
      ls_line = ls_out.
      PERFORM pl_item USING lv_packno ls_item lt_inv_ok CHANGING ls_line.
      APPEND ls_line TO gt_out.
    ENDLOOP.

  ENDLOOP.

  IF gt_out IS INITIAL.
    APPEND VALUE #( type = 'E' text = 'No packing list in the file belongs to the entered sales orders' ) TO gt_log.
  ENDIF.

ENDFORM.

*&---------------------------------------------------------------------*
*& Rows of one packing list: header values, invoice numbers, items
*&---------------------------------------------------------------------*
FORM pl_data USING    iv_packno TYPE string
             CHANGING cs_head   TYPE ty_xl
                      ct_inv    TYPE ty_vbeln_tab
                      ct_items  TYPE ty_xl_tab.

  DATA lt_token TYPE ty_string_tab.
  DATA lv_vbeln TYPE vbeln_vf.
  DATA lv_text  TYPE string.

  FIELD-SYMBOLS <lv_head> TYPE any.
  FIELD-SYMBOLS <lv_row>  TYPE any.

  CLEAR: cs_head, ct_inv, ct_items.

  LOOP AT gt_xl INTO DATA(ls_xl) WHERE packno = iv_packno.

    " Header fields: first filled value of the packing list
    DO.
      ASSIGN COMPONENT sy-index OF STRUCTURE cs_head TO <lv_head>.
      IF sy-subrc <> 0.
        EXIT.
      ENDIF.
      ASSIGN COMPONENT sy-index OF STRUCTURE ls_xl TO <lv_row>.
      IF <lv_head> IS INITIAL AND <lv_row> IS NOT INITIAL.
        <lv_head> = <lv_row>.
      ENDIF.
    ENDDO.

    " Invoice numbers - several in one cell: "A, B & C"
    IF ls_xl-invoices IS NOT INITIAL.
      lv_text = ls_xl-invoices.
      REPLACE ALL OCCURRENCES OF REGEX `[,;&/\s]+` IN lv_text WITH ` `.
      SPLIT condense( lv_text ) AT ` ` INTO TABLE lt_token.
      LOOP AT lt_token INTO DATA(lv_token) WHERE table_line IS NOT INITIAL.
        IF lv_token CO '0123456789' AND strlen( lv_token ) <= 10.
          lv_vbeln = |{ lv_token ALPHA = IN WIDTH = 10 }|.
          INSERT lv_vbeln INTO TABLE ct_inv.
        ELSE.
          APPEND VALUE #( type = 'W' packno = iv_packno
                          text = |Row { ls_xl-row }: "{ lv_token }" is not an invoice number - ignored| ) TO gt_log.
        ENDIF.
      ENDLOOP.
    ENDIF.

    " Line items
    IF ls_xl-sn IS NOT INITIAL.
      APPEND ls_xl TO ct_items.
    ENDIF.

  ENDLOOP.

  IF ct_inv IS INITIAL.
    APPEND VALUE #( type = 'E' packno = iv_packno text = 'No SAP invoice number in the file' ) TO gt_log.
  ENDIF.

ENDFORM.

*&---------------------------------------------------------------------*
*& Header data of one packing list (logic sheet B3 - B44, B61)
*&---------------------------------------------------------------------*
FORM pl_header USING    iv_packno TYPE string
                        is_head   TYPE ty_xl
                        it_inv    TYPE ty_vbeln_tab
               CHANGING cs_out    TYPE ty_out.

  DATA ls_bill    TYPE ty_bill.
  DATA ls_party   TYPE ty_party.
  DATA lt_orders  TYPE ty_string_tab.
  DATA lt_ord_no  TYPE ty_string_tab.
  DATA lt_ord_dt  TYPE ty_string_tab.
  DATA lt_goods   TYPE ty_string_tab.
  DATA lt_inv_ext TYPE ty_string_tab.
  DATA lt_steuc   TYPE SORTED TABLE OF marc-steuc WITH UNIQUE KEY table_line.
  DATA lv_text    TYPE string.
  DATA lv_ok      TYPE abap_bool.

  CLEAR cs_out.

  " Header values from VBRK: first billing document of the packing list
  " (A5). Billing items of all its billing documents for B13 / B42.
  READ TABLE it_inv INTO DATA(lv_first) INDEX 1.
  READ TABLE gt_bill INTO ls_bill WITH KEY vbeln = lv_first.

  LOOP AT it_inv INTO DATA(lv_inv).
    APPEND |{ lv_inv ALPHA = OUT }| TO lt_inv_ext.
    LOOP AT gt_bill INTO DATA(ls_item) WHERE vbeln = lv_inv.
      " Sales orders of the packing list
      lv_text = |{ ls_item-aubel ALPHA = OUT }|.
      IF NOT line_exists( lt_orders[ table_line = lv_text ] ).
        APPEND lv_text TO lt_orders.
        " B13: AUBEL -> VBAK-VGBEL -> VBKD-BSTKD / BSTDK
        READ TABLE gt_vbak INTO DATA(ls_vbak) WITH TABLE KEY vbeln = ls_item-aubel.
        IF sy-subrc = 0.
          READ TABLE gt_vbkd INTO DATA(ls_vbkd) WITH TABLE KEY vbeln = ls_vbak-vgbel.
          IF sy-subrc = 0.
            IF ls_vbkd-bstkd IS NOT INITIAL AND NOT line_exists( lt_ord_no[ table_line = ls_vbkd-bstkd ] ).
              APPEND CONV string( ls_vbkd-bstkd ) TO lt_ord_no.
            ENDIF.
            IF ls_vbkd-bstdk IS NOT INITIAL.
              lv_text = |{ ls_vbkd-bstdk+6(2) }.{ ls_vbkd-bstdk+4(2) }.{ ls_vbkd-bstdk(4) }|.
              IF NOT line_exists( lt_ord_dt[ table_line = lv_text ] ).
                APPEND lv_text TO lt_ord_dt.
              ENDIF.
            ENDIF.
          ELSE.
            APPEND VALUE #( type = 'W' packno = iv_packno
                            text = |Sales order { ls_item-aubel ALPHA = OUT }: no quotation reference (VBAK-VGBEL / VBKD) - order no. & date empty| ) TO gt_log.
          ENDIF.
        ENDIF.
      ENDIF.
      " B42: VBRP-MATNR -> MARC-STEUC -> T604N-TEXT1, common HSN clubbed
      READ TABLE gt_marc INTO DATA(ls_marc) WITH TABLE KEY matnr = ls_item-matnr werks = ls_item-werks.
      IF sy-subrc = 0 AND ls_marc-steuc IS NOT INITIAL.
        INSERT ls_marc-steuc INTO TABLE lt_steuc.
        IF sy-subrc = 0.                               " new HSN only
          READ TABLE gt_t604n INTO DATA(ls_t604n) WITH TABLE KEY steuc = ls_marc-steuc.
          IF sy-subrc = 0.
            APPEND |{ ls_t604n-text1 } - { ls_marc-steuc }| TO lt_goods.
          ELSE.
            APPEND |{ ls_marc-steuc }| TO lt_goods.
          ENDIF.
        ENDIF.
      ENDIF.
    ENDLOOP.
  ENDLOOP.

  cs_out-sales_orders = concat_lines_of( table = lt_orders  sep = `, ` ).
  cs_out-invoices     = concat_lines_of( table = lt_inv_ext sep = `, ` ).

  " ---- Exporter (B3 - B7, B11) ------------------------------------------
  READ TABLE gt_t001 INTO DATA(ls_t001) WITH TABLE KEY bukrs = ls_bill-bukrs.
  READ TABLE gt_adrc INTO DATA(ls_adrc) WITH TABLE KEY addrnumber = ls_t001-adrnr.
  IF sy-subrc = 0.
    cs_out-exp_name1     = ls_adrc-name1.
    cs_out-exp_street    = ls_adrc-street.
    cs_out-exp_str_supl1 = ls_adrc-str_suppl1.
    cs_out-exp_city1     = ls_adrc-city1.
    cs_out-exp_post_code = ls_adrc-post_code1.
    cs_out-exp_region    = ls_adrc-region.
    cs_out-exp_country   = ls_adrc-country.
    cs_out-iec           = ls_adrc-building.                        " B11
    READ TABLE gt_t005u INTO DATA(ls_t005u)
         WITH TABLE KEY land1 = ls_adrc-country bland = ls_adrc-region.
    IF sy-subrc = 0.
      cs_out-exp_region_tx = ls_t005u-bezei.                         " B4
    ENDIF.
    READ TABLE gt_t005t INTO DATA(ls_t005t) WITH TABLE KEY land1 = ls_adrc-country.
    IF sy-subrc = 0.
      cs_out-exp_country_tx = ls_t005t-landx.                        " B5
    ENDIF.
  ELSE.
    APPEND VALUE #( type = 'W' packno = iv_packno
                    text = |Company code { ls_bill-bukrs }: no address (T001 / ADRC)| ) TO gt_log.
  ENDIF.

  READ TABLE gt_branch INTO DATA(ls_branch)
       WITH TABLE KEY bukrs = ls_bill-bukrs branch = ls_bill-bupla.
  IF sy-subrc = 0.
    cs_out-gstin = ls_branch-gstin.                                  " B6
  ELSE.
    APPEND VALUE #( type = 'W' packno = iv_packno
                    text = |No GST No. in J_1BBRANCH for { ls_bill-bukrs } / { ls_bill-bupla }| ) TO gt_log.
  ENDIF.

  READ TABLE gt_t001z INTO DATA(ls_t001z) WITH TABLE KEY bukrs = ls_bill-bukrs.
  IF sy-subrc = 0.
    cs_out-cin = ls_t001z-paval.                                     " B7
  ELSE.
    APPEND VALUE #( type = 'W' packno = iv_packno
                    text = |No CIN in T001Z (party ZCIN) for company code { ls_bill-bukrs }| ) TO gt_log.
  ENDIF.

  " ---- ZSD_PACKNO header fields (B9, B10, B14, B16) ----------------------
  cs_out-packno = iv_packno.
  PERFORM to_date USING is_head-packdt iv_packno 'Packing list date' CHANGING cs_out-packdt lv_ok.
  cs_out-order_no   = concat_lines_of( table = lt_ord_no sep = ` & ` ).
  cs_out-order_date = concat_lines_of( table = lt_ord_dt sep = ` & ` ).
  cs_out-buyer_ref  = is_head-party_ref.
  PERFORM to_date USING is_head-ref_dt iv_packno 'Buyer''s reference date' CHANGING cs_out-buyer_ref_dt lv_ok.
  cs_out-cnty_orgn  = is_head-cnty_orgn.

  " ---- Sold-to (B20 - B23) and ship-to (B26 - B29) -----------------------
  PERFORM partner USING ls_bill gc_parvw_sp iv_packno CHANGING ls_party.
  cs_out-sp_kunnr = ls_party-kunnr.
  READ TABLE gt_adrc INTO ls_adrc WITH TABLE KEY addrnumber = ls_party-adrnr.
  IF sy-subrc = 0.
    cs_out-sp_name1     = ls_adrc-name1.
    cs_out-sp_street    = ls_adrc-street.
    cs_out-sp_str_supl1 = ls_adrc-str_suppl1.
    cs_out-sp_city1     = ls_adrc-city1.
    cs_out-sp_city2     = ls_adrc-city2.
    cs_out-sp_post_code = ls_adrc-post_code1.
    cs_out-sp_country   = ls_adrc-country.
    READ TABLE gt_t005t INTO ls_t005t WITH TABLE KEY land1 = ls_adrc-country.
    IF sy-subrc = 0.
      cs_out-sp_country_tx = ls_t005t-landx.
    ENDIF.
  ENDIF.
  PERFORM contact USING ls_party-adrnr
                  CHANGING cs_out-sp_email cs_out-sp_telephone cs_out-sp_mobile.

  PERFORM partner USING ls_bill gc_parvw_sh iv_packno CHANGING ls_party.
  cs_out-sh_kunnr = ls_party-kunnr.
  READ TABLE gt_adrc INTO ls_adrc WITH TABLE KEY addrnumber = ls_party-adrnr.
  IF sy-subrc = 0.
    cs_out-sh_name1     = ls_adrc-name1.
    cs_out-sh_street    = ls_adrc-street.
    cs_out-sh_str_supl1 = ls_adrc-str_suppl1.
    cs_out-sh_city1     = ls_adrc-city1.
    cs_out-sh_city2     = ls_adrc-city2.
    cs_out-sh_post_code = ls_adrc-post_code1.
    cs_out-sh_country   = ls_adrc-country.
    READ TABLE gt_t005t INTO ls_t005t WITH TABLE KEY land1 = ls_adrc-country.
    IF sy-subrc = 0.
      cs_out-sh_country_tx = ls_t005t-landx.
      cs_out-cnty_dest     = ls_t005t-landx.                        " B17
    ENDIF.
  ENDIF.
  PERFORM contact USING ls_party-adrnr
                  CHANGING cs_out-sh_email cs_out-sh_telephone cs_out-sh_mobile.

  " ---- Notifier parties and shipment (B31 - B41) - ZSD_PACKNO ------------
  cs_out-np1           = is_head-np1.
  cs_out-np2           = is_head-np2.
  cs_out-np3           = is_head-np3.
  cs_out-np4           = is_head-np4.
  cs_out-np5           = is_head-np5.
  cs_out-pre_carriage  = is_head-prec.
  cs_out-place_receipt = is_head-rec_pc.
  cs_out-vessel_flight = is_head-vsl_flt.
  cs_out-port_loading  = is_head-pol.
  cs_out-port_disch    = is_head-pod.
  cs_out-place_deliv   = is_head-pld.
  cs_out-marks_nos     = is_head-mark.
  cs_out-container_no  = is_head-cont.
  cs_out-no_packages   = is_head-nopack.

  " ---- Description of goods (B42), payment (B43), shipment terms (B44) ---
  cs_out-goods_desc = concat_lines_of( table = lt_goods sep = ` / ` ).
  IF cs_out-goods_desc IS INITIAL.
    APPEND VALUE #( type = 'W' packno = iv_packno
                    text = 'No HSN code (MARC-STEUC) for the billed materials - description of goods empty' ) TO gt_log.
  ENDIF.

  cs_out-zterm = ls_bill-zterm.
  LOOP AT gt_tvzbt INTO DATA(ls_tvzbt) WHERE zterm = ls_bill-zterm.
    cs_out-payment_term = ls_tvzbt-vtext.                  " lowest ZTAGG
    EXIT.
  ENDLOOP.
  cs_out-inco1 = ls_bill-inco1.
  cs_out-inco2 = ls_bill-inco2.

  " ---- Footer (B61) - ZSD_PACKNO --------------------------------------------
  cs_out-declaration = is_head-decl.

ENDFORM.

*&---------------------------------------------------------------------*
*& Line item (B47 - B58) - ZSD_PACKNO values; product description and
*& HSN from the material master when not given in the file
*&---------------------------------------------------------------------*
FORM pl_item USING    iv_packno TYPE string
                      is_item   TYPE ty_xl
                      it_inv    TYPE ty_vbeln_tab
             CHANGING cs_out    TYPE ty_out.

  DATA lv_matnr TYPE matnr.
  DATA lv_werks TYPE werks_d.
  DATA lv_text  TYPE string.

  PERFORM to_int USING is_item-sn     iv_packno 'Sl.No.'     CHANGING cs_out-sn.
  PERFORM to_int USING is_item-inv_sn iv_packno 'Sr.No. INV' CHANGING cs_out-inv_sn.
  cs_out-article_no = is_item-artno.
  PERFORM to_int USING is_item-std    iv_packno 'STD'        CHANGING cs_out-std.
  PERFORM to_int USING is_item-art    iv_packno 'Articles'   CHANGING cs_out-articles.
  " B52 ZSD_TART: no column in the Excel file - left empty (A7)
  cs_out-part_no = is_item-partno.
  PERFORM to_dec USING is_item-qty iv_packno 'Quantity'  CHANGING cs_out-quantity.
  PERFORM to_dec USING is_item-gwt iv_packno 'Gross Wt.' CHANGING cs_out-gross_wt.
  PERFORM to_dec USING is_item-nwt iv_packno 'Net Wt.'   CHANGING cs_out-net_wt.

  " Material of the Part No. and its plant in the packing list invoices
  READ TABLE gt_partno INTO DATA(ls_partno) WITH TABLE KEY partno = is_item-partno.
  IF sy-subrc = 0.
    lv_matnr = ls_partno-matnr.
  ENDIF.
  IF lv_matnr IS NOT INITIAL.
    LOOP AT it_inv INTO DATA(lv_inv).
      LOOP AT gt_bill INTO DATA(ls_bill) WHERE vbeln = lv_inv AND matnr = lv_matnr.
        lv_werks = ls_bill-werks.
        EXIT.
      ENDLOOP.
      IF lv_werks IS NOT INITIAL.
        EXIT.
      ENDIF.
    ENDLOOP.
  ENDIF.

  " B54: ZSD_MATDESC or material description
  cs_out-prod_desc = is_item-matdesc.
  IF cs_out-prod_desc IS INITIAL.
    READ TABLE gt_makt INTO DATA(ls_makt) WITH TABLE KEY matnr = lv_matnr.
    IF sy-subrc = 0.
      cs_out-prod_desc = ls_makt-maktx.
    ELSE.
      lv_text = |Sl.No. { is_item-sn }: no product description in the file or the material master|.
      APPEND VALUE #( type = 'W' packno = iv_packno text = lv_text ) TO gt_log.
    ENDIF.
  ENDIF.

  " B58: ZSD_HSN or MARC-STEUC of the Part No. (plant of the billing item)
  cs_out-hsn = is_item-hsn.
  IF cs_out-hsn IS INITIAL.
    READ TABLE gt_marc INTO DATA(ls_marc) WITH TABLE KEY matnr = lv_matnr werks = lv_werks.
    IF sy-subrc = 0 AND ls_marc-steuc IS NOT INITIAL.
      cs_out-hsn = ls_marc-steuc.
    ELSE.
      lv_text = |Sl.No. { is_item-sn }: part no. { is_item-partno } has no HSN in the file and is not billed in the packing list invoices (MARC)|.
      APPEND VALUE #( type = 'W' packno = iv_packno text = lv_text ) TO gt_log.
    ENDIF.
  ENDIF.

ENDFORM.

*&---------------------------------------------------------------------*
*& KNVP partner of the billing document: KUNNR = KUNAG, sales area of the
*& billing document, PARVW = SP / SH -> KUNN2 -> KNA1-ADRNR. Several
*& partners: the lowest counter PARZA (A4).
*&---------------------------------------------------------------------*
FORM partner USING    is_bill   TYPE ty_bill
                      iv_parvw  TYPE parvw
                      iv_packno TYPE string
             CHANGING cs_party  TYPE ty_party.

  CLEAR cs_party.

  READ TABLE gt_knvp INTO DATA(ls_knvp)
       WITH KEY kunnr = is_bill-kunag
                vkorg = is_bill-vkorg
                vtweg = is_bill-vtweg
                spart = is_bill-spart
                parvw = iv_parvw.
  IF sy-subrc <> 0.
    APPEND VALUE #( type = 'W' packno = iv_packno
                    text = |Customer { is_bill-kunag ALPHA = OUT }: no partner { iv_parvw } in KNVP for { is_bill-vkorg }/{ is_bill-vtweg }/{ is_bill-spart }| ) TO gt_log.
    RETURN.
  ENDIF.

  cs_party-kunnr = ls_knvp-kunn2.
  cs_party-adrnr = ls_knvp-adrnr.

ENDFORM.

*&---------------------------------------------------------------------*
*& E-mail (ADR6), telephone (ADR2 R3_USER 1), mobile (ADR2 R3_USER 3)
*& Default entry first, else the first one.
*&---------------------------------------------------------------------*
FORM contact USING    iv_adrnr TYPE adrnr
             CHANGING cv_email TYPE adr6-smtp_addr
                      cv_tel   TYPE adr2-telnr_long
                      cv_mob   TYPE adr2-telnr_long.

  CLEAR: cv_email, cv_tel, cv_mob.
  IF iv_adrnr IS INITIAL.
    RETURN.
  ENDIF.

  READ TABLE gt_adr6 INTO DATA(ls_adr6) WITH KEY addrnumber = iv_adrnr BINARY SEARCH.
  IF sy-subrc = 0.
    cv_email = ls_adr6-smtp_addr.
  ENDIF.

  READ TABLE gt_adr2 INTO DATA(ls_adr2) WITH KEY addrnumber = iv_adrnr r3_user = gc_tel BINARY SEARCH.
  IF sy-subrc = 0.
    cv_tel = ls_adr2-telnr_long.
  ENDIF.

  READ TABLE gt_adr2 INTO ls_adr2 WITH KEY addrnumber = iv_adrnr r3_user = gc_mob BINARY SEARCH.
  IF sy-subrc = 0.
    cv_mob = ls_adr2-telnr_long.
  ENDIF.

ENDFORM.

*&---------------------------------------------------------------------*
*& Conversions of Excel values (with a warning when not convertible)
*&---------------------------------------------------------------------*
FORM to_int USING    iv_text   TYPE string
                     iv_packno TYPE string
                     iv_field  TYPE csequence
            CHANGING cv_value  TYPE i.

  DATA lv_dec TYPE p LENGTH 16 DECIMALS 3.

  CLEAR cv_value.
  PERFORM to_dec USING iv_text iv_packno iv_field CHANGING lv_dec.
  cv_value = round( val = lv_dec dec = 0 ).

ENDFORM.

FORM to_dec USING    iv_text   TYPE string
                     iv_packno TYPE string
                     iv_field  TYPE csequence
            CHANGING cv_value  TYPE p.

  DATA lv_text TYPE string.
  DATA lv_num  TYPE decfloat34.

  CLEAR cv_value.
  lv_text = condense( val = iv_text del = ` ` ).
  REPLACE ALL OCCURRENCES OF `,` IN lv_text WITH ``.
  IF lv_text IS INITIAL.
    RETURN.
  ENDIF.

  TRY.
      lv_num   = lv_text.
      cv_value = lv_num.
    CATCH cx_sy_conversion_error cx_sy_arithmetic_error.
      APPEND VALUE #( type = 'W' packno = iv_packno
                      text = |{ iv_field }: "{ iv_text }" is not a number - left empty| ) TO gt_log.
  ENDTRY.

ENDFORM.

FORM to_date USING    iv_text   TYPE string
                      iv_packno TYPE string
                      iv_field  TYPE csequence
             CHANGING cv_date   TYPE d
                      cv_ok     TYPE abap_bool.

  DATA lv_text TYPE string.
  DATA lv_day  TYPE string.
  DATA lv_mon  TYPE string.
  DATA lv_year TYPE string.
  DATA lv_num  TYPE decfloat34.
  DATA lv_days TYPE i.

  CLEAR cv_date.
  cv_ok = abap_true.
  lv_text = condense( val = iv_text del = ` ` ).
  IF lv_text IS INITIAL.
    RETURN.
  ENDIF.

  cv_ok = abap_false.
  " "2026-09-10 00:00:00" -> date part only
  lv_text = segment( val = lv_text index = 1 sep = ` ` ).

  IF matches( val = lv_text regex = `\d+(\.\d+)?` ).
    " Excel date number: days since 30.12.1899
    TRY.
        lv_num = lv_text.
      CATCH cx_sy_conversion_error.
        lv_num = 0.
    ENDTRY.
    IF lv_num >= 1 AND lv_num <= 2958465.
      lv_days = trunc( lv_num ).
      cv_date = '18991230'.
      cv_date = cv_date + lv_days.
    ENDIF.
  ELSE.
    " DD.MM.YYYY / DD-MM-YYYY / DD/MM/YYYY / YYYY-MM-DD
    REPLACE ALL OCCURRENCES OF REGEX `[/-]` IN lv_text WITH `.`.
    SPLIT lv_text AT '.' INTO lv_day lv_mon lv_year.
    IF strlen( lv_day ) = 4.
      lv_text = lv_day.
      lv_day  = lv_year.
      lv_year = lv_text.
    ENDIF.
    IF strlen( lv_year ) = 4 AND lv_day IS NOT INITIAL AND lv_mon IS NOT INITIAL.
      cv_date = |{ lv_year }{ lv_mon WIDTH = 2 ALIGN = RIGHT PAD = '0' }{ lv_day WIDTH = 2 ALIGN = RIGHT PAD = '0' }|.
    ENDIF.
  ENDIF.

  IF cv_date IS NOT INITIAL.
    CALL FUNCTION 'DATE_CHECK_PLAUSIBILITY'
      EXPORTING
        date                      = cv_date
      EXCEPTIONS
        plausibility_check_failed = 1
        OTHERS                    = 2.
    IF sy-subrc = 0.
      cv_ok = abap_true.
      RETURN.
    ENDIF.
  ENDIF.

  CLEAR cv_date.
  APPEND VALUE #( type = 'W' packno = iv_packno
                  text = |{ iv_field }: "{ iv_text }" is not a date - left empty| ) TO gt_log.

ENDFORM.

*&---------------------------------------------------------------------*
*& 5. Messages (popup) and final table (ALV grid, export to spreadsheet
*&    via the standard ALV functions)
*&---------------------------------------------------------------------*
FORM display.

  DATA lo_alv  TYPE REF TO cl_salv_table.
  DATA lo_log  TYPE REF TO cl_salv_table.
  DATA lx_salv TYPE REF TO cx_salv_error.

  " Messages first
  IF gt_log IS NOT INITIAL.
    TRY.
        cl_salv_table=>factory( IMPORTING r_salv_table = lo_log
                                CHANGING  t_table      = gt_log ).
        lo_log->set_screen_popup( start_column = 5  end_column = 150
                                  start_line   = 3  end_line   = 20 ).
        lo_log->get_columns( )->set_optimize( abap_true ).
        PERFORM column_text USING lo_log 'TYPE'   'Type'.
        PERFORM column_text USING lo_log 'PACKNO' 'Packing list'.
        PERFORM column_text USING lo_log 'TEXT'   'Message'.
        lo_log->display( ).
      CATCH cx_salv_error INTO lx_salv.
        MESSAGE lx_salv TYPE 'S' DISPLAY LIKE 'E'.
    ENDTRY.
  ENDIF.

  IF gt_out IS INITIAL.
    MESSAGE 'No data to display' TYPE 'S' DISPLAY LIKE 'E'.
    RETURN.
  ENDIF.

  TRY.
      cl_salv_table=>factory( IMPORTING r_salv_table = lo_alv
                              CHANGING  t_table      = gt_out ).
      " All standard functions incl. export to spreadsheet
      lo_alv->get_functions( )->set_all( abap_true ).
      lo_alv->get_columns( )->set_optimize( abap_true ).
      lo_alv->get_display_settings( )->set_striped_pattern( abap_true ).
      lo_alv->get_display_settings( )->set_list_header( 'Export Packing List' ).
      PERFORM column_texts USING lo_alv.
      lo_alv->display( ).
    CATCH cx_salv_error INTO lx_salv.
      MESSAGE lx_salv TYPE 'S' DISPLAY LIKE 'E'.
  ENDTRY.

ENDFORM.

*&---------------------------------------------------------------------*
*& Column headings of the output (logic sheet labels)
*&---------------------------------------------------------------------*
FORM column_texts USING io_alv TYPE REF TO cl_salv_table.

  PERFORM column_text USING io_alv 'SALES_ORDERS'   'Sales Order(s)'.
  PERFORM column_text USING io_alv 'INVOICES'       'SAP Invoice Number(s)'.
  PERFORM column_text USING io_alv 'EXP_NAME1'      'Exporter Name'.
  PERFORM column_text USING io_alv 'EXP_STREET'     'Exporter Street'.
  PERFORM column_text USING io_alv 'EXP_STR_SUPL1'  'Exporter Street 2'.
  PERFORM column_text USING io_alv 'EXP_CITY1'      'Exporter City'.
  PERFORM column_text USING io_alv 'EXP_POST_CODE'  'Exporter Postal Code'.
  PERFORM column_text USING io_alv 'EXP_REGION'     'Exporter Region'.
  PERFORM column_text USING io_alv 'EXP_REGION_TX'  'Exporter Region Desc.'.
  PERFORM column_text USING io_alv 'EXP_COUNTRY'    'Exporter Country'.
  PERFORM column_text USING io_alv 'EXP_COUNTRY_TX' 'Exporter Country Desc.'.
  PERFORM column_text USING io_alv 'GSTIN'          'GST No.'.
  PERFORM column_text USING io_alv 'CIN'            'CIN No.'.
  PERFORM column_text USING io_alv 'PACKNO'         'Packing List No.'.
  PERFORM column_text USING io_alv 'PACKDT'         'Packing List Date'.
  PERFORM column_text USING io_alv 'IEC'            'Exporter Ref. (IEC No.)'.
  PERFORM column_text USING io_alv 'ORDER_NO'       'Order No.'.
  PERFORM column_text USING io_alv 'ORDER_DATE'     'Order Date'.
  PERFORM column_text USING io_alv 'BUYER_REF'      'Buyer''s Reference No.'.
  PERFORM column_text USING io_alv 'BUYER_REF_DT'   'Buyer''s Reference Date'.
  PERFORM column_text USING io_alv 'CNTY_ORGN'      'Country of Origin of Goods'.
  PERFORM column_text USING io_alv 'CNTY_DEST'      'Country of Final Destination'.
  PERFORM column_text USING io_alv 'SP_KUNNR'       'Buyer (Sold-to)'.
  PERFORM column_text USING io_alv 'SP_NAME1'       'Buyer Name'.
  PERFORM column_text USING io_alv 'SP_STREET'      'Buyer Street'.
  PERFORM column_text USING io_alv 'SP_STR_SUPL1'   'Buyer Street 2'.
  PERFORM column_text USING io_alv 'SP_CITY1'       'Buyer City'.
  PERFORM column_text USING io_alv 'SP_CITY2'       'Buyer District'.
  PERFORM column_text USING io_alv 'SP_POST_CODE'   'Buyer Postal Code'.
  PERFORM column_text USING io_alv 'SP_COUNTRY'     'Buyer Country'.
  PERFORM column_text USING io_alv 'SP_COUNTRY_TX'  'Buyer Country Desc.'.
  PERFORM column_text USING io_alv 'SP_EMAIL'       'Buyer E-mail'.
  PERFORM column_text USING io_alv 'SP_TELEPHONE'   'Buyer Telephone'.
  PERFORM column_text USING io_alv 'SP_MOBILE'      'Buyer Mobile Number'.
  PERFORM column_text USING io_alv 'SH_KUNNR'       'Ship-to Party'.
  PERFORM column_text USING io_alv 'SH_NAME1'       'Ship-to Name'.
  PERFORM column_text USING io_alv 'SH_STREET'      'Ship-to Street'.
  PERFORM column_text USING io_alv 'SH_STR_SUPL1'   'Ship-to Street 2'.
  PERFORM column_text USING io_alv 'SH_CITY1'       'Ship-to City'.
  PERFORM column_text USING io_alv 'SH_CITY2'       'Ship-to District'.
  PERFORM column_text USING io_alv 'SH_POST_CODE'   'Ship-to Postal Code'.
  PERFORM column_text USING io_alv 'SH_COUNTRY'     'Ship-to Country'.
  PERFORM column_text USING io_alv 'SH_COUNTRY_TX'  'Ship-to Country Desc.'.
  PERFORM column_text USING io_alv 'SH_EMAIL'       'Ship-to E-mail'.
  PERFORM column_text USING io_alv 'SH_TELEPHONE'   'Ship-to Telephone'.
  PERFORM column_text USING io_alv 'SH_MOBILE'      'Ship-to Mobile Number'.
  PERFORM column_text USING io_alv 'NP1'            'Notifier Party 1'.
  PERFORM column_text USING io_alv 'NP2'            'Notifier Party 2'.
  PERFORM column_text USING io_alv 'NP3'            'Notifier Party 3'.
  PERFORM column_text USING io_alv 'NP4'            'Notifier Party 4'.
  PERFORM column_text USING io_alv 'NP5'            'Notifier Party 5'.
  PERFORM column_text USING io_alv 'PRE_CARRIAGE'   'Pre-Carriage'.
  PERFORM column_text USING io_alv 'PLACE_RECEIPT'  'Place of Receipt by Pre-carrier'.
  PERFORM column_text USING io_alv 'VESSEL_FLIGHT'  'Vessel / Flight No.'.
  PERFORM column_text USING io_alv 'PORT_LOADING'   'Port of Loading'.
  PERFORM column_text USING io_alv 'PORT_DISCH'     'Port of Discharge'.
  PERFORM column_text USING io_alv 'PLACE_DELIV'    'Place of Delivery'.
  PERFORM column_text USING io_alv 'MARKS_NOS'      'Marks & Nos.'.
  PERFORM column_text USING io_alv 'CONTAINER_NO'   'Container No.'.
  PERFORM column_text USING io_alv 'NO_PACKAGES'    'No. of Packages'.
  PERFORM column_text USING io_alv 'GOODS_DESC'     'Description of Goods'.
  PERFORM column_text USING io_alv 'ZTERM'          'Payment Term'.
  PERFORM column_text USING io_alv 'PAYMENT_TERM'   'Payment Term Desc.'.
  PERFORM column_text USING io_alv 'INCO1'          'Incoterms'.
  PERFORM column_text USING io_alv 'INCO2'          'Incoterms Location'.
  PERFORM column_text USING io_alv 'SN'             'Sl.No.'.
  PERFORM column_text USING io_alv 'INV_SN'         'Sr.No. INV'.
  PERFORM column_text USING io_alv 'ARTICLE_NO'     'Article No.'.
  PERFORM column_text USING io_alv 'STD'            'STD'.
  PERFORM column_text USING io_alv 'ARTICLES'       'Articles'.
  PERFORM column_text USING io_alv 'TOTAL_ART'      'Total Articles'.
  PERFORM column_text USING io_alv 'PART_NO'        'Part No.'.
  PERFORM column_text USING io_alv 'PROD_DESC'      'Product Description'.
  PERFORM column_text USING io_alv 'QUANTITY'       'Quantity Pcs'.
  PERFORM column_text USING io_alv 'GROSS_WT'       'Gross Wt. KGS'.
  PERFORM column_text USING io_alv 'NET_WT'         'Net Wt. KGS'.
  PERFORM column_text USING io_alv 'HSN'            'H.S Code'.
  PERFORM column_text USING io_alv 'DECLARATION'    'Declaration'.

ENDFORM.

FORM column_text USING io_alv   TYPE REF TO cl_salv_table
                       iv_name  TYPE csequence
                       iv_text  TYPE csequence.

  DATA lo_column TYPE REF TO cl_salv_column.
  DATA lv_name   TYPE lvc_fname.
  DATA lv_long   TYPE scrtext_l.
  DATA lv_medium TYPE scrtext_m.
  DATA lv_short  TYPE scrtext_s.

  TRY.
      lv_name   = iv_name.
      lo_column = io_alv->get_columns( )->get_column( lv_name ).
      lv_long   = iv_text.
      lv_medium = iv_text.
      lv_short  = iv_text.
      lo_column->set_long_text( lv_long ).
      lo_column->set_medium_text( lv_medium ).
      lo_column->set_short_text( lv_short ).
    CATCH cx_salv_not_found.
      RETURN.
  ENDTRY.

ENDFORM.

*&---------------------------------------------------------------------*
*& F4 help for the Excel file
*&---------------------------------------------------------------------*
FORM f4_file CHANGING cv_file TYPE rlgrap-filename.

  DATA lt_files TYPE filetable.
  DATA ls_file  TYPE file_table.
  DATA lv_rc    TYPE i.

  cl_gui_frontend_services=>file_open_dialog(
    EXPORTING
      file_filter    = 'Excel (*.xlsx)|*.xlsx'
      multiselection = abap_false
    CHANGING
      file_table     = lt_files
      rc             = lv_rc
    EXCEPTIONS
      OTHERS         = 1 ).
  IF sy-subrc = 0 AND lv_rc = 1.
    READ TABLE lt_files INTO ls_file INDEX 1.
    cv_file = ls_file-filename.
  ENDIF.

ENDFORM.

*&---------------------------------------------------------------------*
*& Text elements (Goto -> Text elements)
*&   Text symbol    B01      Selection
*&   Selection text S_VBELN  Sales Order
*&   Selection text P_FILE   Packing list Excel file (.xlsx)
*&---------------------------------------------------------------------*
