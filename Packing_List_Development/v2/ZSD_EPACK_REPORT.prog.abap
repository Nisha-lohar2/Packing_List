*&---------------------------------------------------------------------*
*& Report ZSD_EPACK_REPORT
*&---------------------------------------------------------------------*
*& WRICEF 102-B  Export consolidated packing list
*&
*& Process (two steps, same program):
*&  STEP 1 - Sales order(s) entered, no file:
*&           the data whose logic is given in the FS logic sheet is read
*&           from SAP and shown in an ALV grid with ALL columns of the
*&           packing list (columns without SAP logic stay empty). One row
*&           per billing document item of the sales orders.
*&           The user exports the ALV to Excel (standard ALV function
*&           "Export -> Spreadsheet") and fills the remaining columns
*&           (packing list no., notifiers, shipment data, line items ...).
*&  STEP 2 - The filled Excel file is uploaded (file field):
*&           the file is read back into the same structure (columns found
*&           by their headings) and the Adobe Form ZSD_EPACK_AF is printed
*&           once per packing list number with exactly this data.
*&
*& Relationship sales order -> billing document: VBRP-AUBEL = VBAK-VBELN.
*& All database reads are bulk reads (JOIN / FOR ALL ENTRIES) before the
*& output is built. No SELECT and no SORT inside a loop.
*& Assumptions: v2/Assumptions_v2.md
*&---------------------------------------------------------------------*
REPORT zsd_epack_report.

*----------------------------------------------------------------------*
* Constants
*----------------------------------------------------------------------*
CONSTANTS:
  gc_langu    TYPE spras VALUE 'E',              " T604N "SPRAS = EN", texts
  gc_land_in  TYPE land1 VALUE 'IN',             " T604N "LAND1 = IN"
  gc_cin      TYPE t001z-party VALUE 'ZCIN',     " T001Z "PARTY = ZCIN"
  " Logic sheet "PARVW = SP / SH" = sold-to / ship-to party. Stored in
  " KNVP with the internal codes AG / WE (A4).
  gc_parvw_sp TYPE parvw VALUE 'AG',
  gc_parvw_sh TYPE parvw VALUE 'WE',
  gc_tel      TYPE adr2-r3_user VALUE '1',       " ADR2-R3_USER telephone
  gc_mob      TYPE adr2-r3_user VALUE '3',       " ADR2-R3_USER mobile
  " Step 2: Adobe Form and output device (print preview, printable)
  gc_form     TYPE fpname    VALUE 'ZSD_EPACK_AF',
  gc_dest     TYPE rspopname VALUE 'LP01'.

*----------------------------------------------------------------------*
* Types - packing list data = ALV columns = Excel columns (in this order)
*----------------------------------------------------------------------*
TYPES:
  BEGIN OF ty_out,
    " Reference
    sales_orders  TYPE string,             " sales order of the row
    invoices      TYPE string,             " billing document of the row
    " Header - Exporter (B3-B5), ONE column, fixed positions (A17):
    " name, street, street 2, city, postal code, region text, country text
    exporter      TYPE c LENGTH 100,
    gstin         TYPE j_1bbranch-gstin,
    cin           TYPE t001z-paval,
    " Header (B9-B17)
    packno        TYPE string,             " ZSD_PACKNO - filled by user
    packdt        TYPE d,                  " ZSD_PACKDT - filled by user
    iec           TYPE adrc-building,
    order_no      TYPE string,
    order_date    TYPE string,
    buyer_ref     TYPE string,             " ZSD_PARTY_REF - by user
    buyer_ref_dt  TYPE d,                  " ZSD_REF_DT    - by user
    cnty_orgn     TYPE string,             " ZSD_CNTY_ORGN - by user
    cnty_dest     TYPE t005t-landx,
    " Sold-to / bill-to party (B20-B23)
    sp_kunnr      TYPE kunnr,
    " ONE column, fixed positions (A19): name, street, street 2, city,
    " district, postal code, country text
    buyer         TYPE c LENGTH 100,
    sp_email      TYPE adr6-smtp_addr,
    sp_telephone  TYPE adr2-telnr_long,
    sp_mobile     TYPE adr2-telnr_long,
    " Ship-to party (B26-B29)
    sh_kunnr      TYPE kunnr,
    " ONE column, fixed positions (A19): name, street, street 2, city,
    " district, postal code, country text
    ship_to       TYPE c LENGTH 100,
    sh_email      TYPE adr6-smtp_addr,
    sh_telephone  TYPE adr2-telnr_long,
    sh_mobile     TYPE adr2-telnr_long,
    " Notifier parties (B31) - by user
    np1           TYPE string,
    np2           TYPE string,
    np3           TYPE string,
    np4           TYPE string,
    np5           TYPE string,
    " Shipment (B33-B41) - by user
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
    " Line item (B47-B58) - ZSD_PACKNO fields by user; part no.,
    " description and HSN pre-filled from the billing item
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
    " Footer (B61) - by user
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

  " One payment term can have several lines (ZTAGG) - non-unique key
  BEGIN OF ty_tvzbt,
    zterm TYPE tvzbt-zterm,
    vtext TYPE tvzbt-vtext,
  END OF ty_tvzbt,
  ty_tvzbt_tab TYPE SORTED TABLE OF ty_tvzbt WITH NON-UNIQUE KEY zterm,

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

  " Partner and its address number
  BEGIN OF ty_party,
    kunnr TYPE kunnr,
    adrnr TYPE adrnr,
  END OF ty_party,

  " Column heading <-> field of TY_OUT
  BEGIN OF ty_label,
    name TYPE lvc_fname,
    text TYPE string,
    len  TYPE lvc_outlen,              " column width in ALV / Excel
  END OF ty_label,
  ty_label_tab TYPE STANDARD TABLE OF ty_label WITH EMPTY KEY,

  BEGIN OF ty_colmap,
    key  TYPE string,                  " normalized heading
    name TYPE lvc_fname,               " '*' = ambiguous heading
  END OF ty_colmap,
  ty_colmap_tab TYPE HASHED TABLE OF ty_colmap WITH UNIQUE KEY key,

  " Message log
  BEGIN OF ty_log,
    type TYPE symsgty,
    ref  TYPE string,                  " sales order / invoice / PL no.
    text TYPE string,
  END OF ty_log,
  ty_log_tab TYPE STANDARD TABLE OF ty_log WITH EMPTY KEY,

  ty_string_tab TYPE STANDARD TABLE OF string WITH EMPTY KEY.

*----------------------------------------------------------------------*
* Global data
*----------------------------------------------------------------------*
DATA gv_vbeln TYPE vbak-vbeln.                       " for SELECT-OPTIONS

DATA gt_out    TYPE ty_out_tab.                      " packing list data
DATA gt_log    TYPE ty_log_tab.                      " messages
DATA gt_labels TYPE ty_label_tab.                    " column headings

DATA gt_vbak   TYPE ty_vbak_tab.
DATA gt_bill   TYPE ty_bill_tab.
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

*----------------------------------------------------------------------*
* Selection screen
*   Step 1: sales order(s) only       -> ALV with all columns
*   Step 2: filled Excel file         -> Adobe Form per packing list
*----------------------------------------------------------------------*
SELECTION-SCREEN BEGIN OF BLOCK b1 WITH FRAME TITLE TEXT-b01.
SELECT-OPTIONS s_vbeln FOR gv_vbeln.
PARAMETERS     p_file  TYPE rlgrap-filename LOWER CASE.
SELECTION-SCREEN END OF BLOCK b1.

*----------------------------------------------------------------------*
INITIALIZATION.
*----------------------------------------------------------------------*
  PERFORM labels.

*----------------------------------------------------------------------*
AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_file.
*----------------------------------------------------------------------*
  PERFORM f4_file CHANGING p_file.

*----------------------------------------------------------------------*
AT SELECTION-SCREEN.
*----------------------------------------------------------------------*
  IF ( sy-ucomm = 'ONLI' OR sy-ucomm = 'PRIN' )
     AND s_vbeln[] IS INITIAL AND p_file IS INITIAL.
    MESSAGE 'Enter sales order(s) to fetch the data, or the filled Excel file to print'
      TYPE 'E'.
  ENDIF.

*----------------------------------------------------------------------*
START-OF-SELECTION.
*----------------------------------------------------------------------*
  IF p_file IS INITIAL.
    " STEP 1 - data from SAP, all columns in the ALV for download
    PERFORM read_sales_orders.     " VBAK, VBRP/VBRK of the sales orders
    PERFORM read_master_data.      " all other tables, bulk
    PERFORM build_output.          " one row per billing item
    PERFORM show_log.
    PERFORM display.
  ELSE.
    " STEP 2 - filled Excel back in, one form per packing list
    PERFORM read_upload.           " file -> GT_OUT (columns by heading)
    PERFORM print_forms.
    PERFORM show_log.
  ENDIF.

*&=====================================================================*
*& STEP 1 - SAP data
*&=====================================================================*

*&---------------------------------------------------------------------*
*& Sales orders -> billing documents (VBRP-AUBEL = sales order)
*&---------------------------------------------------------------------*
FORM read_sales_orders.

  DATA lv_text   TYPE string.
  DATA lt_billed TYPE SORTED TABLE OF vbak-vbeln WITH UNIQUE KEY table_line.

  IF s_vbeln[] IS INITIAL.
    MESSAGE 'Enter at least one sales order' TYPE 'S' DISPLAY LIKE 'E'.
    LEAVE LIST-PROCESSING.
  ENDIF.

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
      APPEND VALUE #( type = 'E' ref = |{ ls_sel-low ALPHA = OUT }|
                      text = 'Sales order does not exist' ) TO gt_log.
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
    INSERT ls_bill-aubel INTO TABLE lt_billed.
  ENDLOOP.

  " Sales orders without billing document
  LOOP AT gt_vbak INTO DATA(ls_vbak).
    IF NOT line_exists( lt_billed[ table_line = ls_vbak-vbeln ] ).
      lv_text = |{ ls_vbak-vbeln ALPHA = OUT }|.
      APPEND VALUE #( type = 'W' ref = lv_text text = 'Sales order has no billing document' ) TO gt_log.
    ENDIF.
  ENDLOOP.

ENDFORM.

*&---------------------------------------------------------------------*
*& All master data, each table read once (FOR ALL ENTRIES / JOIN)
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
    SELECT zterm, vtext
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
*& ALV rows: one per billing item. SAP fields filled from the logic
*& sheet, ZSD_PACKNO fields empty (filled by the user in Excel).
*&---------------------------------------------------------------------*
FORM build_output.

  DATA ls_head  TYPE ty_out.
  DATA ls_out   TYPE ty_out.
  DATA lt_goods TYPE ty_string_tab.
  DATA lt_steuc TYPE SORTED TABLE OF marc-steuc WITH UNIQUE KEY table_line.
  DATA lv_goods TYPE string.
  DATA lv_prev  TYPE vbrk-vbeln.

  " B42: HSN and description of all billing documents, common HSN once
  LOOP AT gt_bill INTO DATA(ls_bill).
    READ TABLE gt_marc INTO DATA(ls_marc) WITH TABLE KEY matnr = ls_bill-matnr werks = ls_bill-werks.
    IF sy-subrc = 0 AND ls_marc-steuc IS NOT INITIAL.
      INSERT ls_marc-steuc INTO TABLE lt_steuc.
      IF sy-subrc = 0.                                  " new HSN only
        READ TABLE gt_t604n INTO DATA(ls_t604n) WITH TABLE KEY steuc = ls_marc-steuc.
        IF sy-subrc = 0.
          APPEND |{ ls_t604n-text1 } - { ls_marc-steuc }| TO lt_goods.
        ELSE.
          APPEND |{ ls_marc-steuc }| TO lt_goods.
        ENDIF.
      ENDIF.
    ENDIF.
  ENDLOOP.
  lv_goods = concat_lines_of( table = lt_goods sep = ` / ` ).
  IF lv_goods IS INITIAL.
    APPEND VALUE #( type = 'W' text = 'No HSN code (MARC-STEUC) for the billed materials - description of goods empty' ) TO gt_log.
  ENDIF.

  " Header data once per billing document, then its items
  LOOP AT gt_bill INTO ls_bill.
    IF ls_bill-vbeln <> lv_prev.
      PERFORM invoice_header USING ls_bill lv_goods CHANGING ls_head.
      lv_prev = ls_bill-vbeln.
    ENDIF.
    ls_out = ls_head.
    PERFORM bill_item USING ls_bill CHANGING ls_out.
    APPEND ls_out TO gt_out.
  ENDLOOP.

ENDFORM.

*&---------------------------------------------------------------------*
*& Header data of one billing document (logic sheet B3 - B44)
*&---------------------------------------------------------------------*
FORM invoice_header USING    is_bill  TYPE ty_bill
                             iv_goods TYPE string
                    CHANGING cs_out   TYPE ty_out.

  DATA ls_party TYPE ty_party.
  DATA lv_ref   TYPE string.
  DATA lv_reg   TYPE string.
  DATA lv_cty   TYPE string.
  DATA lt_parts TYPE ty_string_tab.

  CLEAR cs_out.
  lv_ref = |{ is_bill-vbeln ALPHA = OUT }|.
  cs_out-invoices = lv_ref.

  " ---- Exporter (B3 - B7, B11) ------------------------------------------
  READ TABLE gt_t001 INTO DATA(ls_t001) WITH TABLE KEY bukrs = is_bill-bukrs.
  IF sy-subrc <> 0.
    CLEAR ls_t001.
  ENDIF.
  READ TABLE gt_adrc INTO DATA(ls_adrc) WITH TABLE KEY addrnumber = ls_t001-adrnr.
  IF sy-subrc = 0.
    cs_out-iec = ls_adrc-building.                                   " B11
    READ TABLE gt_t005u INTO DATA(ls_t005u)
         WITH TABLE KEY land1 = ls_adrc-country bland = ls_adrc-region.
    IF sy-subrc = 0.
      lv_reg = ls_t005u-bezei.                                       " B4
    ENDIF.
    READ TABLE gt_t005t INTO DATA(ls_t005t) WITH TABLE KEY land1 = ls_adrc-country.
    IF sy-subrc = 0.
      lv_cty = ls_t005t-landx.                                       " B5
    ENDIF.
    " B3-B5: NAME1, STREET, STR_SUPPL1, CITY1, POST_CODE1, region, country
    lt_parts = VALUE #( ( CONV string( ls_adrc-name1 ) )
                        ( CONV string( ls_adrc-street ) )
                        ( CONV string( ls_adrc-str_suppl1 ) )
                        ( CONV string( ls_adrc-city1 ) )
                        ( CONV string( ls_adrc-post_code1 ) )
                        ( lv_reg )
                        ( lv_cty ) ).
    PERFORM fixed_address USING lt_parts lv_ref 'Exporter address' CHANGING cs_out-exporter.
  ELSE.
    APPEND VALUE #( type = 'W' ref = lv_ref
                    text = |Company code { is_bill-bukrs }: no address (T001 / ADRC)| ) TO gt_log.
  ENDIF.

  READ TABLE gt_branch INTO DATA(ls_branch)
       WITH TABLE KEY bukrs = is_bill-bukrs branch = is_bill-bupla.
  IF sy-subrc = 0.
    cs_out-gstin = ls_branch-gstin.                                  " B6
  ELSE.
    APPEND VALUE #( type = 'W' ref = lv_ref
                    text = |No GST No. in J_1BBRANCH for { is_bill-bukrs } / { is_bill-bupla }| ) TO gt_log.
  ENDIF.

  READ TABLE gt_t001z INTO DATA(ls_t001z) WITH TABLE KEY bukrs = is_bill-bukrs.
  IF sy-subrc = 0.
    cs_out-cin = ls_t001z-paval.                                     " B7
  ELSE.
    APPEND VALUE #( type = 'W' ref = lv_ref
                    text = |No CIN in T001Z (party ZCIN) for company code { is_bill-bukrs }| ) TO gt_log.
  ENDIF.

  " ---- Sold-to (B20 - B23) ------------------------------------------------
  PERFORM partner USING is_bill gc_parvw_sp lv_ref CHANGING ls_party.
  cs_out-sp_kunnr = ls_party-kunnr.
  READ TABLE gt_adrc INTO ls_adrc WITH TABLE KEY addrnumber = ls_party-adrnr.
  IF sy-subrc = 0.
    CLEAR lv_cty.
    READ TABLE gt_t005t INTO ls_t005t WITH TABLE KEY land1 = ls_adrc-country.
    IF sy-subrc = 0.
      lv_cty = ls_t005t-landx.
    ENDIF.
    " B20: NAME1, STREET, STR_SUPPL1, CITY1, CITY2, POST_CODE1, country text
    lt_parts = VALUE #( ( CONV string( ls_adrc-name1 ) )
                        ( CONV string( ls_adrc-street ) )
                        ( CONV string( ls_adrc-str_suppl1 ) )
                        ( CONV string( ls_adrc-city1 ) )
                        ( CONV string( ls_adrc-city2 ) )
                        ( CONV string( ls_adrc-post_code1 ) )
                        ( lv_cty ) ).
    PERFORM fixed_address USING lt_parts lv_ref 'Buyer address' CHANGING cs_out-buyer.
  ENDIF.
  PERFORM contact USING ls_party-adrnr
                  CHANGING cs_out-sp_email cs_out-sp_telephone cs_out-sp_mobile.

  " ---- Ship-to (B26 - B29), country of final destination (B17) -----------
  PERFORM partner USING is_bill gc_parvw_sh lv_ref CHANGING ls_party.
  cs_out-sh_kunnr = ls_party-kunnr.
  READ TABLE gt_adrc INTO ls_adrc WITH TABLE KEY addrnumber = ls_party-adrnr.
  IF sy-subrc = 0.
    CLEAR lv_cty.
    READ TABLE gt_t005t INTO ls_t005t WITH TABLE KEY land1 = ls_adrc-country.
    IF sy-subrc = 0.
      lv_cty           = ls_t005t-landx.
      cs_out-cnty_dest = ls_t005t-landx.                            " B17
    ENDIF.
    " B26: NAME1, STREET, STR_SUPPL1, CITY1, CITY2, POST_CODE1, country text
    lt_parts = VALUE #( ( CONV string( ls_adrc-name1 ) )
                        ( CONV string( ls_adrc-street ) )
                        ( CONV string( ls_adrc-str_suppl1 ) )
                        ( CONV string( ls_adrc-city1 ) )
                        ( CONV string( ls_adrc-city2 ) )
                        ( CONV string( ls_adrc-post_code1 ) )
                        ( lv_cty ) ).
    PERFORM fixed_address USING lt_parts lv_ref 'Ship-to address' CHANGING cs_out-ship_to.
  ENDIF.
  PERFORM contact USING ls_party-adrnr
                  CHANGING cs_out-sh_email cs_out-sh_telephone cs_out-sh_mobile.

  " ---- Description of goods (B42), payment (B43), shipment terms (B44) ---
  cs_out-goods_desc = iv_goods.
  cs_out-zterm      = is_bill-zterm.
  READ TABLE gt_tvzbt INTO DATA(ls_tvzbt) WITH TABLE KEY zterm = is_bill-zterm.
  IF sy-subrc = 0.
    cs_out-payment_term = ls_tvzbt-vtext.
  ENDIF.
  cs_out-inco1 = is_bill-inco1.
  cs_out-inco2 = is_bill-inco2.

ENDFORM.

*&---------------------------------------------------------------------*
*& Item part of a row: sales order, order no. & date (B13), part no.,
*& description (B54) and HSN (B58) of the billed material
*&---------------------------------------------------------------------*
FORM bill_item USING    is_bill TYPE ty_bill
               CHANGING cs_out  TYPE ty_out.

  DATA lv_partno TYPE c LENGTH 40.

  cs_out-sales_orders = |{ is_bill-aubel ALPHA = OUT }|.

  " B13: AUBEL -> VBAK-VGBEL -> VBKD-BSTKD / BSTDK
  READ TABLE gt_vbak INTO DATA(ls_vbak) WITH TABLE KEY vbeln = is_bill-aubel.
  IF sy-subrc = 0.
    READ TABLE gt_vbkd INTO DATA(ls_vbkd) WITH TABLE KEY vbeln = ls_vbak-vgbel.
    IF sy-subrc = 0.
      cs_out-order_no = ls_vbkd-bstkd.
      IF ls_vbkd-bstdk IS NOT INITIAL.
        cs_out-order_date = |{ ls_vbkd-bstdk+6(2) }.{ ls_vbkd-bstdk+4(2) }.{ ls_vbkd-bstdk(4) }|.
      ENDIF.
    ENDIF.
  ENDIF.

  " Part No. = billed material (logic B58 "from STEP no 53")
  IF is_bill-matnr IS NOT INITIAL.
    CALL FUNCTION 'CONVERSION_EXIT_MATN1_OUTPUT'
      EXPORTING
        input  = is_bill-matnr
      IMPORTING
        output = lv_partno.
    cs_out-part_no = condense( lv_partno ).
  ENDIF.

  READ TABLE gt_makt INTO DATA(ls_makt) WITH TABLE KEY matnr = is_bill-matnr.
  IF sy-subrc = 0.
    cs_out-prod_desc = ls_makt-maktx.                                " B54
  ENDIF.

  READ TABLE gt_marc INTO DATA(ls_marc) WITH TABLE KEY matnr = is_bill-matnr werks = is_bill-werks.
  IF sy-subrc = 0.
    cs_out-hsn = ls_marc-steuc.                                      " B58
  ENDIF.

ENDFORM.

*&---------------------------------------------------------------------*
*& KNVP partner of the billing document: KUNNR = KUNAG, sales area of the
*& billing document, PARVW = SP / SH -> KUNN2 -> KNA1-ADRNR. Several
*& partners: the lowest counter PARZA (A4).
*&---------------------------------------------------------------------*
FORM partner USING    is_bill  TYPE ty_bill
                      iv_parvw TYPE parvw
                      iv_ref   TYPE string
             CHANGING cs_party TYPE ty_party.

  CLEAR cs_party.

  READ TABLE gt_knvp INTO DATA(ls_knvp)
       WITH KEY kunnr = is_bill-kunag
                vkorg = is_bill-vkorg
                vtweg = is_bill-vtweg
                spart = is_bill-spart
                parvw = iv_parvw.
  IF sy-subrc <> 0.
    APPEND VALUE #( type = 'W' ref = iv_ref
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
*& ALV grid with ALL columns (filled or not). The standard ALV functions
*& include "Export -> Spreadsheet" for the download.
*&---------------------------------------------------------------------*
FORM display.

  DATA lo_alv  TYPE REF TO cl_salv_table.
  DATA lx_salv TYPE REF TO cx_salv_error.

  IF gt_out IS INITIAL.
    MESSAGE 'No data to display' TYPE 'S' DISPLAY LIKE 'E'.
    RETURN.
  ENDIF.

  TRY.
      cl_salv_table=>factory( IMPORTING r_salv_table = lo_alv
                              CHANGING  t_table      = gt_out ).
      lo_alv->get_functions( )->set_all( abap_true ).
      lo_alv->get_display_settings( )->set_striped_pattern( abap_true ).
      lo_alv->get_display_settings( )->set_list_header(
        'Export Packing List - fill the empty columns in Excel and upload the file to print' ).
      LOOP AT gt_labels INTO DATA(ls_label).
        PERFORM column_text USING lo_alv ls_label-name ls_label-text ls_label-len.
      ENDLOOP.
      lo_alv->display( ).
    CATCH cx_salv_error INTO lx_salv.
      MESSAGE lx_salv TYPE 'S' DISPLAY LIKE 'E'.
  ENDTRY.

ENDFORM.

*&=====================================================================*
*& STEP 2 - filled Excel -> Adobe Form
*&=====================================================================*

*&---------------------------------------------------------------------*
*& Filled Excel (downloaded from the ALV) -> GT_OUT.
*& Columns are found by their heading (ALV column text or field name),
*& so the column order of the file does not matter.
*&---------------------------------------------------------------------*
FORM read_upload.

  TYPES: BEGIN OF ty_col,
           col  TYPE i,                 " column in the file
           name TYPE lvc_fname,         " field of TY_OUT
         END OF ty_col.

  DATA lr_data   TYPE REF TO data.
  DATA lt_map    TYPE ty_colmap_tab.
  DATA lt_cols   TYPE STANDARD TABLE OF ty_col WITH EMPTY KEY.
  DATA ls_out    TYPE ty_out.
  DATA lv_text   TYPE string.
  DATA lv_key    TYPE string.
  DATA lv_col    TYPE i.
  DATA lv_row    TYPE i.
  DATA lv_packno TYPE abap_bool.

  FIELD-SYMBOLS <lt_excel> TYPE STANDARD TABLE.
  FIELD-SYMBOLS <ls_excel> TYPE any.
  FIELD-SYMBOLS <lv_cell>  TYPE any.
  FIELD-SYMBOLS <lv_field> TYPE any.

  PERFORM read_xlsx CHANGING lr_data.
  ASSIGN lr_data->* TO <lt_excel>.
  IF <lt_excel> IS NOT ASSIGNED OR lines( <lt_excel> ) < 2.
    MESSAGE 'The file contains no data below the heading row' TYPE 'S' DISPLAY LIKE 'E'.
    LEAVE LIST-PROCESSING.
  ENDIF.

  " Heading row -> column map
  PERFORM heading_map CHANGING lt_map.
  READ TABLE <lt_excel> ASSIGNING <ls_excel> INDEX 1.
  DO.
    lv_col = sy-index.
    ASSIGN COMPONENT lv_col OF STRUCTURE <ls_excel> TO <lv_cell>.
    IF sy-subrc <> 0.
      EXIT.
    ENDIF.
    lv_text = <lv_cell>.
    PERFORM normalize USING lv_text CHANGING lv_key.
    READ TABLE lt_map INTO DATA(ls_map) WITH TABLE KEY key = lv_key.
    IF sy-subrc = 0 AND ls_map-name <> '*'.
      APPEND VALUE #( col = lv_col name = ls_map-name ) TO lt_cols.
      IF ls_map-name = 'PACKNO'.
        lv_packno = abap_true.
      ENDIF.
    ENDIF.
  ENDDO.

  IF lv_packno = abap_false.
    MESSAGE 'Column "Packing List No." not found in the heading row - use the file downloaded from the ALV'
      TYPE 'S' DISPLAY LIKE 'E'.
    LEAVE LIST-PROCESSING.
  ENDIF.

  " Data rows
  DELETE <lt_excel> INDEX 1.
  CLEAR gt_out.

  LOOP AT <lt_excel> ASSIGNING <ls_excel>.
    lv_row = sy-tabix + 1.
    CLEAR ls_out.
    LOOP AT lt_cols INTO DATA(ls_col).
      ASSIGN COMPONENT ls_col-col OF STRUCTURE <ls_excel> TO <lv_cell>.
      IF sy-subrc <> 0.
        CONTINUE.
      ENDIF.
      lv_text = <lv_cell>.
      CONDENSE lv_text.
      IF lv_text IS INITIAL.
        CONTINUE.
      ENDIF.
      ASSIGN COMPONENT ls_col-name OF STRUCTURE ls_out TO <lv_field>.
      PERFORM convert_cell USING lv_text lv_row ls_col-name CHANGING <lv_field>.
    ENDLOOP.

    IF ls_out IS INITIAL.
      CONTINUE.                                            " empty row
    ENDIF.
    IF ls_out-packno IS INITIAL.
      APPEND VALUE #( type = 'W' ref = |Row { lv_row }| text = 'No packing list number - row not printed' ) TO gt_log.
      CONTINUE.
    ENDIF.
    APPEND ls_out TO gt_out.
  ENDLOOP.

  IF gt_out IS INITIAL.
    MESSAGE 'No row with a packing list number in the file' TYPE 'S' DISPLAY LIKE 'E'.
    LEAVE LIST-PROCESSING.
  ENDIF.

ENDFORM.

*&---------------------------------------------------------------------*
*& Excel file -> first worksheet as internal table (binary upload)
*&---------------------------------------------------------------------*
FORM read_xlsx CHANGING cr_data TYPE REF TO data.

  DATA lt_bin     TYPE solix_tab.
  DATA lv_size    TYPE i.
  DATA lv_xstring TYPE xstring.
  DATA lv_file    TYPE string.
  DATA lo_excel   TYPE REF TO cl_fdt_xl_spreadsheet.
  DATA lt_sheets  TYPE if_fdt_doc_spreadsheet=>t_worksheet_names.
  DATA lv_sheet   TYPE string.

  lv_file = p_file.

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

  cr_data = lo_excel->if_fdt_doc_spreadsheet~get_itab_from_worksheet( lv_sheet ).

ENDFORM.

*&---------------------------------------------------------------------*
*& Heading text -> field: ALV long text (40), medium (20), short (10)
*& and the technical field name are accepted. A heading that fits more
*& than one field is marked ambiguous ('*').
*&---------------------------------------------------------------------*
FORM heading_map CHANGING ct_map TYPE ty_colmap_tab.

  DATA lt_keys TYPE ty_string_tab.
  DATA lv_text TYPE string.
  DATA lv_key  TYPE string.

  CLEAR ct_map.
  LOOP AT gt_labels INTO DATA(ls_label).
    CLEAR lt_keys.
    lv_text = ls_label-name.
    APPEND lv_text TO lt_keys.
    APPEND ls_label-text TO lt_keys.
    IF strlen( ls_label-text ) > 20.
      APPEND substring( val = ls_label-text len = 20 ) TO lt_keys.
    ENDIF.
    IF strlen( ls_label-text ) > 10.
      APPEND substring( val = ls_label-text len = 10 ) TO lt_keys.
    ENDIF.
    LOOP AT lt_keys INTO lv_text.
      PERFORM normalize USING lv_text CHANGING lv_key.
      READ TABLE ct_map ASSIGNING FIELD-SYMBOL(<ls_map>) WITH TABLE KEY key = lv_key.
      IF sy-subrc <> 0.
        INSERT VALUE #( key = lv_key name = ls_label-name ) INTO TABLE ct_map.
      ELSEIF <ls_map>-name <> ls_label-name.
        <ls_map>-name = '*'.
      ENDIF.
    ENDLOOP.
  ENDLOOP.

ENDFORM.

FORM normalize USING    iv_text TYPE string
               CHANGING cv_key  TYPE string.
  cv_key = replace( val = to_upper( iv_text ) regex = `[^A-Z0-9]` with = `` occ = 0 ).
ENDFORM.

*&---------------------------------------------------------------------*
*& Excel cell -> field of TY_OUT according to its type
*&---------------------------------------------------------------------*
FORM convert_cell USING    iv_text  TYPE string
                           iv_row   TYPE i
                           iv_name  TYPE lvc_fname
                  CHANGING cv_field TYPE any.

  DATA lv_kind TYPE c LENGTH 1.
  DATA lv_date TYPE d.
  DATA lv_dec  TYPE p LENGTH 16 DECIMALS 3.
  DATA lv_ok   TYPE abap_bool.
  DATA lv_ref  TYPE string.

  lv_ref = |Row { iv_row } / { iv_name }|.
  DESCRIBE FIELD cv_field TYPE lv_kind.

  CASE lv_kind.
    WHEN 'D'.
      PERFORM to_date USING iv_text lv_ref CHANGING lv_date lv_ok.
      cv_field = lv_date.
    WHEN 'I' OR 'P' OR 'N' OR 'F'.
      PERFORM to_dec USING iv_text lv_ref CHANGING lv_dec.
      cv_field = lv_dec.
    WHEN OTHERS.
      cv_field = iv_text.
  ENDCASE.

ENDFORM.

*&---------------------------------------------------------------------*
*& One Adobe Form per packing list number with the data of the file
*&---------------------------------------------------------------------*
FORM print_forms.

  DATA lt_packno TYPE ty_string_tab.
  DATA lt_rows   TYPE ty_out_tab.
  DATA ls_hdr    TYPE zsd_s_epack_af_hdr.
  DATA lt_itm    TYPE zsd_tt_epack_af_itm.
  DATA lv_goods  TYPE string.
  DATA lv_ok     TYPE abap_bool.
  DATA lv_count  TYPE i.

  " Packing lists in the order of the file
  LOOP AT gt_out INTO DATA(ls_out).
    IF NOT line_exists( lt_packno[ table_line = ls_out-packno ] ).
      APPEND ls_out-packno TO lt_packno.
    ENDIF.
  ENDLOOP.

  LOOP AT lt_packno INTO DATA(lv_packno).
    CLEAR lt_rows.
    LOOP AT gt_out INTO ls_out WHERE packno = lv_packno.
      APPEND ls_out TO lt_rows.
    ENDLOOP.

    PERFORM form_header USING lt_rows CHANGING ls_hdr.
    lv_goods = ls_hdr-goods_desc.
    PERFORM form_items  USING lt_rows lv_goods CHANGING lt_itm.
    IF lt_itm IS INITIAL.
      APPEND VALUE #( type = 'E' ref = lv_packno text = 'No line item (Sl.No. / Part No.) - not printed' ) TO gt_log.
      CONTINUE.
    ENDIF.

    PERFORM send_form USING lv_packno ls_hdr lt_itm CHANGING lv_ok.
    IF lv_ok = abap_true.
      lv_count = lv_count + 1.
      APPEND VALUE #( type = 'S' ref = lv_packno text = |Form { gc_form } output on { gc_dest }| ) TO gt_log.
    ENDIF.
  ENDLOOP.

  MESSAGE |{ lv_count } of { lines( lt_packno ) } packing list(s) printed| TYPE 'S'.

ENDFORM.

*&---------------------------------------------------------------------*
*& Form header from the rows of one packing list: for each field the
*& first filled value of the rows (header columns may be filled on the
*& first row only)
*&---------------------------------------------------------------------*
FORM form_header USING    it_rows TYPE ty_out_tab
                 CHANGING cs_hdr  TYPE zsd_s_epack_af_hdr.

  DATA ls_h      TYPE ty_out.
  DATA lt_parts  TYPE ty_string_tab.
  DATA lt_addr   TYPE ty_string_tab.
  DATA lt_np     TYPE ty_string_tab.
  DATA lv_text   TYPE string.
  DATA lv_nl     TYPE string.
  DATA lv_art    TYPE i.
  DATA lv_qty    TYPE p LENGTH 16 DECIMALS 3.
  DATA lv_gwt    TYPE p LENGTH 16 DECIMALS 3.
  DATA lv_nwt    TYPE p LENGTH 16 DECIMALS 3.

  FIELD-SYMBOLS <lv_h> TYPE any.
  FIELD-SYMBOLS <lv_r> TYPE any.

  CLEAR cs_hdr.
  lv_nl = cl_abap_char_utilities=>newline.

  LOOP AT it_rows INTO DATA(ls_row).
    DO.
      ASSIGN COMPONENT sy-index OF STRUCTURE ls_h TO <lv_h>.
      IF sy-subrc <> 0.
        EXIT.
      ENDIF.
      ASSIGN COMPONENT sy-index OF STRUCTURE ls_row TO <lv_r>.
      IF <lv_h> IS INITIAL AND <lv_r> IS NOT INITIAL.
        <lv_h> = <lv_r>.
      ENDIF.
    ENDDO.
    lv_art = lv_art + ls_row-articles.
    lv_qty = lv_qty + ls_row-quantity.
    lv_gwt = lv_gwt + ls_row-gross_wt.
    lv_nwt = lv_nwt + ls_row-net_wt.
  ENDLOOP.

  " Document
  cs_hdr-packno = ls_h-packno.
  PERFORM date_text USING ls_h-packdt CHANGING lv_text.
  cs_hdr-packdt = lv_text.

  " Notifier parties -> form format (FS 1.1)
  cs_hdr-np1 = ls_h-np1.
  cs_hdr-np2 = ls_h-np2.
  cs_hdr-np3 = ls_h-np3.
  cs_hdr-np4 = ls_h-np4.
  cs_hdr-np5 = ls_h-np5.
  lt_np = VALUE #( ( ls_h-np1 ) ( ls_h-np2 ) ( ls_h-np3 ) ( ls_h-np4 ) ( ls_h-np5 ) ).
  DELETE lt_np WHERE table_line IS INITIAL.
  cs_hdr-np_count = lines( lt_np ).
  cs_hdr-format   = COND #( WHEN cs_hdr-np_count > 0 THEN 'N' ELSE 'O' ).

  " Exporter: fixed positions 1 name, 2 street, 3 street 2, 4 city,
  " 5 postal code, 6 region, 7 country -> name + address lines
  PERFORM split_address USING ls_h-exporter CHANGING lt_addr.
  cs_hdr-exp_name = lt_addr[ 1 ].
  cs_hdr-sign_for = lt_addr[ 1 ].
  lt_parts = VALUE #( ( lt_addr[ 2 ] ) ( lt_addr[ 3 ] )
                      ( condense( |{ lt_addr[ 4 ] } { COND string( WHEN lt_addr[ 5 ] IS NOT INITIAL
                                                                     THEN |- { lt_addr[ 5 ] }| ) }| ) )
                      ( lt_addr[ 6 ] ) ( lt_addr[ 7 ] ) ).
  PERFORM join USING lt_parts lv_nl CHANGING cs_hdr-exp_address.
  cs_hdr-exp_gstin = ls_h-gstin.
  cs_hdr-exp_iec   = ls_h-iec.
  cs_hdr-exp_cin   = ls_h-cin.

  " Consignee = ship-to, buyer = sold-to
  cs_hdr-same_party = xsdbool( ls_h-sp_kunnr = ls_h-sh_kunnr AND ls_h-sh_kunnr IS NOT INITIAL ).
  " Consignee: fixed positions 1 name, 2 street, 3 street 2, 4 city,
  " 5 district, 6 postal code, 7 country
  PERFORM split_address USING ls_h-ship_to CHANGING lt_addr.
  cs_hdr-cns_name = lt_addr[ 1 ].
  lt_parts = VALUE #( ( lt_addr[ 2 ] ) ( lt_addr[ 3 ] ) ( lt_addr[ 5 ] )
                      ( condense( |{ lt_addr[ 6 ] } { lt_addr[ 4 ] }| ) ) ( lt_addr[ 7 ] ) ).
  PERFORM join USING lt_parts `, ` CHANGING cs_hdr-cns_address.
  cs_hdr-cns_tel    = ls_h-sh_telephone.
  cs_hdr-cns_mobile = ls_h-sh_mobile.
  cs_hdr-cns_email  = ls_h-sh_email.
  " Buyer: same fixed positions as the consignee
  PERFORM split_address USING ls_h-buyer CHANGING lt_addr.
  cs_hdr-buy_name = lt_addr[ 1 ].
  lt_parts = VALUE #( ( lt_addr[ 2 ] ) ( lt_addr[ 3 ] ) ( lt_addr[ 5 ] )
                      ( condense( |{ lt_addr[ 6 ] } { lt_addr[ 4 ] }| ) ) ( lt_addr[ 7 ] ) ).
  PERFORM join USING lt_parts `, ` CHANGING cs_hdr-buy_address.
  cs_hdr-buy_tel    = ls_h-sp_telephone.
  cs_hdr-buy_mobile = ls_h-sp_mobile.
  cs_hdr-buy_email  = ls_h-sp_email.

  " References and countries
  cs_hdr-order_ref = ls_h-order_no.
  IF ls_h-order_date IS NOT INITIAL.
    cs_hdr-order_ref = |{ cs_hdr-order_ref } DATE: { ls_h-order_date }|.
  ENDIF.
  cs_hdr-buyer_ref = ls_h-buyer_ref.
  IF ls_h-buyer_ref_dt IS NOT INITIAL.
    PERFORM date_text USING ls_h-buyer_ref_dt CHANGING lv_text.
    cs_hdr-buyer_ref = |{ cs_hdr-buyer_ref } DATE: { lv_text }|.
  ENDIF.
  cs_hdr-cnty_orgn = ls_h-cnty_orgn.
  cs_hdr-cnty_dest = ls_h-cnty_dest.

  " Shipment
  cs_hdr-pre_carr   = ls_h-pre_carriage.
  cs_hdr-rec_place  = ls_h-place_receipt.
  cs_hdr-vessel     = ls_h-vessel_flight.
  cs_hdr-pol        = ls_h-port_loading.
  cs_hdr-pod        = ls_h-port_disch.
  cs_hdr-pld        = ls_h-place_deliv.
  cs_hdr-marks      = ls_h-marks_nos.
  cs_hdr-containers = ls_h-container_no.
  cs_hdr-no_pkgs    = ls_h-no_packages.
  cs_hdr-goods_desc = ls_h-goods_desc.
  cs_hdr-pay_terms  = |PAYMENT TERM: { COND string( WHEN ls_h-payment_term IS NOT INITIAL
                                                     THEN ls_h-payment_term ELSE ls_h-zterm ) }|.
  cs_hdr-ship_terms = condense( |{ ls_h-inco1 } { ls_h-inco2 }| ).

  " Totals row of the item table (A14) and declaration
  IF ls_h-total_art IS NOT INITIAL.
    lv_art = ls_h-total_art.
  ENDIF.
  cs_hdr-tot_art  = |{ lv_art }|.
  cs_hdr-tot_text = |(Total { lv_art } Packages)|.
  PERFORM qty_text USING lv_qty CHANGING lv_text.
  cs_hdr-tot_qty  = lv_text.
  cs_hdr-tot_gwt  = |{ lv_gwt DECIMALS = 3 }|.
  cs_hdr-tot_nwt  = |{ lv_nwt DECIMALS = 3 }|.
  cs_hdr-declaration = ls_h-declaration.

ENDFORM.

*&---------------------------------------------------------------------*
*& Form item rows: HSN heading rows and one row per line item.
*& An empty Articles / weight cell belongs to the package above (merged
*& cell in the samples) - A15.
*&---------------------------------------------------------------------*
FORM form_items USING    it_rows  TYPE ty_out_tab
                         iv_goods TYPE string
                CHANGING ct_itm   TYPE zsd_tt_epack_af_itm.

  DATA ls_itm      TYPE zsd_s_epack_af_itm.
  DATA ls_next     TYPE zsd_s_epack_af_itm.
  DATA lt_goods    TYPE ty_string_tab.
  DATA lv_prev_hsn TYPE string VALUE '#'.
  DATA lv_hsn      TYPE string.
  DATA lv_digits   TYPE string.
  DATA lv_line_dig TYPE string.
  DATA lv_text     TYPE string.
  DATA lv_art_grp  TYPE i.
  DATA lv_wt_grp   TYPE i.
  DATA lv_next     TYPE i.

  CLEAR ct_itm.
  SPLIT iv_goods AT ` / ` INTO TABLE lt_goods.

  LOOP AT it_rows INTO DATA(ls_row) WHERE sn IS NOT INITIAL OR part_no IS NOT INITIAL.

    " Section heading when the HSN changes
    lv_hsn = condense( val = ls_row-hsn del = ` ` ).
    IF lv_hsn <> lv_prev_hsn.
      CLEAR ls_itm.
      ls_itm-row_type = 'H'.
      lv_digits = replace( val = lv_hsn regex = `[^0-9]` with = `` occ = 0 ).
      IF lv_digits IS NOT INITIAL.
        LOOP AT lt_goods INTO lv_text.
          lv_line_dig = replace( val = lv_text regex = `[^0-9]` with = `` occ = 0 ).
          IF lv_line_dig CS lv_digits.
            ls_itm-row_text = lv_text.
            EXIT.
          ENDIF.
        ENDLOOP.
      ENDIF.
      IF ls_itm-row_text IS INITIAL AND lv_hsn IS NOT INITIAL.
        PERFORM hsn_text USING lv_hsn CHANGING lv_text.
        ls_itm-row_text = |HS CODE : { lv_text }|.
      ENDIF.
      IF ls_itm-row_text IS NOT INITIAL.
        APPEND ls_itm TO ct_itm.
      ENDIF.
      lv_prev_hsn = lv_hsn.
    ENDIF.

    CLEAR ls_itm.
    ls_itm-row_type   = 'I'.
    ls_itm-sl_no      = |{ ls_row-sn }|.
    IF ls_row-inv_sn IS NOT INITIAL.
      ls_itm-inv_sn = |{ ls_row-inv_sn }|.
    ENDIF.
    ls_itm-article_no = ls_row-article_no.
    IF ls_row-std IS NOT INITIAL.
      ls_itm-std = |{ ls_row-std }|.
    ENDIF.
    ls_itm-part_no    = ls_row-part_no.
    ls_itm-prod_desc  = ls_row-prod_desc.
    PERFORM qty_text USING ls_row-quantity CHANGING lv_text.
    ls_itm-quantity   = lv_text.
    PERFORM hsn_text USING lv_hsn CHANGING lv_text.
    ls_itm-hsn        = lv_text.

    " Packages: a filled Articles / weight cell starts a new group
    IF ls_row-articles IS NOT INITIAL OR lv_art_grp = 0.
      lv_art_grp = lv_art_grp + 1.
      ls_itm-art_first = abap_true.
      IF ls_row-articles IS NOT INITIAL.
        ls_itm-articles = |{ ls_row-articles }|.
      ENDIF.
    ENDIF.
    IF ls_row-gross_wt IS NOT INITIAL OR ls_row-net_wt IS NOT INITIAL OR lv_wt_grp = 0.
      lv_wt_grp = lv_wt_grp + 1.
      ls_itm-wt_first = abap_true.
      IF ls_row-gross_wt IS NOT INITIAL OR ls_row-net_wt IS NOT INITIAL.
        ls_itm-gross_wt = |{ ls_row-gross_wt DECIMALS = 3 }|.
        ls_itm-net_wt   = |{ ls_row-net_wt DECIMALS = 3 }|.
      ENDIF.
    ENDIF.
    ls_itm-art_grp = lv_art_grp.
    ls_itm-wt_grp  = lv_wt_grp.
    APPEND ls_itm TO ct_itm.

  ENDLOOP.

  " Last row and number of rows of each group (borders of merged cells)
  LOOP AT ct_itm ASSIGNING FIELD-SYMBOL(<ls_itm>) WHERE row_type = 'I'.
    lv_next = sy-tabix + 1.
    CLEAR ls_next.
    READ TABLE ct_itm INTO ls_next INDEX lv_next.
    IF sy-subrc <> 0 OR ls_next-row_type <> 'I'.
      CLEAR ls_next.
    ENDIF.
    <ls_itm>-art_last = xsdbool( ls_next-art_grp <> <ls_itm>-art_grp ).
    <ls_itm>-wt_last  = xsdbool( ls_next-wt_grp  <> <ls_itm>-wt_grp ).
    <ls_itm>-art_rows = REDUCE i( INIT n = 0 FOR r IN ct_itm
                                  WHERE ( row_type = 'I' AND art_grp = <ls_itm>-art_grp ) NEXT n = n + 1 ).
    <ls_itm>-wt_rows  = REDUCE i( INIT n = 0 FOR r IN ct_itm
                                  WHERE ( row_type = 'I' AND wt_grp = <ls_itm>-wt_grp ) NEXT n = n + 1 ).
  ENDLOOP.

ENDFORM.

*&---------------------------------------------------------------------*
*& Adobe Form call for one packing list: print preview on LP01
*&---------------------------------------------------------------------*
FORM send_form USING    iv_packno TYPE string
                        is_hdr    TYPE zsd_s_epack_af_hdr
                        it_itm    TYPE zsd_tt_epack_af_itm
               CHANGING cv_ok     TYPE abap_bool.

  DATA lv_fm      TYPE funcname.
  DATA ls_outpar  TYPE sfpoutputparams.
  DATA ls_docpar  TYPE sfpdocparams.
  DATA ls_formout TYPE fpformoutput.
  DATA lv_subrc   TYPE sysubrc.

  cv_ok = abap_false.

  TRY.
      CALL FUNCTION 'FP_FUNCTION_MODULE_NAME'
        EXPORTING
          i_name     = gc_form
        IMPORTING
          e_funcname = lv_fm.
    CATCH cx_fp_api.
      APPEND VALUE #( type = 'E' ref = iv_packno
                      text = |Adobe Form { gc_form } does not exist or is not active| ) TO gt_log.
      RETURN.
  ENDTRY.

  ls_outpar-nodialog = abap_true.
  ls_outpar-preview  = abap_true.
  ls_outpar-reqnew   = abap_true.
  ls_outpar-dest     = gc_dest.

  CALL FUNCTION 'FP_JOB_OPEN'
    CHANGING
      ie_outputparams = ls_outpar
    EXCEPTIONS
      cancel          = 1
      usage_error     = 2
      system_error    = 3
      internal_error  = 4
      OTHERS          = 5.
  IF sy-subrc <> 0.
    lv_subrc = sy-subrc.
    APPEND VALUE #( type = 'E' ref = iv_packno text = |Print job could not be opened (FP_JOB_OPEN { lv_subrc })| ) TO gt_log.
    RETURN.
  ENDIF.

  ls_docpar-langu   = gc_langu.
  ls_docpar-country = gc_land_in.

  CALL FUNCTION lv_fm
    EXPORTING
      /1bcdwb/docparams  = ls_docpar
      is_header          = is_hdr
      it_item            = it_itm
    IMPORTING
      /1bcdwb/formoutput = ls_formout
    EXCEPTIONS
      usage_error        = 1
      system_error       = 2
      internal_error     = 3
      OTHERS             = 4.
  lv_subrc = sy-subrc.

  CALL FUNCTION 'FP_JOB_CLOSE'
    EXCEPTIONS
      usage_error    = 1
      system_error   = 2
      internal_error = 3
      OTHERS         = 4.

  IF lv_subrc <> 0.
    APPEND VALUE #( type = 'E' ref = iv_packno text = |Form output failed ({ lv_subrc })| ) TO gt_log.
    RETURN.
  ENDIF.

  cv_ok = abap_true.

ENDFORM.

*&=====================================================================*
*& Common routines
*&=====================================================================*

*&---------------------------------------------------------------------*
*& Column headings = ALV texts = expected Excel headings
*&---------------------------------------------------------------------*
FORM labels.

  gt_labels = VALUE #(
    ( name = 'SALES_ORDERS'   text = 'Sales Order' len = 12 )
    ( name = 'INVOICES'       text = 'SAP Invoice Number' len = 12 )
    ( name = 'EXPORTER'       text = 'Exporter Address' len = 60 )
    ( name = 'GSTIN'          text = 'GST No.' len = 16 )
    ( name = 'CIN'            text = 'CIN No.' len = 22 )
    ( name = 'PACKNO'         text = 'Packing List No.' len = 12 )
    ( name = 'PACKDT'         text = 'Packing List Date' len = 10 )
    ( name = 'IEC'            text = 'Exporter Ref. (IEC No.)' len = 12 )
    ( name = 'ORDER_NO'       text = 'Order No.' len = 20 )
    ( name = 'ORDER_DATE'     text = 'Order Date' len = 10 )
    ( name = 'BUYER_REF'      text = 'Buyer''s Reference No.' len = 25 )
    ( name = 'BUYER_REF_DT'   text = 'Buyer''s Reference Date' len = 10 )
    ( name = 'CNTY_ORGN'      text = 'Country of Origin of Goods' len = 15 )
    ( name = 'CNTY_DEST'      text = 'Country of Final Destination' len = 15 )
    ( name = 'SP_KUNNR'       text = 'Buyer (Sold-to)' len = 10 )
    ( name = 'BUYER'          text = 'Buyer Address' len = 60 )
    ( name = 'SP_EMAIL'       text = 'Buyer E-mail' len = 25 )
    ( name = 'SP_TELEPHONE'   text = 'Buyer Telephone' len = 15 )
    ( name = 'SP_MOBILE'      text = 'Buyer Mobile Number' len = 15 )
    ( name = 'SH_KUNNR'       text = 'Ship-to Party' len = 10 )
    ( name = 'SHIP_TO'        text = 'Ship-to Address' len = 60 )
    ( name = 'SH_EMAIL'       text = 'Ship-to E-mail' len = 25 )
    ( name = 'SH_TELEPHONE'   text = 'Ship-to Telephone' len = 15 )
    ( name = 'SH_MOBILE'      text = 'Ship-to Mobile Number' len = 15 )
    ( name = 'NP1'            text = 'Notifier Party 1' len = 30 )
    ( name = 'NP2'            text = 'Notifier Party 2' len = 30 )
    ( name = 'NP3'            text = 'Notifier Party 3' len = 30 )
    ( name = 'NP4'            text = 'Notifier Party 4' len = 30 )
    ( name = 'NP5'            text = 'Notifier Party 5' len = 30 )
    ( name = 'PRE_CARRIAGE'   text = 'Pre-Carriage' len = 15 )
    ( name = 'PLACE_RECEIPT'  text = 'Place of Receipt by Pre-carrier' len = 20 )
    ( name = 'VESSEL_FLIGHT'  text = 'Vessel / Flight No.' len = 15 )
    ( name = 'PORT_LOADING'   text = 'Port of Loading' len = 20 )
    ( name = 'PORT_DISCH'     text = 'Port of Discharge' len = 20 )
    ( name = 'PLACE_DELIV'    text = 'Place of Delivery' len = 20 )
    ( name = 'MARKS_NOS'      text = 'Marks & Nos.' len = 15 )
    ( name = 'CONTAINER_NO'   text = 'Container No.' len = 25 )
    ( name = 'NO_PACKAGES'    text = 'No. of Packages' len = 12 )
    ( name = 'GOODS_DESC'     text = 'Description of Goods' len = 40 )
    ( name = 'ZTERM'          text = 'Payment Term' len = 6 )
    ( name = 'PAYMENT_TERM'   text = 'Payment Term Desc.' len = 25 )
    ( name = 'INCO1'          text = 'Incoterms' len = 6 )
    ( name = 'INCO2'          text = 'Incoterms Location' len = 15 )
    ( name = 'SN'             text = 'Sl.No.' len = 6 )
    ( name = 'INV_SN'         text = 'Sr.No. INV' len = 6 )
    ( name = 'ARTICLE_NO'     text = 'Article No.' len = 12 )
    ( name = 'STD'            text = 'STD' len = 6 )
    ( name = 'ARTICLES'       text = 'Articles' len = 8 )
    ( name = 'TOTAL_ART'      text = 'Total Articles' len = 8 )
    ( name = 'PART_NO'        text = 'Part No.' len = 18 )
    ( name = 'PROD_DESC'      text = 'Product Description' len = 30 )
    ( name = 'QUANTITY'       text = 'Quantity Pcs' len = 12 )
    ( name = 'GROSS_WT'       text = 'Gross Wt. KGS' len = 12 )
    ( name = 'NET_WT'         text = 'Net Wt. KGS' len = 12 )
    ( name = 'HSN'            text = 'H.S Code' len = 10 )
    ( name = 'DECLARATION'    text = 'Declaration' len = 40 ) ).

ENDFORM.

FORM column_text USING io_alv  TYPE REF TO cl_salv_table
                       iv_name TYPE csequence
                       iv_text TYPE csequence
                       iv_len  TYPE lvc_outlen.

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
      " Fixed width: visible also when empty, reasonable in the Excel export
      lo_column->set_output_length( iv_len ).
    CATCH cx_salv_not_found.
      RETURN.
  ENDTRY.

ENDFORM.

*&---------------------------------------------------------------------*
*& Message log as popup
*&---------------------------------------------------------------------*
FORM show_log.

  DATA lo_log     TYPE REF TO cl_salv_table.
  DATA lx_salv    TYPE REF TO cx_salv_error.
  DATA lv_len_4   TYPE lvc_outlen VALUE 4.
  DATA lv_len_20  TYPE lvc_outlen VALUE 20.
  DATA lv_len_100 TYPE lvc_outlen VALUE 100.

  IF gt_log IS INITIAL.
    RETURN.
  ENDIF.

  TRY.
      cl_salv_table=>factory( IMPORTING r_salv_table = lo_log
                              CHANGING  t_table      = gt_log ).
      lo_log->set_screen_popup( start_column = 5  end_column = 150
                                start_line   = 3  end_line   = 20 ).
      PERFORM column_text USING lo_log 'TYPE' 'Type'      lv_len_4.
      PERFORM column_text USING lo_log 'REF'  'Reference' lv_len_20.
      PERFORM column_text USING lo_log 'TEXT' 'Message'   lv_len_100.
      lo_log->display( ).
    CATCH cx_salv_error INTO lx_salv.
      MESSAGE lx_salv TYPE 'S' DISPLAY LIKE 'E'.
  ENDTRY.

ENDFORM.

*&---------------------------------------------------------------------*
*& Conversions (with a warning when not convertible)
*&---------------------------------------------------------------------*
FORM to_dec USING    iv_text  TYPE string
                     iv_ref   TYPE string
            CHANGING cv_value TYPE p.

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
      APPEND VALUE #( type = 'W' ref = iv_ref
                      text = |"{ iv_text }" is not a number - left empty| ) TO gt_log.
  ENDTRY.

ENDFORM.

FORM to_date USING    iv_text TYPE string
                      iv_ref  TYPE string
             CHANGING cv_date TYPE d
                      cv_ok   TYPE abap_bool.

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

  IF matches( val = lv_text regex = `(19|20)\d{6}` ).
    cv_date = lv_text.                                      " YYYYMMDD
  ELSEIF matches( val = lv_text regex = `\d+(\.\d+)?` ).
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
  APPEND VALUE #( type = 'W' ref = iv_ref
                  text = |"{ iv_text }" is not a date - left empty| ) TO gt_log.

ENDFORM.

FORM date_text USING    iv_date TYPE d
               CHANGING cv_text TYPE string.
  CLEAR cv_text.
  IF iv_date IS NOT INITIAL.
    cv_text = |{ iv_date+6(2) }-{ iv_date+4(2) }-{ iv_date(4) }|.
  ENDIF.
ENDFORM.

FORM qty_text USING    iv_qty  TYPE p
              CHANGING cv_text TYPE string.
  IF frac( iv_qty ) = 0.
    cv_text = |{ CONV i( iv_qty ) }|.
  ELSE.
    cv_text = |{ iv_qty DECIMALS = 3 }|.
  ENDIF.
ENDFORM.

FORM hsn_text USING    iv_hsn  TYPE string
              CHANGING cv_text TYPE string.
  " "39172390" -> "3917 2390" (format of the samples)
  DATA(lv_digits) = replace( val = iv_hsn regex = `\s` with = `` occ = 0 ).
  IF strlen( lv_digits ) = 8 AND lv_digits CO '0123456789'.
    cv_text = |{ lv_digits(4) } { lv_digits+4(4) }|.
  ELSE.
    cv_text = iv_hsn.
  ENDIF.
ENDFORM.

*&---------------------------------------------------------------------*
*& Address parts -> one CHAR100 column with FIXED positions: every part
*& keeps its place, an empty part stays empty between the commas
*& ("NAME, STREET, , CITY, , 380059, INDIA"), so the column can be split
*& back into the right fields. A comma inside a part is replaced by a
*& space (it would shift the positions).
*&---------------------------------------------------------------------*
FORM fixed_address USING    it_parts TYPE ty_string_tab
                            iv_ref   TYPE string
                            iv_what  TYPE csequence
                   CHANGING cv_addr  TYPE c.

  DATA lt_clean TYPE ty_string_tab.
  DATA lv_text  TYPE string.

  LOOP AT it_parts INTO lv_text.
    REPLACE ALL OCCURRENCES OF `,` IN lv_text WITH ` `.
    CONDENSE lv_text.
    APPEND lv_text TO lt_clean.
  ENDLOOP.
  lv_text = concat_lines_of( table = lt_clean sep = `, ` ).
  cv_addr = lv_text.
  IF strlen( lv_text ) > 100.
    APPEND VALUE #( type = 'W' ref = iv_ref
                    text = |{ iv_what } longer than 100 characters - cut off: { lv_text }| ) TO gt_log.
  ENDIF.

ENDFORM.

*&---------------------------------------------------------------------*
*& Fixed-position address column -> its 7 parts (empty parts kept)
*&---------------------------------------------------------------------*
FORM split_address USING    iv_addr  TYPE c
                   CHANGING ct_parts TYPE ty_string_tab.

  DATA lv_text TYPE string.

  CLEAR ct_parts.
  lv_text = iv_addr.
  SPLIT lv_text AT `,` INTO TABLE ct_parts.
  LOOP AT ct_parts ASSIGNING FIELD-SYMBOL(<lv_part>).
    CONDENSE <lv_part>.
  ENDLOOP.
  WHILE lines( ct_parts ) < 7.
    APPEND INITIAL LINE TO ct_parts.
  ENDWHILE.

ENDFORM.

FORM join USING    it_parts TYPE ty_string_tab
                   iv_sep   TYPE csequence
          CHANGING cv_text  TYPE any.

  DATA lv_text TYPE string.

  LOOP AT it_parts INTO DATA(lv_part).
    lv_part = condense( lv_part ).
    IF lv_part IS INITIAL.
      CONTINUE.
    ENDIF.
    IF lv_text IS INITIAL.
      lv_text = lv_part.
    ELSE.
      lv_text = |{ lv_text }{ iv_sep }{ lv_part }|.
    ENDIF.
  ENDLOOP.
  cv_text = lv_text.

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
*&   Selection text S_VBELN  Sales Order (step 1: fetch data)
*&   Selection text P_FILE   Filled Excel file (step 2: print)
*&---------------------------------------------------------------------*
