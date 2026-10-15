*&---------------------------------------------------------------------*
*& Report ZSD_EPACK_EXTRACT                       (T-code ZSD_EPACK)
*&---------------------------------------------------------------------*
*& WRICEF 102-B  Export consolidated packing list - REPORT 1
*&
*& 1. Selection: sales order(s) (from / to, multiple selection) - sheet
*&    "Input" of Logic_sheet1.xlsx.
*& 2. The data of sheet "Data From SAP" is read for every item (VBAP, all
*&    POSNR) of the sales orders. All reads are bulk reads before the
*&    output is built - no SELECT inside a loop.
*& 3. Editable ALV in the layout of sheet "Z table" (ZSD_EPACK_DATA):
*&    columns of sheet "Data From Excel" are editable, SAP columns are
*&    protected. Entries are checked in the DATA_CHANGED event.
*& 4. Button "Download Excel": the ALV entries are taken over
*&    (CHECK_CHANGED_DATA) and an .xlsx file is written with exactly the
*&    columns, headings and order of sheet "Z table" (columns B - AU).
*&    This file is uploaded unchanged by report 2 ZSD_EPACK_UPLOAD.
*& Nothing is saved to ZSD_EPACK_DATA by this report.
*&
*& Field mapping, assumptions: v3/Field_Mapping.md, v3/Analysis_and_Open_Points.md
*&---------------------------------------------------------------------*
REPORT zsd_epack_extract.

INCLUDE zsd_epack_common.

*----------------------------------------------------------------------*
* Constants
*----------------------------------------------------------------------*
CONSTANTS:
  gc_langu       TYPE spras VALUE 'E',           " texts (T005U, T005T, TVZBT, MAKT)
  gc_cin         TYPE t001z-party VALUE 'ZCIN',  " "PARTY = ZCIN"
  gc_parvw_sh    TYPE parvw VALUE 'WE',          " "PARVW = SH" (internal code WE)
  gc_tel         TYPE adr2-r3_user VALUE '1',    " "R3_USER = 1" telephone
  gc_mob         TYPE adr2-r3_user VALUE '3',    " "R3_USER = 3" mobile
  gc_gewei       TYPE gewei VALUE 'KG',          " weights in KGS (column heading)
  " GUI status: copy of program SAPLKKBL status STANDARD_FULLSCREEN with
  " the additional function ZDOWNLOAD (see Activation steps)
  gc_status      TYPE sypfkey VALUE 'ZEPACK_STD',
  gc_fc_download TYPE syucomm VALUE 'ZDOWNLOAD'.

*----------------------------------------------------------------------*
* Types - SAP data buffers
*----------------------------------------------------------------------*
TYPES:
  BEGIN OF ty_vbak,
    vbeln  TYPE vbak-vbeln,
    vkorg  TYPE vbak-vkorg,
    vtweg  TYPE vbak-vtweg,
    spart  TYPE vbak-spart,
    kunnr  TYPE vbak-kunnr,
    vgbel  TYPE vbak-vgbel,
    bukrs  TYPE tvko-bukrs,
    denied TYPE abap_bool,              " no display authorization
  END OF ty_vbak,
  ty_vbak_tab TYPE SORTED TABLE OF ty_vbak WITH UNIQUE KEY vbeln,

  BEGIN OF ty_vbap,
    vbeln  TYPE vbap-vbeln,
    posnr  TYPE vbap-posnr,
    matnr  TYPE vbap-matnr,
    werks  TYPE vbap-werks,
    kwmeng TYPE vbap-kwmeng,
    vrkme  TYPE vbap-vrkme,
  END OF ty_vbap,
  ty_vbap_tab TYPE SORTED TABLE OF ty_vbap WITH UNIQUE KEY vbeln posnr,

  BEGIN OF ty_vbkd,
    vbeln TYPE vbkd-vbeln,
    zterm TYPE vbkd-zterm,
    inco1 TYPE vbkd-inco1,
    inco2 TYPE vbkd-inco2,
    bstkd TYPE vbkd-bstkd,
    bstdk TYPE vbkd-bstdk,
  END OF ty_vbkd,
  ty_vbkd_tab TYPE HASHED TABLE OF ty_vbkd WITH UNIQUE KEY vbeln,

  BEGIN OF ty_t001,
    bukrs TYPE t001-bukrs,
    adrnr TYPE t001-adrnr,
  END OF ty_t001,
  ty_t001_tab TYPE HASHED TABLE OF ty_t001 WITH UNIQUE KEY bukrs,

  BEGIN OF ty_t001w,
    werks  TYPE t001w-werks,
    branch TYPE t001w-j_1bbranch,
  END OF ty_t001w,
  ty_t001w_tab TYPE HASHED TABLE OF ty_t001w WITH UNIQUE KEY werks,

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

  BEGIN OF ty_kna1,
    kunnr TYPE kna1-kunnr,
    adrnr TYPE kna1-adrnr,
  END OF ty_kna1,
  ty_kna1_tab TYPE HASHED TABLE OF ty_kna1 WITH UNIQUE KEY kunnr,

  BEGIN OF ty_knvp,
    kunnr TYPE knvp-kunnr,
    vkorg TYPE knvp-vkorg,
    vtweg TYPE knvp-vtweg,
    spart TYPE knvp-spart,
    parza TYPE knvp-parza,
    kunn2 TYPE knvp-kunn2,
    adrnr TYPE kna1-adrnr,
  END OF ty_knvp,
  ty_knvp_tab TYPE SORTED TABLE OF ty_knvp WITH NON-UNIQUE KEY kunnr vkorg vtweg spart parza,

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
  ty_adrc_std TYPE STANDARD TABLE OF ty_adrc WITH EMPTY KEY,
  ty_adrc_tab TYPE HASHED TABLE OF ty_adrc WITH UNIQUE KEY addrnumber,

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

  BEGIN OF ty_makt,
    matnr TYPE makt-matnr,
    maktx TYPE makt-maktx,
  END OF ty_makt,
  ty_makt_tab TYPE HASHED TABLE OF ty_makt WITH UNIQUE KEY matnr.

*----------------------------------------------------------------------*
* Global data
*----------------------------------------------------------------------*
DATA gv_vbeln TYPE vbak-vbeln.                         " for SELECT-OPTIONS

" Output = ALV = Excel: typed like the Z table
DATA gt_data TYPE STANDARD TABLE OF zsd_epack_data WITH EMPTY KEY.
DATA gt_cols TYPE ty_col_tab.                          " Excel layout
DATA gt_msg  TYPE ty_msg_tab.                          " message log

DATA gt_vbak   TYPE ty_vbak_tab.
DATA gt_vbap   TYPE ty_vbap_tab.
DATA gt_vbkd   TYPE ty_vbkd_tab.
DATA gt_t001   TYPE ty_t001_tab.
DATA gt_t001w  TYPE ty_t001w_tab.
DATA gt_branch TYPE ty_branch_tab.
DATA gt_t001z  TYPE ty_t001z_tab.
DATA gt_kna1   TYPE ty_kna1_tab.
DATA gt_knvp   TYPE ty_knvp_tab.
DATA gt_adrc   TYPE ty_adrc_tab.
DATA gt_adr6   TYPE ty_adr6_tab.
DATA gt_adr2   TYPE ty_adr2_tab.
DATA gt_t005u  TYPE ty_t005u_tab.
DATA gt_t005t  TYPE ty_t005t_tab.
DATA gt_tvzbt  TYPE ty_tvzbt_tab.
DATA gt_marc   TYPE ty_marc_tab.
DATA gt_makt   TYPE ty_makt_tab.

*----------------------------------------------------------------------*
* Selection screen - sheet "Input": sales order from / to, multiple
*----------------------------------------------------------------------*
SELECTION-SCREEN BEGIN OF BLOCK b1 WITH FRAME TITLE TEXT-b01.
SELECT-OPTIONS s_vbeln FOR gv_vbeln OBLIGATORY.
SELECTION-SCREEN END OF BLOCK b1.

*----------------------------------------------------------------------*
INITIALIZATION.
*----------------------------------------------------------------------*
  PERFORM epack_columns CHANGING gt_cols.

*----------------------------------------------------------------------*
START-OF-SELECTION.
*----------------------------------------------------------------------*
  PERFORM read_orders.           " VBAK + TVKO, authorization, VBAP
  PERFORM read_master_data.      " all other tables, bulk
  PERFORM build_rows.            " one row per sales order item
  PERFORM show_messages USING gt_msg 'Messages - data from SAP' abap_true.
  PERFORM display_alv.

*&---------------------------------------------------------------------*
*& Sales orders (VBAK), company code (TVKO), items (VBAP)
*&---------------------------------------------------------------------*
FORM read_orders.

  " VBELN -> VKORG -> TVKO-BUKRS (sheet "Data From SAP", Exporter)
  SELECT k~vbeln, k~vkorg, k~vtweg, k~spart, k~kunnr, k~vgbel, o~bukrs
    FROM vbak AS k
    LEFT OUTER JOIN tvko AS o ON o~vkorg = k~vkorg
    WHERE k~vbeln IN @s_vbeln
    INTO CORRESPONDING FIELDS OF TABLE @gt_vbak.
  IF gt_vbak IS INITIAL.
    MESSAGE 'No sales order found for the selection' TYPE 'S' DISPLAY LIKE 'E'.
    LEAVE LIST-PROCESSING.
  ENDIF.

  " Single values entered that do not exist
  LOOP AT s_vbeln INTO DATA(ls_sel) WHERE sign = 'I' AND option = 'EQ'.
    IF NOT line_exists( gt_vbak[ vbeln = ls_sel-low ] ).
      APPEND VALUE #( type = 'E' column = |Sales order { ls_sel-low ALPHA = OUT }|
                      text = 'Sales order does not exist' ) TO gt_msg.
    ENDIF.
  ENDLOOP.

  " Display authorization for the sales area (standard object V_VBAK_VKO)
  LOOP AT gt_vbak ASSIGNING FIELD-SYMBOL(<ls_vbak>).
    AUTHORITY-CHECK OBJECT 'V_VBAK_VKO'
      ID 'VKORG' FIELD <ls_vbak>-vkorg
      ID 'VTWEG' FIELD <ls_vbak>-vtweg
      ID 'SPART' FIELD <ls_vbak>-spart
      ID 'ACTVT' FIELD '03'.
    IF sy-subrc <> 0.
      <ls_vbak>-denied = abap_true.
      APPEND VALUE #( type = 'E' column = |Sales order { <ls_vbak>-vbeln ALPHA = OUT }|
                      text = |No authorization for sales area { <ls_vbak>-vkorg }/{ <ls_vbak>-vtweg }/{ <ls_vbak>-spart }| ) TO gt_msg.
    ENDIF.
  ENDLOOP.
  DELETE gt_vbak WHERE denied = abap_true.
  IF gt_vbak IS INITIAL.
    PERFORM show_messages USING gt_msg 'Messages - data from SAP' abap_true.
    MESSAGE 'No authorized sales order in the selection' TYPE 'S' DISPLAY LIKE 'E'.
    LEAVE LIST-PROCESSING.
  ENDIF.

  " Line items: all POSNR (sheet "Data From SAP", Line Item Data)
  SELECT vbeln, posnr, matnr, werks, kwmeng, vrkme
    FROM vbap
    FOR ALL ENTRIES IN @gt_vbak
    WHERE vbeln = @gt_vbak-vbeln
    INTO TABLE @gt_vbap.

  LOOP AT gt_vbak INTO DATA(ls_vbak).
    IF NOT line_exists( gt_vbap[ vbeln = ls_vbak-vbeln ] ).
      APPEND VALUE #( type = 'W' column = |Sales order { ls_vbak-vbeln ALPHA = OUT }|
                      text = 'Sales order has no items' ) TO gt_msg.
    ENDIF.
  ENDLOOP.
  IF gt_vbap IS INITIAL.
    PERFORM show_messages USING gt_msg 'Messages - data from SAP' abap_true.
    MESSAGE 'The selected sales orders have no items' TYPE 'S' DISPLAY LIKE 'E'.
    LEAVE LIST-PROCESSING.
  ENDIF.

ENDFORM.

*&---------------------------------------------------------------------*
*& Master data, each table read once (FOR ALL ENTRIES on unique keys)
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

  DATA lt_vbkd_key TYPE SORTED TABLE OF vbak-vbeln WITH UNIQUE KEY table_line.
  DATA lt_bukrs    TYPE SORTED TABLE OF bukrs WITH UNIQUE KEY table_line.
  DATA lt_werks    TYPE SORTED TABLE OF werks_d WITH UNIQUE KEY table_line.
  DATA lt_branch   TYPE SORTED TABLE OF ty_branch_key WITH UNIQUE KEY bukrs branch.
  DATA lt_kunnr    TYPE SORTED TABLE OF kunnr WITH UNIQUE KEY table_line.
  DATA lt_area     TYPE SORTED TABLE OF ty_area WITH UNIQUE KEY kunnr vkorg vtweg spart.
  DATA lt_adrnr    TYPE SORTED TABLE OF adrnr WITH UNIQUE KEY table_line.
  DATA lt_land     TYPE SORTED TABLE OF land1 WITH UNIQUE KEY table_line.
  DATA lt_region   TYPE SORTED TABLE OF ty_region WITH UNIQUE KEY land1 bland.
  DATA lt_zterm    TYPE SORTED TABLE OF vbkd-zterm WITH UNIQUE KEY table_line.
  DATA lt_matwerks TYPE SORTED TABLE OF ty_matwerks WITH UNIQUE KEY matnr werks.
  DATA lt_matnr    TYPE SORTED TABLE OF matnr WITH UNIQUE KEY table_line.
  DATA lt_adrc     TYPE ty_adrc_std.

  " ---- Keys -----------------------------------------------------------
  LOOP AT gt_vbak INTO DATA(ls_vbak).
    INSERT ls_vbak-vbeln INTO TABLE lt_vbkd_key.          " payment / Incoterms
    IF ls_vbak-vgbel IS NOT INITIAL.
      INSERT ls_vbak-vgbel INTO TABLE lt_vbkd_key.        " order no. & date
    ENDIF.
    IF ls_vbak-bukrs IS NOT INITIAL.
      INSERT ls_vbak-bukrs INTO TABLE lt_bukrs.
    ENDIF.
    INSERT ls_vbak-kunnr INTO TABLE lt_kunnr.
    INSERT VALUE #( kunnr = ls_vbak-kunnr vkorg = ls_vbak-vkorg
                    vtweg = ls_vbak-vtweg spart = ls_vbak-spart ) INTO TABLE lt_area.
  ENDLOOP.
  LOOP AT gt_vbap INTO DATA(ls_vbap) WHERE matnr IS NOT INITIAL.
    INSERT ls_vbap-matnr INTO TABLE lt_matnr.
    INSERT VALUE #( matnr = ls_vbap-matnr werks = ls_vbap-werks ) INTO TABLE lt_matwerks.
    INSERT ls_vbap-werks INTO TABLE lt_werks.
  ENDLOOP.

  " ---- VBKD header data (POSNR 000000) of sales orders and quotations --
  SELECT vbeln, zterm, inco1, inco2, bstkd, bstdk
    FROM vbkd
    FOR ALL ENTRIES IN @lt_vbkd_key
    WHERE vbeln = @lt_vbkd_key-table_line
      AND posnr = '000000'
    INTO TABLE @gt_vbkd.
  LOOP AT gt_vbkd INTO DATA(ls_vbkd) WHERE zterm IS NOT INITIAL.
    INSERT ls_vbkd-zterm INTO TABLE lt_zterm.
  ENDLOOP.

  " ---- Company code: address, CIN --------------------------------------
  IF lt_bukrs IS NOT INITIAL.
    SELECT bukrs, adrnr
      FROM t001
      FOR ALL ENTRIES IN @lt_bukrs
      WHERE bukrs = @lt_bukrs-table_line
      INTO TABLE @gt_t001.
    LOOP AT gt_t001 INTO DATA(ls_t001) WHERE adrnr IS NOT INITIAL.
      INSERT ls_t001-adrnr INTO TABLE lt_adrnr.
    ENDLOOP.

    SELECT bukrs, paval
      FROM t001z
      FOR ALL ENTRIES IN @lt_bukrs
      WHERE bukrs = @lt_bukrs-table_line
        AND party = @gc_cin
      INTO TABLE @gt_t001z.
  ENDIF.

  " ---- GST: business place (BUPLA) of the item plant -> J_1BBRANCH (A10)
  IF lt_werks IS NOT INITIAL.
    SELECT werks, j_1bbranch AS branch
      FROM t001w
      FOR ALL ENTRIES IN @lt_werks
      WHERE werks = @lt_werks-table_line
      INTO TABLE @gt_t001w.
  ENDIF.
  LOOP AT gt_vbap INTO ls_vbap.
    READ TABLE gt_vbak INTO ls_vbak WITH TABLE KEY vbeln = ls_vbap-vbeln.
    READ TABLE gt_t001w INTO DATA(ls_t001w) WITH TABLE KEY werks = ls_vbap-werks.
    IF sy-subrc = 0 AND ls_t001w-branch IS NOT INITIAL AND ls_vbak-bukrs IS NOT INITIAL.
      INSERT VALUE #( bukrs = ls_vbak-bukrs branch = ls_t001w-branch ) INTO TABLE lt_branch.
    ENDIF.
  ENDLOOP.
  IF lt_branch IS NOT INITIAL.
    SELECT bukrs, branch, gstin
      FROM j_1bbranch
      FOR ALL ENTRIES IN @lt_branch
      WHERE bukrs  = @lt_branch-bukrs
        AND branch = @lt_branch-branch
      INTO TABLE @gt_branch.
  ENDIF.

  " ---- Sold-to: VBAK-KUNNR -> KNA1-ADRNR ---------------------------------
  SELECT kunnr, adrnr
    FROM kna1
    FOR ALL ENTRIES IN @lt_kunnr
    WHERE kunnr = @lt_kunnr-table_line
    INTO TABLE @gt_kna1.
  LOOP AT gt_kna1 INTO DATA(ls_kna1) WHERE adrnr IS NOT INITIAL.
    INSERT ls_kna1-adrnr INTO TABLE lt_adrnr.
  ENDLOOP.

  " ---- Ship-to: KNVP (KUNNR, sales area, PARVW = SH) -> KUNN2 -> KNA1 ---
  SELECT p~kunnr, p~vkorg, p~vtweg, p~spart, p~parza, p~kunn2, c~adrnr
    FROM knvp AS p
    INNER JOIN kna1 AS c ON c~kunnr = p~kunn2
    FOR ALL ENTRIES IN @lt_area
    WHERE p~kunnr = @lt_area-kunnr
      AND p~vkorg = @lt_area-vkorg
      AND p~vtweg = @lt_area-vtweg
      AND p~spart = @lt_area-spart
      AND p~parvw = @gc_parvw_sh
    INTO TABLE @gt_knvp.
  LOOP AT gt_knvp INTO DATA(ls_knvp) WHERE adrnr IS NOT INITIAL.
    INSERT ls_knvp-adrnr INTO TABLE lt_adrnr.
  ENDLOOP.

  " ---- Addresses (ADRC), e-mail (ADR6), telephone / mobile (ADR2) -------
  IF lt_adrnr IS NOT INITIAL.
    SELECT addrnumber, date_from, name1, street, str_suppl1, city1, city2,
           post_code1, region, country, building
      FROM adrc
      FOR ALL ENTRIES IN @lt_adrnr
      WHERE addrnumber = @lt_adrnr-table_line
        AND date_from <= @sy-datum
        AND nation     = @space
      INTO TABLE @lt_adrc.
    " Latest valid version per address (A12)
    SORT lt_adrc BY addrnumber date_from DESCENDING.
    DELETE ADJACENT DUPLICATES FROM lt_adrc COMPARING addrnumber.
    gt_adrc = lt_adrc.

    SELECT addrnumber, flgdefault, consnumber, smtp_addr
      FROM adr6
      FOR ALL ENTRIES IN @lt_adrnr
      WHERE addrnumber = @lt_adrnr-table_line
      INTO TABLE @gt_adr6.
    SORT gt_adr6 BY addrnumber flgdefault DESCENDING consnumber.

    SELECT addrnumber, r3_user, flgdefault, consnumber, telnr_long
      FROM adr2
      FOR ALL ENTRIES IN @lt_adrnr
      WHERE addrnumber = @lt_adrnr-table_line
        AND r3_user IN ( @gc_tel, @gc_mob )
      INTO TABLE @gt_adr2.
    SORT gt_adr2 BY addrnumber r3_user flgdefault DESCENDING consnumber.
  ENDIF.

  " ---- Region text (T005U) and country text (T005T) (A11) ---------------
  LOOP AT gt_adrc INTO DATA(ls_adrc) WHERE country IS NOT INITIAL.
    INSERT ls_adrc-country INTO TABLE lt_land.
    IF ls_adrc-region IS NOT INITIAL.
      INSERT VALUE #( land1 = ls_adrc-country bland = ls_adrc-region ) INTO TABLE lt_region.
    ENDIF.
  ENDLOOP.
  IF lt_region IS NOT INITIAL.
    SELECT land1, bland, bezei
      FROM t005u
      FOR ALL ENTRIES IN @lt_region
      WHERE spras = @gc_langu
        AND land1 = @lt_region-land1
        AND bland = @lt_region-bland
      INTO TABLE @gt_t005u.
  ENDIF.
  IF lt_land IS NOT INITIAL.
    SELECT land1, landx
      FROM t005t
      FOR ALL ENTRIES IN @lt_land
      WHERE spras = @gc_langu
        AND land1 = @lt_land-table_line
      INTO TABLE @gt_t005t.
  ENDIF.

  " ---- Payment term text ------------------------------------------------
  IF lt_zterm IS NOT INITIAL.
    SELECT zterm, vtext
      FROM tvzbt
      FOR ALL ENTRIES IN @lt_zterm
      WHERE spras = @gc_langu
        AND zterm = @lt_zterm-table_line
      INTO TABLE @gt_tvzbt.
  ENDIF.

  " ---- Material: HSN (MARC-STEUC), description (MAKT-MAKTX, A11) -------
  IF lt_matwerks IS NOT INITIAL.
    SELECT matnr, werks, steuc
      FROM marc
      FOR ALL ENTRIES IN @lt_matwerks
      WHERE matnr = @lt_matwerks-matnr
        AND werks = @lt_matwerks-werks
      INTO TABLE @gt_marc.

    SELECT matnr, maktx
      FROM makt
      FOR ALL ENTRIES IN @lt_matnr
      WHERE matnr = @lt_matnr-table_line
        AND spras = @gc_langu
      INTO TABLE @gt_makt.
  ENDIF.

ENDFORM.

*&---------------------------------------------------------------------*
*& One row per sales order item. Header data once per sales order.
*& Columns of sheet "Data From Excel" stay empty (entered by the user).
*&---------------------------------------------------------------------*
FORM build_rows.

  DATA ls_head TYPE zsd_epack_data.
  DATA ls_row  TYPE zsd_epack_data.
  DATA lv_prev TYPE vbak-vbeln.

  LOOP AT gt_vbap INTO DATA(ls_vbap).                   " sorted by VBELN POSNR
    READ TABLE gt_vbak INTO DATA(ls_vbak) WITH TABLE KEY vbeln = ls_vbap-vbeln.
    IF ls_vbap-vbeln <> lv_prev.
      PERFORM order_header USING ls_vbak CHANGING ls_head.
      lv_prev = ls_vbap-vbeln.
    ENDIF.
    ls_row = ls_head.
    PERFORM order_item USING ls_vbak ls_vbap CHANGING ls_row.
    APPEND ls_row TO gt_data.
  ENDLOOP.

ENDFORM.

*&---------------------------------------------------------------------*
*& Header data of one sales order (sheet "Data From SAP", Header Data)
*&---------------------------------------------------------------------*
FORM order_header USING    is_vbak TYPE ty_vbak
                  CHANGING cs_row  TYPE zsd_epack_data.

  DATA lt_parts TYPE ty_string_tab.
  DATA lv_ref   TYPE string.
  DATA lv_reg   TYPE string.
  DATA lv_cty   TYPE string.
  DATA lv_long  TYPE abap_bool.
  DATA lv_kunnr TYPE kunnr.
  DATA lv_adrnr TYPE adrnr.
  DATA lv_text  TYPE string.

  CLEAR cs_row.
  lv_ref = |Sales order { is_vbak-vbeln ALPHA = OUT }|.
  cs_row-vbak_vbeln = |{ is_vbak-vbeln ALPHA = OUT }|.

  " ---- Exporter: TVKO-BUKRS -> T001-ADRNR -> ADRC -------------------------
  READ TABLE gt_t001 INTO DATA(ls_t001) WITH TABLE KEY bukrs = is_vbak-bukrs.
  IF sy-subrc <> 0.
    CLEAR ls_t001.
  ENDIF.
  READ TABLE gt_adrc INTO DATA(ls_adrc) WITH TABLE KEY addrnumber = ls_t001-adrnr.
  IF sy-subrc = 0.
    READ TABLE gt_t005u INTO DATA(ls_t005u)
         WITH TABLE KEY land1 = ls_adrc-country bland = ls_adrc-region.
    IF sy-subrc = 0.
      lv_reg = ls_t005u-bezei.
    ENDIF.
    READ TABLE gt_t005t INTO DATA(ls_t005t) WITH TABLE KEY land1 = ls_adrc-country.
    IF sy-subrc = 0.
      lv_cty = ls_t005t-landx.
    ENDIF.
    " NAME1, STREET, STR_SUPPL1, CITY1, POST_CODE1, REGION text, COUNTRY text
    lt_parts = VALUE #( ( CONV string( ls_adrc-name1 ) )
                        ( CONV string( ls_adrc-street ) )
                        ( CONV string( ls_adrc-str_suppl1 ) )
                        ( CONV string( ls_adrc-city1 ) )
                        ( CONV string( ls_adrc-post_code1 ) )
                        ( lv_reg )
                        ( lv_cty ) ).
    PERFORM fixed_address USING lt_parts CHANGING cs_row-zsd_exp lv_long.
    IF lv_long = abap_true.
      PERFORM warn USING lv_ref 'ZSD_EXP' 'Exporter address longer than 100 characters - cut off'.
    ENDIF.
    " Exporter reference number - IEC No.: ADRC-BUILDING
    PERFORM put_text USING ls_adrc-building lv_ref 'ZSD_EXPREF' CHANGING cs_row-zsd_expref.
  ELSE.
    lv_text = |Company code { is_vbak-bukrs }: no address (TVKO / T001 / ADRC)|.
    PERFORM warn USING lv_ref 'ZSD_EXP' lv_text.
  ENDIF.

  " ---- CIN: T001Z PARTY = ZCIN --------------------------------------------
  READ TABLE gt_t001z INTO DATA(ls_t001z) WITH TABLE KEY bukrs = is_vbak-bukrs.
  IF sy-subrc = 0.
    PERFORM put_text USING ls_t001z-paval lv_ref 'ZSD_CIN' CHANGING cs_row-zsd_cin.
  ELSE.
    lv_text = |No CIN in T001Z (party ZCIN) for company code { is_vbak-bukrs }|.
    PERFORM warn USING lv_ref 'ZSD_CIN' lv_text.
  ENDIF.

  " ---- Order no. & date: VBAK-VGBEL -> VBKD BSTKD, BSTDK -------------------
  IF is_vbak-vgbel IS NOT INITIAL.
    READ TABLE gt_vbkd INTO DATA(ls_vbkd) WITH TABLE KEY vbeln = is_vbak-vgbel.
    IF sy-subrc = 0.
      PERFORM put_text USING ls_vbkd-bstkd lv_ref 'ZSD_ORDNO' CHANGING cs_row-zsd_ordno.
      cs_row-zsd_ord_date = ls_vbkd-bstdk.
    ENDIF.
  ELSE.
    PERFORM warn USING lv_ref 'ZSD_ORDNO' 'Sales order has no preceding document (VBAK-VGBEL) - order no. & date empty'.
  ENDIF.

  " ---- Payment term, terms of shipment: VBKD of the sales order ------------
  READ TABLE gt_vbkd INTO ls_vbkd WITH TABLE KEY vbeln = is_vbak-vbeln.
  IF sy-subrc = 0.
    READ TABLE gt_tvzbt INTO DATA(ls_tvzbt) WITH TABLE KEY zterm = ls_vbkd-zterm.
    IF sy-subrc = 0.
      PERFORM put_text USING ls_tvzbt-vtext lv_ref 'ZSD_PTERMS' CHANGING cs_row-zsd_pterms.
    ENDIF.
    lv_text = |{ ls_vbkd-inco1 } { ls_vbkd-inco2 }|.
    PERFORM put_text USING lv_text lv_ref 'ZSD_INCOT' CHANGING cs_row-zsd_incot.
  ENDIF.

  " ---- Sold-to: VBAK-KUNNR -> KNA1-ADRNR ----------------------------------
  READ TABLE gt_kna1 INTO DATA(ls_kna1) WITH TABLE KEY kunnr = is_vbak-kunnr.
  IF sy-subrc = 0.
    lv_adrnr = ls_kna1-adrnr.
  ELSE.
    CLEAR lv_adrnr.
  ENDIF.
  PERFORM party USING lv_adrnr lv_ref 'ZSD_SOLD_PART' CHANGING cs_row-zsd_sold_part lv_cty.

  " ---- Ship-to: KNVP PARVW = SH, lowest PARZA (A13) -------------------------
  CLEAR: lv_kunnr, lv_adrnr.
  LOOP AT gt_knvp INTO DATA(ls_knvp) WHERE kunnr = is_vbak-kunnr
                                       AND vkorg = is_vbak-vkorg
                                       AND vtweg = is_vbak-vtweg
                                       AND spart = is_vbak-spart.
    lv_kunnr = ls_knvp-kunn2.
    lv_adrnr = ls_knvp-adrnr.
    EXIT.
  ENDLOOP.
  IF lv_kunnr IS INITIAL.
    lv_text = |Customer { is_vbak-kunnr ALPHA = OUT }: no ship-to party in KNVP for { is_vbak-vkorg }/{ is_vbak-vtweg }/{ is_vbak-spart }|.
    PERFORM warn USING lv_ref 'ZSD_SHIP_PART' lv_text.
  ENDIF.
  PERFORM party USING lv_adrnr lv_ref 'ZSD_SHIP_PART' CHANGING cs_row-zsd_ship_part lv_cty.
  " Country of final destination = ship-to party country
  cs_row-zsd_cnty_find = lv_cty.

ENDFORM.

*&---------------------------------------------------------------------*
*& Item data (sheet "Data From SAP", Line Item Data)
*&---------------------------------------------------------------------*
FORM order_item USING    is_vbak TYPE ty_vbak
                         is_vbap TYPE ty_vbap
                CHANGING cs_row  TYPE zsd_epack_data.

  DATA lv_partno TYPE c LENGTH 40.
  DATA lv_ref    TYPE string.
  DATA lv_text   TYPE string.

  lv_ref = |Sales order { is_vbap-vbeln ALPHA = OUT } item { is_vbap-posnr ALPHA = OUT }|.

  " Part number = VBAP-MATNR (external format)
  IF is_vbap-matnr IS NOT INITIAL.
    CALL FUNCTION 'CONVERSION_EXIT_MATN1_OUTPUT'
      EXPORTING
        input  = is_vbap-matnr
      IMPORTING
        output = lv_partno.
    cs_row-zsd_partno = condense( lv_partno ).
  ENDIF.

  " Part description = MAKT-MAKTX
  READ TABLE gt_makt INTO DATA(ls_makt) WITH TABLE KEY matnr = is_vbap-matnr.
  IF sy-subrc = 0.
    cs_row-zsd_matdesc = ls_makt-maktx.
  ENDIF.

  " Quantity = VBAP-KWMENG (unit kept for the QUAN reference field)
  cs_row-zsd_qty   = is_vbap-kwmeng.
  cs_row-zsd_meins = is_vbap-vrkme.
  cs_row-zsd_gewei = gc_gewei.

  " HS code = MARC-STEUC (material, item plant)
  READ TABLE gt_marc INTO DATA(ls_marc) WITH TABLE KEY matnr = is_vbap-matnr werks = is_vbap-werks.
  IF sy-subrc = 0.
    PERFORM put_text USING ls_marc-steuc lv_ref 'ZSD_HSN' CHANGING cs_row-zsd_hsn.
  ENDIF.

  " GST No.: plant business place -> J_1BBRANCH-GSTIN (A10)
  READ TABLE gt_t001w INTO DATA(ls_t001w) WITH TABLE KEY werks = is_vbap-werks.
  IF sy-subrc = 0.
    READ TABLE gt_branch INTO DATA(ls_branch)
         WITH TABLE KEY bukrs = is_vbak-bukrs branch = ls_t001w-branch.
  ENDIF.
  IF sy-subrc = 0.
    PERFORM put_text USING ls_branch-gstin lv_ref 'ZSD_GST' CHANGING cs_row-zsd_gst.
  ELSE.
    lv_text = |No GST No. (T001W business place / J_1BBRANCH) for plant { is_vbap-werks }|.
    PERFORM warn USING lv_ref 'ZSD_GST' lv_text.
  ENDIF.

ENDFORM.

*&---------------------------------------------------------------------*
*& Sold-to / ship-to column (CHAR 150), fixed positions (A14):
*& 1 NAME1, 2 STREET, 3 STR_SUPPL1, 4 CITY1, 5 CITY2, 6 POST_CODE1,
*& 7 country text, 8 e-mail (ADR6), 9 telephone (ADR2 R3_USER 1),
*& 10 mobile (ADR2 R3_USER 3)
*&---------------------------------------------------------------------*
FORM party USING    iv_adrnr TYPE adrnr
                    iv_ref   TYPE string
                    iv_field TYPE csequence
           CHANGING cv_party TYPE c
                    cv_cty   TYPE string.

  DATA lt_parts TYPE ty_string_tab.
  DATA lv_email TYPE string.
  DATA lv_tel   TYPE string.
  DATA lv_mob   TYPE string.
  DATA lv_long  TYPE abap_bool.
  DATA lv_text  TYPE string.

  CLEAR: cv_party, cv_cty.
  IF iv_adrnr IS INITIAL.
    RETURN.
  ENDIF.
  READ TABLE gt_adrc INTO DATA(ls_adrc) WITH TABLE KEY addrnumber = iv_adrnr.
  IF sy-subrc <> 0.
    lv_text = |No valid address { iv_adrnr } in ADRC|.
    PERFORM warn USING iv_ref iv_field lv_text.
    RETURN.
  ENDIF.

  READ TABLE gt_t005t INTO DATA(ls_t005t) WITH TABLE KEY land1 = ls_adrc-country.
  IF sy-subrc = 0.
    cv_cty = ls_t005t-landx.
  ENDIF.

  " Default entry first, then the lowest sequence number (A12)
  READ TABLE gt_adr6 INTO DATA(ls_adr6) WITH KEY addrnumber = iv_adrnr BINARY SEARCH.
  IF sy-subrc = 0.
    lv_email = ls_adr6-smtp_addr.
  ENDIF.
  READ TABLE gt_adr2 INTO DATA(ls_adr2) WITH KEY addrnumber = iv_adrnr r3_user = gc_tel BINARY SEARCH.
  IF sy-subrc = 0.
    lv_tel = ls_adr2-telnr_long.
  ENDIF.
  READ TABLE gt_adr2 INTO ls_adr2 WITH KEY addrnumber = iv_adrnr r3_user = gc_mob BINARY SEARCH.
  IF sy-subrc = 0.
    lv_mob = ls_adr2-telnr_long.
  ENDIF.

  lt_parts = VALUE #( ( CONV string( ls_adrc-name1 ) )
                      ( CONV string( ls_adrc-street ) )
                      ( CONV string( ls_adrc-str_suppl1 ) )
                      ( CONV string( ls_adrc-city1 ) )
                      ( CONV string( ls_adrc-city2 ) )
                      ( CONV string( ls_adrc-post_code1 ) )
                      ( cv_cty )
                      ( lv_email )
                      ( lv_tel )
                      ( lv_mob ) ).
  PERFORM fixed_address USING lt_parts CHANGING cv_party lv_long.
  IF lv_long = abap_true.
    PERFORM warn USING iv_ref iv_field 'Address longer than 150 characters - cut off'.
  ENDIF.

ENDFORM.

*&---------------------------------------------------------------------*
*& Text into a CHAR column, warning when it is cut off
*&---------------------------------------------------------------------*
FORM put_text USING    iv_text  TYPE csequence
                       iv_ref   TYPE string
                       iv_field TYPE csequence
              CHANGING cv_field TYPE c.

  DATA lv_text TYPE string.
  DATA lv_len  TYPE i.

  lv_text = condense( CONV string( iv_text ) ).
  cv_field = lv_text.
  DESCRIBE FIELD cv_field LENGTH lv_len IN CHARACTER MODE.
  IF strlen( lv_text ) > lv_len.
    lv_text = |"{ lv_text }" longer than { lv_len } characters - cut off|.
    PERFORM warn USING iv_ref iv_field lv_text.
  ENDIF.

ENDFORM.

FORM warn USING iv_ref   TYPE string
                iv_field TYPE csequence
                iv_text  TYPE csequence.

  DATA lv_header TYPE string.

  READ TABLE gt_cols INTO DATA(ls_col) WITH KEY field = iv_field.
  IF sy-subrc = 0.
    lv_header = ls_col-header.
  ELSE.
    lv_header = iv_field.
  ENDIF.
  APPEND VALUE #( type = 'W' column = |{ iv_ref }: { lv_header }| text = iv_text ) TO gt_msg.

ENDFORM.

*&=====================================================================*
*& Editable ALV
*&=====================================================================*
FORM display_alv.

  DATA lt_fcat   TYPE lvc_t_fcat.
  DATA ls_layout TYPE lvc_s_layo.
  DATA ls_grid   TYPE lvc_s_glay.
  DATA lt_events TYPE slis_t_event.
  DATA lv_width  TYPE i.

  IF gt_data IS INITIAL.
    MESSAGE 'No data to display' TYPE 'S' DISPLAY LIKE 'E'.
    RETURN.
  ENDIF.

  CALL FUNCTION 'LVC_FIELDCATALOG_MERGE'
    EXPORTING
      i_structure_name       = 'ZSD_EPACK_DATA'
      i_bypassing_buffer     = abap_true
    CHANGING
      ct_fieldcat            = lt_fcat
    EXCEPTIONS
      inconsistent_interface = 1
      program_error          = 2
      OTHERS                 = 3.
  IF sy-subrc <> 0.
    MESSAGE 'Field catalog of ZSD_EPACK_DATA could not be built' TYPE 'S' DISPLAY LIKE 'E'.
    RETURN.
  ENDIF.

  " Z-table order and headings; manual columns editable, SAP columns not
  LOOP AT lt_fcat ASSIGNING FIELD-SYMBOL(<ls_fcat>).
    READ TABLE gt_cols INTO DATA(ls_col) WITH KEY field = <ls_fcat>-fieldname.
    IF sy-subrc <> 0 OR ls_col-source = gc_src_copy.
      <ls_fcat>-tech = abap_true.          " MANDT, units, admin fields
      CONTINUE.
    ENDIF.
    <ls_fcat>-col_pos    = ls_col-pos.
    <ls_fcat>-key        = space.
    <ls_fcat>-coltext    = ls_col-header.
    <ls_fcat>-reptext    = ls_col-header.
    <ls_fcat>-scrtext_l  = ls_col-header.
    <ls_fcat>-scrtext_m  = ls_col-header.
    <ls_fcat>-scrtext_s  = ls_col-header.
    <ls_fcat>-colddictxt = 'L'.
    <ls_fcat>-edit       = xsdbool( ls_col-source = gc_src_manual ).
    IF <ls_fcat>-inttype = 'N' OR <ls_fcat>-inttype = 'P'.
      <ls_fcat>-no_zero = abap_true.       " empty cell instead of 0
    ENDIF.
    PERFORM column_width USING <ls_fcat>-fieldname ls_col-header CHANGING lv_width.
    <ls_fcat>-outputlen = lv_width.
  ENDLOOP.
  SORT lt_fcat BY col_pos.

  ls_layout-zebra      = abap_true.
  ls_layout-no_rowins  = abap_true.        " rows = sales order items (A16)
  ls_layout-grid_title = 'Export packing list - fill the editable columns, then Download Excel'.
  ls_grid-edt_cll_cb   = abap_true.        " cell entries to GT_DATA at once

  lt_events = VALUE #( ( name = slis_ev_data_changed form = 'DATA_CHANGED' ) ).

  CALL FUNCTION 'REUSE_ALV_GRID_DISPLAY_LVC'
    EXPORTING
      i_callback_program       = sy-repid
      i_callback_pf_status_set = 'SET_STATUS'
      i_callback_user_command  = 'USER_COMMAND'
      i_grid_settings          = ls_grid
      is_layout_lvc            = ls_layout
      it_fieldcat_lvc          = lt_fcat
      it_events                = lt_events
      i_save                   = 'A'
    TABLES
      t_outtab                 = gt_data
    EXCEPTIONS
      program_error            = 1
      OTHERS                   = 2.
  IF sy-subrc <> 0.
    MESSAGE 'ALV could not be displayed' TYPE 'S' DISPLAY LIKE 'E'.
  ENDIF.

ENDFORM.

*&---------------------------------------------------------------------*
*& Column width = the longer of heading and longest value (max. 60)
*&---------------------------------------------------------------------*
FORM column_width USING    iv_field TYPE csequence
                           iv_text  TYPE string
                  CHANGING cv_width TYPE i.

  CONSTANTS lc_max TYPE i VALUE 60.
  DATA lv_text TYPE string.

  FIELD-SYMBOLS <lv_value> TYPE any.

  cv_width = strlen( iv_text ).
  LOOP AT gt_data ASSIGNING FIELD-SYMBOL(<ls_data>).
    ASSIGN COMPONENT iv_field OF STRUCTURE <ls_data> TO <lv_value>.
    IF sy-subrc <> 0.
      EXIT.
    ENDIF.
    PERFORM field_to_text USING <lv_value> CHANGING lv_text.
    cv_width = nmax( val1 = cv_width val2 = nmin( val1 = strlen( lv_text ) val2 = lc_max ) ).
  ENDLOOP.
  cv_width = nmax( val1 = cv_width val2 = 10 ) + 1.     " room to type a date

ENDFORM.

FORM set_status USING it_extab TYPE slis_t_extab.
  SET PF-STATUS gc_status EXCLUDING it_extab.
ENDFORM.

*&---------------------------------------------------------------------*
*& DATA_CHANGED: entries are checked when the user leaves a cell.
*& Type checks (date, number, length) are done by the ALV itself.
*& Gross weight must not be less than net weight (FS, Additional
*& requirement) - checked as soon as both weights are entered.
*&---------------------------------------------------------------------*
FORM data_changed USING ir_data_changed TYPE REF TO cl_alv_changed_data_protocol.

  DATA lv_gwt TYPE zsd_epack_data-zsd_gwt.
  DATA lv_nwt TYPE zsd_epack_data-zsd_nwt.
  DATA lt_rows TYPE SORTED TABLE OF i WITH UNIQUE KEY table_line.

  LOOP AT ir_data_changed->mt_mod_cells INTO DATA(ls_mod)
       WHERE fieldname = 'ZSD_GWT' OR fieldname = 'ZSD_NWT'.
    INSERT ls_mod-row_id INTO TABLE lt_rows.
  ENDLOOP.

  LOOP AT lt_rows INTO DATA(lv_row).
    READ TABLE gt_data INTO DATA(ls_data) INDEX lv_row.
    IF sy-subrc <> 0.
      CONTINUE.
    ENDIF.
    lv_gwt = ls_data-zsd_gwt.
    lv_nwt = ls_data-zsd_nwt.
    " New values of this event override the table values
    IF line_exists( ir_data_changed->mt_mod_cells[ row_id = lv_row fieldname = 'ZSD_GWT' ] ).
      ir_data_changed->get_cell_value( EXPORTING i_row_id = lv_row i_fieldname = 'ZSD_GWT'
                                       IMPORTING e_value  = lv_gwt ).
    ENDIF.
    IF line_exists( ir_data_changed->mt_mod_cells[ row_id = lv_row fieldname = 'ZSD_NWT' ] ).
      ir_data_changed->get_cell_value( EXPORTING i_row_id = lv_row i_fieldname = 'ZSD_NWT'
                                       IMPORTING e_value  = lv_nwt ).
    ENDIF.
    IF lv_gwt IS NOT INITIAL AND lv_nwt IS NOT INITIAL AND lv_gwt < lv_nwt.
      " 00/001 = "&1&2&3&4": text of the FS split into 50-character parts
      ir_data_changed->add_protocol_entry(
        i_msgid     = '00'
        i_msgty     = 'E'
        i_msgno     = '001'
        i_msgv1     = 'Invalid Weight: Gross Weight cannot be less than '
        i_msgv2     = 'Net Weight.'
        i_fieldname = 'ZSD_GWT'
        i_row_id    = lv_row ).
    ENDIF.
  ENDLOOP.

ENDFORM.

*&---------------------------------------------------------------------*
*& Toolbar functions
*&---------------------------------------------------------------------*
FORM user_command USING iv_ucomm    TYPE sy-ucomm
                        cs_selfield TYPE slis_selfield.

  DATA lo_grid  TYPE REF TO cl_gui_alv_grid.
  DATA lv_valid TYPE c LENGTH 1.

  CASE iv_ucomm.
    WHEN gc_fc_download.
      " Take over the entries of the grid into GT_DATA (runs DATA_CHANGED)
      CALL FUNCTION 'GET_GLOBALS_FROM_SLVC_FULLSCR'
        IMPORTING
          e_grid = lo_grid.
      IF lo_grid IS BOUND.
        lo_grid->check_changed_data( IMPORTING e_valid = lv_valid ).
        IF lv_valid = abap_false.
          MESSAGE 'Correct the marked entries before the download' TYPE 'S' DISPLAY LIKE 'E'.
          RETURN.
        ENDIF.
      ENDIF.
      PERFORM download_excel.
      cs_selfield-refresh    = abap_true.
      cs_selfield-row_stable = abap_true.
      cs_selfield-col_stable = abap_true.
  ENDCASE.

ENDFORM.

*&---------------------------------------------------------------------*
*& Excel file in the layout of sheet "Z table" (FORM epack_columns):
*& .xlsx written with SALV (headings = sheet headings) and saved on the
*& front end.
*&---------------------------------------------------------------------*
FORM download_excel.

  DATA lr_xl      TYPE REF TO data.
  DATA lo_salv    TYPE REF TO cl_salv_table.
  DATA lo_column  TYPE REF TO cl_salv_column.
  DATA lx_salv    TYPE REF TO cx_salv_error.
  DATA lv_xstring TYPE xstring.
  DATA lt_bin     TYPE solix_tab.
  DATA lv_size    TYPE i.
  DATA lv_name    TYPE lvc_fname.
  DATA lv_long    TYPE scrtext_l.
  DATA lv_medium  TYPE scrtext_m.
  DATA lv_short   TYPE scrtext_s.
  DATA lv_width   TYPE i.
  DATA lv_outlen  TYPE lvc_outlen.
  DATA lv_file    TYPE string.
  DATA lv_path    TYPE string.
  DATA lv_full    TYPE string.
  DATA lv_action  TYPE i.

  FIELD-SYMBOLS <lt_xl>    TYPE STANDARD TABLE.
  FIELD-SYMBOLS <lv_value> TYPE any.
  FIELD-SYMBOLS <lv_cell>  TYPE any.

  " Rows: one STRING cell per column of the layout
  PERFORM create_xl_table USING gt_cols CHANGING lr_xl.
  ASSIGN lr_xl->* TO <lt_xl>.
  LOOP AT gt_data ASSIGNING FIELD-SYMBOL(<ls_data>).
    APPEND INITIAL LINE TO <lt_xl> ASSIGNING FIELD-SYMBOL(<ls_xl>).
    LOOP AT gt_cols INTO DATA(ls_col).
      ASSIGN COMPONENT ls_col-field OF STRUCTURE <ls_data> TO <lv_value>.
      ASSIGN COMPONENT ls_col-pos   OF STRUCTURE <ls_xl>   TO <lv_cell>.
      IF <lv_value> IS ASSIGNED AND <lv_cell> IS ASSIGNED.
        PERFORM field_to_text USING <lv_value> CHANGING <lv_cell>.
      ENDIF.
      UNASSIGN: <lv_value>, <lv_cell>.
    ENDLOOP.
  ENDLOOP.

  TRY.
      cl_salv_table=>factory( IMPORTING r_salv_table = lo_salv
                              CHANGING  t_table      = <lt_xl> ).
      LOOP AT gt_cols INTO ls_col.
        lv_name   = |C{ ls_col-pos WIDTH = 2 ALIGN = RIGHT PAD = '0' }|.
        lo_column = lo_salv->get_columns( )->get_column( lv_name ).
        lv_long   = ls_col-header.
        lv_medium = ls_col-header.
        lv_short  = ls_col-header.
        lo_column->set_long_text( lv_long ).
        lo_column->set_medium_text( lv_medium ).
        lo_column->set_short_text( lv_short ).
        lo_column->set_fixed_header_text( 'L' ).
        PERFORM column_width USING ls_col-field ls_col-header CHANGING lv_width.
        lv_outlen = lv_width.
        lo_column->set_output_length( lv_outlen ).
      ENDLOOP.
      lv_xstring = lo_salv->to_xml( xml_type = if_salv_bs_xml=>c_type_xlsx ).
    CATCH cx_salv_error INTO lx_salv.
      MESSAGE lx_salv TYPE 'S' DISPLAY LIKE 'E'.
      RETURN.
  ENDTRY.

  " File name
  cl_gui_frontend_services=>file_save_dialog(
    EXPORTING
      default_extension = 'xlsx'
      default_file_name = |ZSD_EPACK_{ sy-datum }.xlsx|
      file_filter       = 'Excel (*.xlsx)|*.xlsx'
    CHANGING
      filename          = lv_file
      path              = lv_path
      fullpath          = lv_full
      user_action       = lv_action
    EXCEPTIONS
      OTHERS            = 1 ).
  IF sy-subrc <> 0 OR lv_action = cl_gui_frontend_services=>action_cancel OR lv_full IS INITIAL.
    MESSAGE 'Download cancelled' TYPE 'S'.
    RETURN.
  ENDIF.

  lt_bin  = cl_bcs_convert=>xstring_to_solix( iv_xstring = lv_xstring ).
  lv_size = xstrlen( lv_xstring ).
  cl_gui_frontend_services=>gui_download(
    EXPORTING
      bin_filesize = lv_size
      filename     = lv_full
      filetype     = 'BIN'
    CHANGING
      data_tab     = lt_bin
    EXCEPTIONS
      OTHERS       = 1 ).
  IF sy-subrc <> 0.
    MESSAGE |File { lv_full } could not be written| TYPE 'S' DISPLAY LIKE 'E'.
    RETURN.
  ENDIF.

  MESSAGE |{ lines( gt_data ) } rows downloaded to { lv_full }| TYPE 'S'.

ENDFORM.

*&---------------------------------------------------------------------*
*& Text elements (Goto -> Text elements)
*&   Text symbol    B01      Sales orders
*&   Selection text S_VBELN  Sales Order
*&---------------------------------------------------------------------*
