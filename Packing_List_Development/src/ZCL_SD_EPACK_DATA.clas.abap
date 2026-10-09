"! <p class="shorttext synchronized">Export Packing List - data provider</p>
"!
"! WRICEF 102-B "Export consolidated packing list".
"!
"! Builds the Smart Form data of MANY packing lists in one go, following
"! the FS logic sheet (Logic_sheet1.xlsx, sheet "Logic") row by row. The
"! row numbers (B3, B20, ...) are quoted at each derivation.
"!
"! PERFORMANCE CONTRACT
"!   READ_SAP_DATA reads every SAP table ONCE for the invoices of ALL
"!   packing lists (FOR ALL ENTRIES on de-duplicated keys). BUILD_ONE then
"!   works only on internal HASHED / SORTED tables. There is no SELECT
"!   inside any loop, whatever the number of packing lists, invoices or
"!   line items.
"!
"! ISOLATION
"!   Every packing list is assembled only from its own invoices and its
"!   own line items; errors of one packing list never stop the others.
"!
"! WHERE THE LOGIC SHEET NAMES A FIELD THAT DOES NOT EXIST IN THE NAMED
"! TABLE, the table that holds the field is read instead (documented in
"! docs/Packing_List_Report_Design.md, section 3):
"!   B4   BEZEI  is in T005U (T005S has no texts)
"!   B5   LANDX  is in T005T (T005 has no texts)
"!   B54  MAKTX  is in MAKT  (MARA has no texts)
"!   B20/B26  "PARVW = SP / SH" are the English display codes of the
"!        internal values AG / WE; the partner number is KNVP-KUNN2
"!   B13  AUBEL is a field of VBRP (billing item), not of VBRK
"! ⚠ SAP-VERIFY in DEV before activation: J_1BBRANCH-GSTIN, T001Z,
"!   T604N (key SPRAS/LAND1/STEUC, text TEXT1), TVZBT-ZTAGG, KNVP-DEFPA.
CLASS zcl_sd_epack_data DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC.

  PUBLIC SECTION.

    TYPES tr_packno TYPE RANGE OF zsd_packno.

    "! Prepare the form data of all packing lists.
    "! @parameter iv_format | ZIF_SD_EPACK=>GC_FORMAT-*
    "! @parameter rt_result | one entry per packing list - status, messages, form data
    METHODS prepare
      IMPORTING it_packing_lists TYPE zif_sd_epack=>tt_packing_list
                iv_format        TYPE char1 DEFAULT zif_sd_epack=>gc_format-auto
      RETURNING VALUE(rt_result) TYPE zif_sd_epack=>tt_result.

    "! Read saved packing lists (reprint) for invoices and/or numbers.
    "! A packing list found through one of its invoices is returned
    "! completely, with all its invoices and items.
    "! @raising zcx_sd_epack | nothing selected / nothing found
    METHODS load_saved
      IMPORTING it_vbeln         TYPE zif_sd_epack=>tr_vbeln
                it_packno        TYPE tr_packno
      EXPORTING et_packing_lists TYPE zif_sd_epack=>tt_packing_list
                et_messages      TYPE bapiret2_t
      RAISING   zcx_sd_epack.

  PRIVATE SECTION.

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
      tt_vbrk TYPE HASHED TABLE OF ty_vbrk WITH UNIQUE KEY vbeln,
      "! Checked invoices of ONE packing list (index access needed)
      tt_vbrk_list TYPE SORTED TABLE OF ty_vbrk WITH UNIQUE KEY vbeln,

      BEGIN OF ty_vbrp,
        vbeln TYPE vbrp-vbeln,
        posnr TYPE vbrp-posnr,
        matnr TYPE vbrp-matnr,
        werks TYPE vbrp-werks,
        aubel TYPE vbrp-aubel,
      END OF ty_vbrp,
      tt_vbrp TYPE SORTED TABLE OF ty_vbrp WITH UNIQUE KEY vbeln posnr,

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
      tt_knvp TYPE SORTED TABLE OF ty_knvp
              WITH UNIQUE KEY kunnr vkorg vtweg spart parvw parza,

      BEGIN OF ty_key_adrnr,
        id    TYPE char10,
        adrnr TYPE adrnr,
      END OF ty_key_adrnr,
      tt_key_adrnr TYPE HASHED TABLE OF ty_key_adrnr WITH UNIQUE KEY id,

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
      tt_adrc TYPE HASHED TABLE OF ty_adrc WITH UNIQUE KEY addrnumber,

      BEGIN OF ty_adr6,
        addrnumber TYPE adr6-addrnumber,
        flgdefault TYPE adr6-flgdefault,
        consnumber TYPE adr6-consnumber,
        smtp_addr  TYPE adr6-smtp_addr,
      END OF ty_adr6,
      tt_adr6 TYPE SORTED TABLE OF ty_adr6 WITH NON-UNIQUE KEY addrnumber,

      BEGIN OF ty_adr2,
        addrnumber TYPE adr2-addrnumber,
        r3_user    TYPE adr2-r3_user,
        flgdefault TYPE adr2-flgdefault,
        consnumber TYPE adr2-consnumber,
        telnr_long TYPE adr2-telnr_long,
      END OF ty_adr2,
      tt_adr2 TYPE SORTED TABLE OF ty_adr2 WITH NON-UNIQUE KEY addrnumber r3_user,

      BEGIN OF ty_t005t,
        land1 TYPE t005t-land1,
        landx TYPE t005t-landx,
      END OF ty_t005t,
      tt_t005t TYPE HASHED TABLE OF ty_t005t WITH UNIQUE KEY land1,

      BEGIN OF ty_t005u,
        land1 TYPE t005u-land1,
        bland TYPE t005u-bland,
        bezei TYPE t005u-bezei,
      END OF ty_t005u,
      tt_t005u TYPE HASHED TABLE OF ty_t005u WITH UNIQUE KEY land1 bland,

      BEGIN OF ty_branch,
        bukrs  TYPE j_1bbranch-bukrs,
        branch TYPE j_1bbranch-branch,
        gstin  TYPE j_1bbranch-gstin,
      END OF ty_branch,
      tt_branch TYPE HASHED TABLE OF ty_branch WITH UNIQUE KEY bukrs branch,

      BEGIN OF ty_t001z,
        bukrs TYPE t001z-bukrs,
        paval TYPE t001z-paval,
      END OF ty_t001z,
      tt_t001z TYPE HASHED TABLE OF ty_t001z WITH UNIQUE KEY bukrs,

      BEGIN OF ty_vbak,
        vbeln TYPE vbak-vbeln,
        vgbel TYPE vbak-vgbel,
      END OF ty_vbak,
      tt_vbak TYPE HASHED TABLE OF ty_vbak WITH UNIQUE KEY vbeln,

      BEGIN OF ty_vbkd,
        vbeln TYPE vbkd-vbeln,
        bstkd TYPE vbkd-bstkd,
        bstdk TYPE vbkd-bstdk,
      END OF ty_vbkd,
      tt_vbkd TYPE HASHED TABLE OF ty_vbkd WITH UNIQUE KEY vbeln,

      BEGIN OF ty_tvzbt,
        zterm TYPE tvzbt-zterm,
        ztagg TYPE tvzbt-ztagg,
        vtext TYPE tvzbt-vtext,
      END OF ty_tvzbt,
      tt_tvzbt TYPE SORTED TABLE OF ty_tvzbt WITH UNIQUE KEY zterm ztagg,

      BEGIN OF ty_makt,
        matnr TYPE makt-matnr,
        maktx TYPE makt-maktx,
      END OF ty_makt,
      tt_makt TYPE HASHED TABLE OF ty_makt WITH UNIQUE KEY matnr,

      BEGIN OF ty_marc,
        matnr TYPE marc-matnr,
        werks TYPE marc-werks,
        steuc TYPE marc-steuc,
      END OF ty_marc,
      tt_marc TYPE HASHED TABLE OF ty_marc WITH UNIQUE KEY matnr werks,

      BEGIN OF ty_t604n,
        steuc TYPE t604n-steuc,
        text1 TYPE t604n-text1,
      END OF ty_t604n,
      tt_t604n TYPE HASHED TABLE OF ty_t604n WITH UNIQUE KEY steuc,

      BEGIN OF ty_matnr,
        partno TYPE zsd_partno,
        matnr  TYPE matnr,
      END OF ty_matnr,
      tt_matnr TYPE HASHED TABLE OF ty_matnr WITH UNIQUE KEY partno,

      BEGIN OF ty_auth,
        object TYPE xuobject,
        value  TYPE char10,
        ok     TYPE abap_bool,
      END OF ty_auth,
      tt_auth TYPE HASHED TABLE OF ty_auth WITH UNIQUE KEY object value,

      "! Resolved partner (logic B20 / B26)
      BEGIN OF ty_partner,
        kunnr TYPE kunnr,
        adrnr TYPE adrnr,
      END OF ty_partner.

    " --- Buffers: filled once by READ_SAP_DATA -----------------------------
    DATA mt_vbrk   TYPE tt_vbrk.
    DATA mt_vbrp   TYPE tt_vbrp.
    DATA mt_knvp   TYPE tt_knvp.
    DATA mt_kna1   TYPE tt_key_adrnr.
    DATA mt_t001   TYPE tt_key_adrnr.
    DATA mt_adrc   TYPE tt_adrc.
    DATA mt_adr6   TYPE tt_adr6.
    DATA mt_adr2   TYPE tt_adr2.
    DATA mt_t005t  TYPE tt_t005t.
    DATA mt_t005u  TYPE tt_t005u.
    DATA mt_branch TYPE tt_branch.
    DATA mt_t001z  TYPE tt_t001z.
    DATA mt_vbak   TYPE tt_vbak.
    DATA mt_vbkd   TYPE tt_vbkd.
    DATA mt_tvzbt  TYPE tt_tvzbt.
    DATA mt_makt   TYPE tt_makt.
    DATA mt_marc   TYPE tt_marc.
    DATA mt_t604n  TYPE tt_t604n.
    DATA mt_matnr  TYPE tt_matnr.
    DATA mt_auth   TYPE tt_auth.
    DATA mr_fkart  TYPE RANGE OF fkart.

    " --- Working data of the packing list being built ----------------------
    DATA mt_messages TYPE bapiret2_t.

    METHODS add
      IMPORTING iv_type TYPE bapi_mtype DEFAULT 'E'
                iv_no   TYPE symsgno
                iv_v1   TYPE simple OPTIONAL
                iv_v2   TYPE simple OPTIONAL
                iv_v3   TYPE simple OPTIONAL
                iv_v4   TYPE simple OPTIONAL.

    METHODS read_sap_data
      IMPORTING it_packing_lists TYPE zif_sd_epack=>tt_packing_list.

    METHODS read_addresses
      IMPORTING it_adrnr TYPE tt_key_adrnr.

    METHODS build_one
      IMPORTING is_pl            TYPE zif_sd_epack=>ts_packing_list
                iv_format        TYPE char1
      RETURNING VALUE(rs_result) TYPE zif_sd_epack=>ts_result.

    METHODS check_invoices
      IMPORTING is_pl          TYPE zif_sd_epack=>ts_packing_list
      RETURNING VALUE(rt_vbrk) TYPE tt_vbrk_list.

    METHODS check_consistency
      IMPORTING it_vbrk TYPE tt_vbrk_list.

    METHODS check_authority
      IMPORTING it_vbrk TYPE tt_vbrk_list.

    METHODS partner
      IMPORTING is_vbrk           TYPE ty_vbrk
                iv_parvw          TYPE parvw
      RETURNING VALUE(rs_partner) TYPE ty_partner.

    METHODS determine_form
      IMPORTING is_header      TYPE zsd_epack_hdr
                iv_format      TYPE char1
      EXPORTING ev_np_count    TYPE i
      RETURNING VALUE(rv_form) TYPE tdsfname.

    METHODS goods_lines
      IMPORTING is_pl           TYPE zif_sd_epack=>ts_packing_list
                it_vbrp         TYPE tt_vbrp
      RETURNING VALUE(rt_lines) TYPE zif_sd_epack=>tt_text_line.

    METHODS resolve_items
      IMPORTING is_pl          TYPE zif_sd_epack=>ts_packing_list
                it_vbrp        TYPE tt_vbrp
                it_goods       TYPE zif_sd_epack=>tt_text_line
      EXPORTING et_items       TYPE zif_sd_epack=>tt_item
                et_headings    TYPE zif_sd_epack=>tt_hsn_heading.

    METHODS order_reference
      IMPORTING it_vbrp         TYPE tt_vbrp
      RETURNING VALUE(rt_lines) TYPE zif_sd_epack=>tt_text_line.

    METHODS exporter_lines
      IMPORTING iv_adrnr        TYPE adrnr
                iv_gstin        TYPE csequence
      RETURNING VALUE(rt_lines) TYPE zif_sd_epack=>tt_text_line.

    METHODS party_lines
      IMPORTING is_partner      TYPE ty_partner
      RETURNING VALUE(rt_lines) TYPE zif_sd_epack=>tt_text_line.

    METHODS country_text
      IMPORTING iv_land1       TYPE land1
      RETURNING VALUE(rv_text) TYPE string.

    METHODS matnr_of
      IMPORTING iv_partno       TYPE zsd_partno
      RETURNING VALUE(rv_matnr) TYPE matnr.

    CLASS-METHODS join
      IMPORTING it_parts       TYPE string_table
                iv_sep         TYPE string DEFAULT `, `
      RETURNING VALUE(rv_text) TYPE string.

    CLASS-METHODS append_block
      IMPORTING iv_block TYPE zsd_epack_block
                it_lines TYPE zif_sd_epack=>tt_text_line
      CHANGING  ct_texts TYPE zsd_tt_epack_prt_txt.

    CLASS-METHODS steuc_of
      IMPORTING iv_hsn          TYPE csequence
      RETURNING VALUE(rv_steuc) TYPE steuc.

ENDCLASS.


CLASS zcl_sd_epack_data IMPLEMENTATION.

  METHOD prepare.

    read_sap_data( it_packing_lists ).

    LOOP AT it_packing_lists INTO DATA(ls_pl).
      INSERT build_one( is_pl     = ls_pl
                        iv_format = iv_format ) INTO TABLE rt_result.
    ENDLOOP.

  ENDMETHOD.


  METHOD load_saved.

    DATA lt_packno TYPE SORTED TABLE OF zsd_packno WITH UNIQUE KEY table_line.
    DATA lt_hdr    TYPE STANDARD TABLE OF zsd_epack_hdr  WITH EMPTY KEY.
    DATA lt_inv    TYPE STANDARD TABLE OF zsd_epack_inv  WITH EMPTY KEY.
    DATA lt_item   TYPE STANDARD TABLE OF zsd_epack_data WITH EMPTY KEY.

    CLEAR: et_packing_lists, et_messages.

    " An empty range would select everything - never allowed here
    IF it_vbeln IS INITIAL AND it_packno IS INITIAL.
      zcx_sd_epack=>raise( iv_msgno = '003' ).
    ENDIF.

    IF it_vbeln IS NOT INITIAL.
      SELECT zsd_packno, vbeln FROM zsd_epack_inv
        WHERE vbeln IN @it_vbeln
        INTO TABLE @DATA(lt_link).
      LOOP AT lt_link INTO DATA(ls_link).
        INSERT ls_link-zsd_packno INTO TABLE lt_packno.
      ENDLOOP.
      " Invoices entered one by one must belong to a packing list
      LOOP AT it_vbeln INTO DATA(ls_sel) WHERE sign = 'I' AND option = 'EQ'.
        IF NOT line_exists( lt_link[ vbeln = ls_sel-low ] ).
          APPEND zcl_sd_epack_rules=>msg( iv_type = 'E' iv_no = '004'
                                          iv_v1 = |{ ls_sel-low ALPHA = OUT }| ) TO et_messages.
        ENDIF.
      ENDLOOP.
    ENDIF.

    IF it_packno IS NOT INITIAL.
      SELECT zsd_packno FROM zsd_epack_hdr
        WHERE zsd_packno IN @it_packno
        INTO TABLE @DATA(lt_hdr_key).
      LOOP AT lt_hdr_key INTO DATA(ls_hdr_key).
        INSERT ls_hdr_key-zsd_packno INTO TABLE lt_packno.
      ENDLOOP.
    ENDIF.

    IF lt_packno IS INITIAL.
      zcx_sd_epack=>raise( iv_msgno = '046' it_messages = et_messages ).
    ENDIF.

    SELECT * FROM zsd_epack_hdr
      FOR ALL ENTRIES IN @lt_packno
      WHERE zsd_packno = @lt_packno-table_line
      INTO TABLE @lt_hdr.

    SELECT * FROM zsd_epack_inv
      FOR ALL ENTRIES IN @lt_packno
      WHERE zsd_packno = @lt_packno-table_line
      INTO TABLE @lt_inv.

    SELECT * FROM zsd_epack_data
      FOR ALL ENTRIES IN @lt_packno
      WHERE zsd_packno = @lt_packno-table_line
      INTO TABLE @lt_item.

    SORT lt_inv  BY zsd_packno vbeln.
    SORT lt_item BY zsd_packno zsd_sn.

    LOOP AT lt_hdr INTO DATA(ls_hdr).
      DATA(ls_pl) = VALUE zif_sd_epack=>ts_packing_list( packno = ls_hdr-zsd_packno
                                                         header = ls_hdr ).
      LOOP AT lt_inv INTO DATA(ls_inv) WHERE zsd_packno = ls_hdr-zsd_packno.
        INSERT ls_inv-vbeln INTO TABLE ls_pl-invoices.
        IF it_vbeln IS NOT INITIAL AND ls_inv-vbeln NOT IN it_vbeln.
          APPEND zcl_sd_epack_rules=>msg( iv_type = 'I' iv_no = '006' iv_v1 = ls_hdr-zsd_packno
                                          iv_v2 = |{ ls_inv-vbeln ALPHA = OUT }| ) TO ls_pl-messages.
        ENDIF.
      ENDLOOP.
      LOOP AT lt_item INTO DATA(ls_item) WHERE zsd_packno = ls_hdr-zsd_packno.
        APPEND ls_item TO ls_pl-items.
      ENDLOOP.
      IF ls_pl-invoices IS INITIAL OR ls_pl-items IS INITIAL.
        APPEND zcl_sd_epack_rules=>msg( iv_type = 'E' iv_no = '047' iv_v1 = ls_hdr-zsd_packno ) TO ls_pl-messages.
      ENDIF.
      INSERT ls_pl INTO TABLE et_packing_lists.
    ENDLOOP.

  ENDMETHOD.


  METHOD add.
    APPEND zcl_sd_epack_rules=>msg( iv_type = iv_type iv_no = iv_no
                                    iv_v1 = iv_v1 iv_v2 = iv_v2
                                    iv_v3 = iv_v3 iv_v4 = iv_v4 ) TO mt_messages.
  ENDMETHOD.


  METHOD read_sap_data.

    TYPES: BEGIN OF ty_area,
             kunnr TYPE kunnr,
             vkorg TYPE vkorg,
             vtweg TYPE vtweg,
             spart TYPE spart,
           END OF ty_area.
    DATA lt_inv    TYPE SORTED TABLE OF vbeln_vf WITH UNIQUE KEY table_line.
    DATA lt_area   TYPE SORTED TABLE OF ty_area WITH UNIQUE KEY kunnr vkorg vtweg spart.
    DATA lt_kunnr  TYPE SORTED TABLE OF kunnr WITH UNIQUE KEY table_line.
    DATA lt_bukrs  TYPE SORTED TABLE OF bukrs WITH UNIQUE KEY table_line.
    DATA lt_adrnr  TYPE tt_key_adrnr.
    DATA lt_aubel  TYPE SORTED TABLE OF vbeln WITH UNIQUE KEY table_line.
    DATA lt_vgbel  TYPE SORTED TABLE OF vbeln WITH UNIQUE KEY table_line.
    DATA lt_zterm  TYPE SORTED TABLE OF dzterm WITH UNIQUE KEY table_line.
    DATA lt_matnr  TYPE SORTED TABLE OF matnr WITH UNIQUE KEY table_line.
    DATA lt_steuc  TYPE SORTED TABLE OF steuc WITH UNIQUE KEY table_line.
    DATA lt_branch TYPE SORTED TABLE OF ty_branch WITH UNIQUE KEY bukrs branch.

    " Invoices of all packing lists, each one once
    LOOP AT it_packing_lists INTO DATA(ls_pl).
      INSERT LINES OF ls_pl-invoices INTO TABLE lt_inv.
      LOOP AT ls_pl-items INTO DATA(ls_item) WHERE zsd_partno IS NOT INITIAL.
        IF NOT line_exists( mt_matnr[ partno = ls_item-zsd_partno ] ).
          INSERT VALUE #( partno = ls_item-zsd_partno matnr = matnr_of( ls_item-zsd_partno ) )
                 INTO TABLE mt_matnr.
        ENDIF.
        IF ls_item-zsd_hsn IS NOT INITIAL.
          INSERT steuc_of( ls_item-zsd_hsn ) INTO TABLE lt_steuc.
        ENDIF.
      ENDLOOP.
    ENDLOOP.
    IF lt_inv IS INITIAL.
      RETURN.
    ENDIF.

    " Billing documents - logic B3, B6, B11, B20, B26, B43, B44
    SELECT vbeln, fkart, fksto, sfakn, bukrs, vkorg, vtweg, spart,
           kunag, zterm, inco1, inco2, bupla
      FROM vbrk
      FOR ALL ENTRIES IN @lt_inv
      WHERE vbeln = @lt_inv-table_line
      INTO TABLE @mt_vbrk.

    " Billing items - logic B13 (AUBEL), B42 (MATNR)
    SELECT vbeln, posnr, matnr, werks, aubel
      FROM vbrp
      FOR ALL ENTRIES IN @lt_inv
      WHERE vbeln = @lt_inv-table_line
      INTO TABLE @mt_vbrp.

    " Allowed billing types (A11) - one read, empty = all allowed
    SELECT sign, opti AS option, low, high
      FROM tvarvc
      WHERE name = @zif_sd_epack=>gc_tvarv_fkart
        AND type = 'S'
      INTO CORRESPONDING FIELDS OF TABLE @mr_fkart.
    DELETE mr_fkart WHERE low IS INITIAL AND high IS INITIAL.

    LOOP AT mt_vbrk INTO DATA(ls_vbrk).
      INSERT VALUE #( kunnr = ls_vbrk-kunag vkorg = ls_vbrk-vkorg
                      vtweg = ls_vbrk-vtweg spart = ls_vbrk-spart ) INTO TABLE lt_area.
      INSERT ls_vbrk-bukrs INTO TABLE lt_bukrs.
      INSERT VALUE #( bukrs = ls_vbrk-bukrs branch = ls_vbrk-bupla ) INTO TABLE lt_branch.
      INSERT ls_vbrk-zterm INTO TABLE lt_zterm.
    ENDLOOP.

    LOOP AT mt_vbrp INTO DATA(ls_vbrp).
      IF ls_vbrp-aubel IS NOT INITIAL.
        INSERT ls_vbrp-aubel INTO TABLE lt_aubel.
      ENDIF.
    ENDLOOP.

    " Partners - logic B20 / B26: KNVP with KUNNR = KUNAG, PARVW SP / SH
    IF lt_area IS NOT INITIAL.
      SELECT kunnr, vkorg, vtweg, spart, parvw, parza, kunn2, defpa
        FROM knvp
        FOR ALL ENTRIES IN @lt_area
        WHERE kunnr = @lt_area-kunnr
          AND vkorg = @lt_area-vkorg
          AND vtweg = @lt_area-vtweg
          AND spart = @lt_area-spart
          AND ( parvw = @zif_sd_epack=>gc_parvw-sold_to OR parvw = @zif_sd_epack=>gc_parvw-ship_to )
        INTO TABLE @mt_knvp.
    ENDIF.

    LOOP AT mt_knvp INTO DATA(ls_knvp).
      INSERT ls_knvp-kunn2 INTO TABLE lt_kunnr.
    ENDLOOP.

    IF lt_kunnr IS NOT INITIAL.
      SELECT kunnr AS id, adrnr FROM kna1
        FOR ALL ENTRIES IN @lt_kunnr
        WHERE kunnr = @lt_kunnr-table_line
        INTO TABLE @mt_kna1.
    ENDIF.

    " Exporter address - logic B3 / B11
    IF lt_bukrs IS NOT INITIAL.
      SELECT bukrs AS id, adrnr FROM t001
        FOR ALL ENTRIES IN @lt_bukrs
        WHERE bukrs = @lt_bukrs-table_line
        INTO TABLE @mt_t001.

      " CIN - logic B7
      SELECT bukrs, paval FROM t001z
        FOR ALL ENTRIES IN @lt_bukrs
        WHERE bukrs = @lt_bukrs-table_line
          AND party = @zif_sd_epack=>gc_cin_party
        INTO TABLE @mt_t001z.
    ENDIF.

    " GST No. - logic B6
    IF lt_branch IS NOT INITIAL.
      SELECT bukrs, branch, gstin FROM j_1bbranch
        FOR ALL ENTRIES IN @lt_branch
        WHERE bukrs  = @lt_branch-bukrs
          AND branch = @lt_branch-branch
        INTO TABLE @mt_branch.
    ENDIF.

    LOOP AT mt_kna1 INTO DATA(ls_key).
      INSERT VALUE #( id = ls_key-adrnr adrnr = ls_key-adrnr ) INTO TABLE lt_adrnr.
    ENDLOOP.
    LOOP AT mt_t001 INTO ls_key.
      INSERT VALUE #( id = ls_key-adrnr adrnr = ls_key-adrnr ) INTO TABLE lt_adrnr.
    ENDLOOP.
    DELETE lt_adrnr WHERE adrnr IS INITIAL.
    read_addresses( lt_adrnr ).

    " Order no. & date - logic B13: VBRP-AUBEL -> VBAK-VGBEL -> VBKD
    IF lt_aubel IS NOT INITIAL.
      SELECT vbeln, vgbel FROM vbak
        FOR ALL ENTRIES IN @lt_aubel
        WHERE vbeln = @lt_aubel-table_line
        INTO TABLE @mt_vbak.
      LOOP AT mt_vbak INTO DATA(ls_vbak) WHERE vgbel IS NOT INITIAL.
        INSERT ls_vbak-vgbel INTO TABLE lt_vgbel.
      ENDLOOP.
    ENDIF.
    IF lt_vgbel IS NOT INITIAL.
      " Header business data of the quotation (item number 000000)
      SELECT vbeln, bstkd, bstdk FROM vbkd
        FOR ALL ENTRIES IN @lt_vgbel
        WHERE vbeln = @lt_vgbel-table_line
          AND posnr = '000000'
        INTO TABLE @mt_vbkd.
    ENDIF.

    " Payment term text - logic B43
    IF lt_zterm IS NOT INITIAL.
      SELECT zterm, ztagg, vtext FROM tvzbt
        FOR ALL ENTRIES IN @lt_zterm
        WHERE spras = @zif_sd_epack=>gc_langu
          AND zterm = @lt_zterm-table_line
        INTO TABLE @mt_tvzbt.
    ENDIF.

    " Materials: billed materials (B42) + Part No. of the line items (B54, B58)
    LOOP AT mt_vbrp INTO ls_vbrp WHERE matnr IS NOT INITIAL.
      INSERT ls_vbrp-matnr INTO TABLE lt_matnr.
    ENDLOOP.
    LOOP AT mt_matnr INTO DATA(ls_matnr) WHERE matnr IS NOT INITIAL.
      INSERT ls_matnr-matnr INTO TABLE lt_matnr.
    ENDLOOP.

    IF lt_matnr IS NOT INITIAL.
      SELECT matnr, maktx FROM makt
        FOR ALL ENTRIES IN @lt_matnr
        WHERE matnr = @lt_matnr-table_line
          AND spras = @zif_sd_epack=>gc_langu
        INTO TABLE @mt_makt.
    ENDIF.

    " HSN of the billed material in the billing plant (B42, B58)
    IF mt_vbrp IS NOT INITIAL.
      SELECT matnr, werks, steuc FROM marc
        FOR ALL ENTRIES IN @mt_vbrp
        WHERE matnr = @mt_vbrp-matnr
          AND werks = @mt_vbrp-werks
        INTO TABLE @mt_marc.
    ENDIF.

    LOOP AT mt_marc INTO DATA(ls_marc) WHERE steuc IS NOT INITIAL.
      INSERT ls_marc-steuc INTO TABLE lt_steuc.
    ENDLOOP.

    " HSN texts - logic B42: T604N with SPRAS = EN, LAND1 = IN
    IF lt_steuc IS NOT INITIAL.
      SELECT steuc, text1 FROM t604n
        FOR ALL ENTRIES IN @lt_steuc
        WHERE spras = @zif_sd_epack=>gc_langu
          AND land1 = @zif_sd_epack=>gc_hsn_country
          AND steuc = @lt_steuc-table_line
        INTO TABLE @mt_t604n.
    ENDIF.

  ENDMETHOD.


  METHOD read_addresses.

    DATA lt_adrc TYPE STANDARD TABLE OF ty_adrc WITH EMPTY KEY.
    TYPES: BEGIN OF ty_region,
             land1 TYPE land1,
             bland TYPE regio,
           END OF ty_region.
    DATA lt_land   TYPE SORTED TABLE OF land1 WITH UNIQUE KEY table_line.
    DATA lt_region TYPE SORTED TABLE OF ty_region WITH UNIQUE KEY land1 bland.

    IF it_adrnr IS INITIAL.
      RETURN.
    ENDIF.

    " Logic B3 / B20 / B26 - address fields
    SELECT addrnumber, date_from, name1, street, str_suppl1, city1, city2,
           post_code1, region, country, building
      FROM adrc
      FOR ALL ENTRIES IN @it_adrnr
      WHERE addrnumber = @it_adrnr-adrnr
        AND nation     = @space
        AND date_from <= @sy-datum
      INTO TABLE @lt_adrc.

    " Latest valid version of each address
    SORT lt_adrc BY addrnumber date_from DESCENDING.
    DELETE ADJACENT DUPLICATES FROM lt_adrc COMPARING addrnumber.
    mt_adrc = lt_adrc.

    " Logic B21 / B27 - e-mail
    SELECT addrnumber, flgdefault, consnumber, smtp_addr
      FROM adr6
      FOR ALL ENTRIES IN @it_adrnr
      WHERE addrnumber = @it_adrnr-adrnr
        AND persnumber = @space
      INTO TABLE @mt_adr6.

    " Logic B22 / B23 / B28 / B29 - telephone (R3_USER 1) and mobile (3)
    SELECT addrnumber, r3_user, flgdefault, consnumber, telnr_long
      FROM adr2
      FOR ALL ENTRIES IN @it_adrnr
      WHERE addrnumber = @it_adrnr-adrnr
        AND persnumber = @space
        AND ( r3_user = '1' OR r3_user = '3' )
      INTO TABLE @mt_adr2.

    " Logic B4 / B5 - country and region texts
    LOOP AT mt_adrc INTO DATA(ls_adrc).
      IF ls_adrc-country IS NOT INITIAL.
        INSERT ls_adrc-country INTO TABLE lt_land.
        IF ls_adrc-region IS NOT INITIAL.
          INSERT VALUE #( land1 = ls_adrc-country bland = ls_adrc-region ) INTO TABLE lt_region.
        ENDIF.
      ENDIF.
    ENDLOOP.

    IF lt_land IS NOT INITIAL.
      SELECT land1, landx FROM t005t
        FOR ALL ENTRIES IN @lt_land
        WHERE spras = @zif_sd_epack=>gc_langu
          AND land1 = @lt_land-table_line
        INTO TABLE @mt_t005t.
    ENDIF.

    IF lt_region IS NOT INITIAL.
      SELECT land1, bland, bezei FROM t005u
        FOR ALL ENTRIES IN @lt_region
        WHERE spras = @zif_sd_epack=>gc_langu
          AND land1 = @lt_region-land1
          AND bland = @lt_region-bland
        INTO TABLE @mt_t005u.
    ENDIF.

  ENDMETHOD.


  METHOD build_one.

    DATA lt_vbrp     TYPE tt_vbrp.
    DATA lt_items    TYPE zif_sd_epack=>tt_item.
    DATA lt_headings TYPE zif_sd_epack=>tt_hsn_heading.
    DATA lt_vbrk     TYPE tt_vbrk_list.
    DATA lv_np_count TYPE i.
    DATA ls_vbrk     TYPE ty_vbrk.
    DATA ls_sold_to  TYPE ty_partner.
    DATA ls_ship_to  TYPE ty_partner.

    mt_messages = is_pl-messages.

    rs_result-packno     = is_pl-packno.
    rs_result-inv_count  = lines( is_pl-invoices ).
    rs_result-item_count = lines( is_pl-items ).

    " Errors from the upload (structure, mandatory data, weights) stop
    " this packing list - but no other one
    IF NOT zcl_sd_epack_rules=>has_errors( mt_messages ).

      lt_vbrk = check_invoices( is_pl ).
      check_consistency( lt_vbrk ).
      check_authority( lt_vbrk ).

      " FS 2.1 - GW >= NW at output time (again, for saved data)
      IF NOT line_exists( mt_messages[ number = '001' ] ).
        APPEND LINES OF zcl_sd_epack_rules=>check_weights( is_pl-items ) TO mt_messages.
      ENDIF.

      determine_form( EXPORTING is_header   = is_pl-header
                                iv_format   = iv_format
                      IMPORTING ev_np_count = lv_np_count
                      RECEIVING rv_form     = rs_result-print-form ).

      " ---- All invoices agree (CHECK_CONSISTENCY), so the header values
      "      of any of them are the values of all of them -----------------
      IF NOT zcl_sd_epack_rules=>has_errors( mt_messages ) AND lt_vbrk IS NOT INITIAL.
        ls_vbrk    = lt_vbrk[ 1 ].
        " Logic B20 / B26 - a missing partner is an error (no consignee)
        ls_sold_to = partner( is_vbrk = ls_vbrk iv_parvw = zif_sd_epack=>gc_parvw-sold_to ).
        ls_ship_to = partner( is_vbrk = ls_vbrk iv_parvw = zif_sd_epack=>gc_parvw-ship_to ).
      ENDIF.
    ENDIF.

    IF zcl_sd_epack_rules=>has_errors( mt_messages ).
      rs_result-status   = zif_sd_epack=>gc_status-error.
      rs_result-messages = mt_messages.
      RETURN.
    ENDIF.

    " Billing items of THIS packing list's invoices only
    LOOP AT is_pl-invoices INTO DATA(lv_vbeln).
      LOOP AT mt_vbrp INTO DATA(ls_vbrp) WHERE vbeln = lv_vbeln.
        INSERT ls_vbrp INTO TABLE lt_vbrp.
      ENDLOOP.
    ENDLOOP.

    DATA(lv_company_adrnr) = VALUE adrnr( mt_t001[ id = ls_vbrk-bukrs ]-adrnr OPTIONAL ).
    DATA(ls_company)       = VALUE ty_adrc( mt_adrc[ addrnumber = lv_company_adrnr ] OPTIONAL ).
    DATA(lv_gstin)         = VALUE ty_branch-gstin( mt_branch[ bukrs = ls_vbrk-bukrs branch = ls_vbrk-bupla ]-gstin OPTIONAL ).

    DATA(ls_header) = VALUE zsd_s_epack_prt_hdr(
      packno     = is_pl-header-zsd_packno
      packdt_txt = zcl_sd_epack_rules=>format_date( is_pl-header-zsd_packdt )
      revno      = is_pl-header-zsd_revno
      gstin      = lv_gstin
      cin        = VALUE #( mt_t001z[ bukrs = ls_vbrk-bukrs ]-paval OPTIONAL )
      cnty_orgn  = is_pl-header-zsd_cnty_orgn
      cnty_dest  = country_text( VALUE #( mt_adrc[ addrnumber = ls_ship_to-adrnr ]-country OPTIONAL ) )
      same_party = xsdbool( ls_sold_to-kunnr = ls_ship_to-kunnr )
      prec       = is_pl-header-zsd_prec
      rec_pc     = is_pl-header-zsd_rec_pc
      vsl_flt    = is_pl-header-zsd_vsl_flt
      pol        = is_pl-header-zsd_pol
      pod        = is_pl-header-zsd_pod
      pld        = is_pl-header-zsd_pld
      np_count   = lv_np_count ).

    " Exporter's Ref. - logic B11: company address field BUILDING
    IF ls_company-building IS NOT INITIAL.
      ls_header-iec = COND #( WHEN ls_company-building CS 'IEC'
                              THEN ls_company-building
                              ELSE |IEC:{ ls_company-building }| ).
    ENDIF.

    IF lv_gstin IS INITIAL.
      add( iv_type = 'W' iv_no = '049' iv_v1 = 'GST No.' ).
    ENDIF.
    IF ls_header-cin IS INITIAL.
      add( iv_type = 'W' iv_no = '049' iv_v1 = 'CIN No.' ).
    ENDIF.
    IF ls_header-iec IS INITIAL.
      add( iv_type = 'W' iv_no = '049' iv_v1 = 'IEC (Exporter''s Ref.)' ).
    ENDIF.

    " ---- Description of goods (B42) and line items (B47-B58) ------------
    DATA(lt_goods) = goods_lines( is_pl = is_pl it_vbrp = lt_vbrp ).

    resolve_items( EXPORTING is_pl       = is_pl
                             it_vbrp     = lt_vbrp
                             it_goods    = lt_goods
                   IMPORTING et_items    = lt_items
                             et_headings = lt_headings ).

    DATA(ls_totals) = zcl_sd_epack_rules=>calc_totals( lt_items ).
    ls_header-tot_art_txt = |{ ls_totals-articles }|.
    ls_header-tot_qty_txt = zcl_sd_epack_rules=>format_quantity( ls_totals-quantity ).
    ls_header-tot_gwt_txt = zcl_sd_epack_rules=>format_weight( ls_totals-gross ).
    ls_header-tot_nwt_txt = zcl_sd_epack_rules=>format_weight( ls_totals-net ).
    ls_header-tot_text    = |({ zcl_sd_epack_rules=>total_text( ls_totals ) })|.

    " ---- Text blocks -----------------------------------------------------
    DATA lt_texts TYPE zsd_tt_epack_prt_txt.

    append_block( EXPORTING iv_block = zif_sd_epack=>gc_block-exporter
                            it_lines = exporter_lines( iv_adrnr = lv_company_adrnr iv_gstin = lv_gstin )
                  CHANGING  ct_texts = lt_texts ).
    append_block( EXPORTING iv_block = zif_sd_epack=>gc_block-sold_to
                            it_lines = party_lines( ls_sold_to )
                  CHANGING  ct_texts = lt_texts ).
    append_block( EXPORTING iv_block = zif_sd_epack=>gc_block-ship_to
                            it_lines = party_lines( ls_ship_to )
                  CHANGING  ct_texts = lt_texts ).

    " Logic B31 - notify parties 1-5
    DO zif_sd_epack=>gc_max_np TIMES.
      DATA(lv_np) = sy-index.
      ASSIGN COMPONENT |ZSD_NP{ lv_np }| OF STRUCTURE is_pl-header TO FIELD-SYMBOL(<lv_np>).
      IF sy-subrc = 0.
        append_block( EXPORTING iv_block = CONV #( |NP{ lv_np }| )
                                it_lines = zcl_sd_epack_rules=>split_lines( <lv_np> )
                      CHANGING  ct_texts = lt_texts ).
      ENDIF.
    ENDDO.

    append_block( EXPORTING iv_block = zif_sd_epack=>gc_block-order_ref
                            it_lines = order_reference( lt_vbrp )
                  CHANGING  ct_texts = lt_texts ).

    " Logic B14 - Buyer's reference no. & date
    DATA(lv_buyer_ref) = condense( CONV string( is_pl-header-zsd_party_ref ) ).
    IF is_pl-header-zsd_ref_dt IS NOT INITIAL.
      lv_buyer_ref = |{ lv_buyer_ref } DATE: { zcl_sd_epack_rules=>format_date( is_pl-header-zsd_ref_dt ) }|.
    ENDIF.
    append_block( EXPORTING iv_block = zif_sd_epack=>gc_block-buyer_ref
                            it_lines = zcl_sd_epack_rules=>split_lines( lv_buyer_ref )
                  CHANGING  ct_texts = lt_texts ).

    " Logic B43 / B44 - payment term text, Incoterms 1 + 2
    DATA(lv_vtext) = VALUE tvzbt-vtext( ).
    LOOP AT mt_tvzbt INTO DATA(ls_tvzbt) WHERE zterm = ls_vbrk-zterm.
      lv_vtext = ls_tvzbt-vtext.          " first day limit = lowest ZTAGG
      EXIT.
    ENDLOOP.
    append_block( EXPORTING iv_block = zif_sd_epack=>gc_block-payment
                            it_lines = VALUE #( ( |PAYMENT TERM: { COND string( WHEN lv_vtext IS NOT INITIAL
                                                                                THEN lv_vtext
                                                                                ELSE ls_vbrk-zterm ) }| )
                                                ( condense( |TERMS OF SHIPMENT :- { ls_vbrk-inco1 } { ls_vbrk-inco2 }| ) ) )
                  CHANGING  ct_texts = lt_texts ).

    " Logic B39 / B40 / B41 / B61
    append_block( EXPORTING iv_block = zif_sd_epack=>gc_block-marks
                            it_lines = zcl_sd_epack_rules=>split_lines( is_pl-header-zsd_mark )
                  CHANGING  ct_texts = lt_texts ).
    append_block( EXPORTING iv_block = zif_sd_epack=>gc_block-container
                            it_lines = zcl_sd_epack_rules=>split_lines( is_pl-header-zsd_cont )
                  CHANGING  ct_texts = lt_texts ).
    append_block( EXPORTING iv_block = zif_sd_epack=>gc_block-packages
                            it_lines = zcl_sd_epack_rules=>split_lines( is_pl-header-zsd_nopack )
                  CHANGING  ct_texts = lt_texts ).
    append_block( EXPORTING iv_block = zif_sd_epack=>gc_block-goods
                            it_lines = lt_goods
                  CHANGING  ct_texts = lt_texts ).
    append_block( EXPORTING iv_block = zif_sd_epack=>gc_block-declaration
                            it_lines = zcl_sd_epack_rules=>split_lines( is_pl-header-zsd_pl_decl )
                  CHANGING  ct_texts = lt_texts ).

    rs_result-bukrs        = ls_vbrk-bukrs.
    rs_result-print-header = ls_header.
    rs_result-print-texts  = lt_texts.
    rs_result-print-items  = zcl_sd_epack_rules=>build_print_items( it_items    = lt_items
                                                                    it_headings = lt_headings ).

    rs_result-messages = mt_messages.
    rs_result-status   = COND #( WHEN line_exists( mt_messages[ type = 'W' ] )
                                 THEN zif_sd_epack=>gc_status-warning
                                 ELSE zif_sd_epack=>gc_status-success ).

  ENDMETHOD.


  METHOD check_invoices.

    LOOP AT is_pl-invoices INTO DATA(lv_vbeln).
      DATA(lv_ext) = |{ lv_vbeln ALPHA = OUT }|.
      READ TABLE mt_vbrk INTO DATA(ls_vbrk) WITH TABLE KEY vbeln = lv_vbeln.
      IF sy-subrc <> 0.
        add( iv_no = '008' iv_v1 = lv_ext ).
        CONTINUE.
      ENDIF.
      IF ls_vbrk-fksto = abap_true OR ls_vbrk-sfakn IS NOT INITIAL.
        add( iv_no = '009' iv_v1 = lv_ext ).
        CONTINUE.
      ENDIF.
      IF mr_fkart IS NOT INITIAL AND ls_vbrk-fkart NOT IN mr_fkart.
        add( iv_no = '010' iv_v1 = lv_ext iv_v2 = ls_vbrk-fkart ).
        CONTINUE.
      ENDIF.
      INSERT ls_vbrk INTO TABLE rt_vbrk.
    ENDLOOP.

    IF rt_vbrk IS INITIAL AND NOT zcl_sd_epack_rules=>has_errors( mt_messages ).
      add( iv_no = '047' iv_v1 = is_pl-packno ).
    ENDIF.

  ENDMETHOD.


  METHOD check_consistency.

    " One packing list prints ONE exporter, GSTIN, buyer, consignee,
    " payment term and Incoterm. The logic sheet reads them "from VBRK";
    " if the invoices of a packing list disagree, no single value is
    " correct, so the packing list is rejected rather than printing the
    " value of an arbitrary invoice (A04 / Q04).
    IF lines( it_vbrk ) < 2.
      RETURN.
    ENDIF.

    DATA(lt_vbrk) = it_vbrk.
    DATA(ls_first) = VALUE ty_vbrk( ).
    LOOP AT lt_vbrk INTO DATA(ls_vbrk).
      IF sy-tabix = 1.
        ls_first = ls_vbrk.
        CONTINUE.
      ENDIF.
      DATA(lv_inv) = |{ ls_first-vbeln ALPHA = OUT }/{ ls_vbrk-vbeln ALPHA = OUT }|.
      IF ls_vbrk-bukrs <> ls_first-bukrs.
        add( iv_no = '011' iv_v1 = 'Company code' iv_v2 = |{ ls_first-bukrs } / { ls_vbrk-bukrs }| iv_v3 = lv_inv ).
      ENDIF.
      IF ls_vbrk-bupla <> ls_first-bupla.
        add( iv_no = '011' iv_v1 = 'Business place (GSTIN)' iv_v2 = |{ ls_first-bupla } / { ls_vbrk-bupla }| iv_v3 = lv_inv ).
      ENDIF.
      IF ls_vbrk-kunag <> ls_first-kunag.
        add( iv_no = '011' iv_v1 = 'Sold-to party' iv_v2 = |{ ls_first-kunag ALPHA = OUT } / { ls_vbrk-kunag ALPHA = OUT }| iv_v3 = lv_inv ).
      ENDIF.
      IF ls_vbrk-vkorg <> ls_first-vkorg OR ls_vbrk-vtweg <> ls_first-vtweg OR ls_vbrk-spart <> ls_first-spart.
        add( iv_no = '011' iv_v1 = 'Sales area' iv_v2 = |{ ls_first-vkorg }/{ ls_first-vtweg }/{ ls_first-spart } / {
                                                          ls_vbrk-vkorg }/{ ls_vbrk-vtweg }/{ ls_vbrk-spart }| iv_v3 = lv_inv ).
      ENDIF.
      IF ls_vbrk-zterm <> ls_first-zterm.
        add( iv_no = '011' iv_v1 = 'Payment terms' iv_v2 = |{ ls_first-zterm } / { ls_vbrk-zterm }| iv_v3 = lv_inv ).
      ENDIF.
      IF ls_vbrk-inco1 <> ls_first-inco1 OR ls_vbrk-inco2 <> ls_first-inco2.
        add( iv_no = '011' iv_v1 = 'Incoterms' iv_v2 = |{ ls_first-inco1 } { ls_first-inco2 } / { ls_vbrk-inco1 } { ls_vbrk-inco2 }|
                                               iv_v3 = lv_inv ).
      ENDIF.
    ENDLOOP.

  ENDMETHOD.


  METHOD check_authority.

    " FS 2.3 is empty - A24. Results are buffered per value.
    LOOP AT it_vbrk INTO DATA(ls_vbrk).

      READ TABLE mt_auth INTO DATA(ls_auth)
           WITH TABLE KEY object = zif_sd_epack=>gc_auth_object value = ls_vbrk-bukrs.
      IF sy-subrc <> 0.
        AUTHORITY-CHECK OBJECT zif_sd_epack=>gc_auth_object
          ID 'ACTVT' FIELD zif_sd_epack=>gc_actvt-print
          ID 'BUKRS' FIELD ls_vbrk-bukrs.
        ls_auth = VALUE #( object = zif_sd_epack=>gc_auth_object value = ls_vbrk-bukrs ok = xsdbool( sy-subrc = 0 ) ).
        INSERT ls_auth INTO TABLE mt_auth.
      ENDIF.
      IF ls_auth-ok = abap_false.
        add( iv_no = '012' iv_v1 = 'Company code' iv_v2 = ls_vbrk-bukrs iv_v3 = zif_sd_epack=>gc_actvt-print ).
      ENDIF.

      READ TABLE mt_auth INTO ls_auth WITH TABLE KEY object = 'V_VBRK_VKO' value = ls_vbrk-vkorg.
      IF sy-subrc <> 0.
        AUTHORITY-CHECK OBJECT 'V_VBRK_VKO'
          ID 'VKORG' FIELD ls_vbrk-vkorg
          ID 'ACTVT' FIELD zif_sd_epack=>gc_actvt-print.
        ls_auth = VALUE #( object = 'V_VBRK_VKO' value = ls_vbrk-vkorg ok = xsdbool( sy-subrc = 0 ) ).
        INSERT ls_auth INTO TABLE mt_auth.
      ENDIF.
      IF ls_auth-ok = abap_false.
        add( iv_no = '012' iv_v1 = 'Sales org.' iv_v2 = ls_vbrk-vkorg iv_v3 = zif_sd_epack=>gc_actvt-print ).
      ENDIF.

    ENDLOOP.

    " Report each missing authorization once per packing list
    SORT mt_messages BY type id number message_v1 message_v2 message_v3 message_v4.
    DELETE ADJACENT DUPLICATES FROM mt_messages COMPARING type id number message_v1 message_v2 message_v3 message_v4.

  ENDMETHOD.


  METHOD partner.

    " Logic B20 / B26: KNVP with KUNNR = KUNAG (and the sales area of the
    " invoice, which is part of the KNVP key), PARVW = SP / SH -> KUNN2.
    " Several partners of one function: default partner (DEFPA), else
    " the first by counter PARZA (A31 / Q07).
    DATA lv_count TYPE i.
    DATA ls_hit   TYPE ty_knvp.

    LOOP AT mt_knvp INTO DATA(ls_knvp)
         WHERE kunnr = is_vbrk-kunag
           AND vkorg = is_vbrk-vkorg
           AND vtweg = is_vbrk-vtweg
           AND spart = is_vbrk-spart
           AND parvw = iv_parvw.
      lv_count += 1.
      IF ls_hit IS INITIAL OR ( ls_knvp-defpa = abap_true AND ls_hit-defpa = abap_false ).
        ls_hit = ls_knvp.
      ENDIF.
    ENDLOOP.

    DATA(lv_function) = COND string( WHEN iv_parvw = zif_sd_epack=>gc_parvw-sold_to THEN `Sold-to (SP)`
                                                                                    ELSE `Ship-to (SH)` ).
    IF lv_count = 0.
      add( iv_no = '042' iv_v1 = lv_function iv_v2 = |{ is_vbrk-kunag ALPHA = OUT }|
                                       iv_v3 = |{ is_vbrk-vkorg }/{ is_vbrk-vtweg }/{ is_vbrk-spart }| ).
      RETURN.
    ELSEIF lv_count > 1.
      add( iv_type = 'W' iv_no = '060' iv_v1 = |{ is_vbrk-kunag ALPHA = OUT }| iv_v2 = lv_count
                                       iv_v3 = lv_function iv_v4 = |{ ls_hit-kunn2 ALPHA = OUT }| ).
    ENDIF.

    rs_partner = VALUE #( kunnr = ls_hit-kunn2
                          adrnr = VALUE #( mt_kna1[ id = ls_hit-kunn2 ]-adrnr OPTIONAL ) ).

  ENDMETHOD.


  METHOD determine_form.

    " FS 1.1: format with notify party when notify parties are maintained
    " (A05). A manual "without" while notify parties exist is an error.
    ev_np_count = 0.
    DO zif_sd_epack=>gc_max_np TIMES.
      ASSIGN COMPONENT |ZSD_NP{ sy-index }| OF STRUCTURE is_header TO FIELD-SYMBOL(<lv_np>).
      IF sy-subrc = 0 AND <lv_np> IS NOT INITIAL.
        ev_np_count += 1.
      ENDIF.
    ENDDO.

    CASE iv_format.
      WHEN zif_sd_epack=>gc_format-with_np.
        rv_form = zif_sd_epack=>gc_form-with_np.
      WHEN zif_sd_epack=>gc_format-without_np.
        IF ev_np_count > 0.
          add( iv_no = '013' iv_v1 = is_header-zsd_packno ).
        ENDIF.
        rv_form = zif_sd_epack=>gc_form-without_np.
      WHEN OTHERS.
        rv_form = COND #( WHEN ev_np_count > 0 THEN zif_sd_epack=>gc_form-with_np
                                               ELSE zif_sd_epack=>gc_form-without_np ).
    ENDCASE.

  ENDMETHOD.


  METHOD goods_lines.

    " Uploaded "Description of goods" (template column AG, field
    " ZSD_HSN_DESC - "Same HSN having different description") is used
    " when filled (A14 / Q05) ...
    rt_lines = zcl_sd_epack_rules=>split_lines( is_pl-header-zsd_hsn_desc ).
    IF rt_lines IS NOT INITIAL.
      RETURN.
    ENDIF.

    " ... otherwise logic B42: VBRP-MATNR -> MARC-STEUC -> T604N-TEXT1,
    " "club common HSN code and description in all billing documents"
    DATA lt_seen TYPE SORTED TABLE OF steuc WITH UNIQUE KEY table_line.

    LOOP AT it_vbrp INTO DATA(ls_vbrp) WHERE matnr IS NOT INITIAL.
      DATA(lv_steuc) = VALUE steuc( mt_marc[ matnr = ls_vbrp-matnr werks = ls_vbrp-werks ]-steuc OPTIONAL ).
      IF lv_steuc IS INITIAL.
        CONTINUE.
      ENDIF.
      INSERT lv_steuc INTO TABLE lt_seen.
      IF sy-subrc <> 0.
        CONTINUE.                         " HSN already clubbed
      ENDIF.
      DATA(lv_text) = VALUE t604n-text1( mt_t604n[ steuc = lv_steuc ]-text1 OPTIONAL ).
      APPEND COND string( WHEN lv_text IS INITIAL
                          THEN |HS CODE : { zcl_sd_epack_rules=>format_hsn( lv_steuc ) }|
                          ELSE |{ lv_text } - HS CODE : { zcl_sd_epack_rules=>format_hsn( lv_steuc ) }| )
             TO rt_lines.
    ENDLOOP.

  ENDMETHOD.


  METHOD resolve_items.

    et_items = is_pl-items.

    LOOP AT et_items ASSIGNING FIELD-SYMBOL(<ls_item>).

      DATA(lv_sn)    = |{ CONV i( <ls_item>-zsd_sn ) }|.
      DATA(lv_matnr) = VALUE matnr( mt_matnr[ partno = <ls_item>-zsd_partno ]-matnr OPTIONAL ).

      " Plant of the material in this packing list's invoices
      DATA(lv_werks) = VALUE werks_d( ).
      LOOP AT it_vbrp INTO DATA(ls_vbrp) WHERE matnr = lv_matnr.
        lv_werks = ls_vbrp-werks.
        EXIT.
      ENDLOOP.
      IF lv_werks IS INITIAL.
        " Reconciliation hint only - not a logic-sheet rule
        add( iv_type = 'W' iv_no = '030' iv_v1 = lv_sn iv_v2 = <ls_item>-zsd_partno ).
      ENDIF.

      " Logic B54: ZSD_MATDESC OR material text
      IF <ls_item>-zsd_matdesc IS INITIAL.
        <ls_item>-zsd_matdesc = VALUE #( mt_makt[ matnr = lv_matnr ]-maktx OPTIONAL ).
        IF <ls_item>-zsd_matdesc IS INITIAL.
          add( iv_type = 'W' iv_no = '048' iv_v1 = lv_sn ).
        ENDIF.
      ENDIF.

      " Logic B58: ZSD_HSN OR MARC-STEUC of the Part No. (step 53)
      IF <ls_item>-zsd_hsn IS INITIAL.
        <ls_item>-zsd_hsn = VALUE #( mt_marc[ matnr = lv_matnr werks = lv_werks ]-steuc OPTIONAL ).
        IF <ls_item>-zsd_hsn IS INITIAL.
          add( iv_type = 'W' iv_no = '044' iv_v1 = lv_sn ).
        ENDIF.
      ENDIF.
      <ls_item>-zsd_hsn = zcl_sd_epack_rules=>format_hsn( <ls_item>-zsd_hsn ).

    ENDLOOP.

    " Section heading per HSN, in order of first appearance
    LOOP AT et_items INTO DATA(ls_item).
      IF line_exists( et_headings[ hsn = ls_item-zsd_hsn ] ).
        CONTINUE.
      ENDIF.
      APPEND VALUE #( hsn  = ls_item-zsd_hsn
                      text = zcl_sd_epack_rules=>heading_for_hsn(
                               iv_hsn         = ls_item-zsd_hsn
                               it_goods_lines = it_goods
                               iv_hsn_text    = CONV #( VALUE t604n-text1(
                                                  mt_t604n[ steuc = steuc_of( ls_item-zsd_hsn ) ]-text1 OPTIONAL ) ) ) )
             TO et_headings.
    ENDLOOP.

  ENDMETHOD.


  METHOD order_reference.

    " Logic B13: VBRP-AUBEL -> VBAK-VGBEL (quotation) -> VBKD-BSTKD / BSTDK
    " ("Quotation - Customer reference and date"). Several orders of the
    " packing list: all distinct references, joined with " & " (A17 / Q06).
    DATA lt_bstkd TYPE string_table.
    DATA lt_bstdk TYPE string_table.
    DATA lt_seen  TYPE SORTED TABLE OF vbeln WITH UNIQUE KEY table_line.

    LOOP AT it_vbrp INTO DATA(ls_vbrp) WHERE aubel IS NOT INITIAL.
      INSERT ls_vbrp-aubel INTO TABLE lt_seen.
      IF sy-subrc <> 0.
        CONTINUE.
      ENDIF.

      DATA(lv_vgbel) = VALUE vbeln( mt_vbak[ vbeln = ls_vbrp-aubel ]-vgbel OPTIONAL ).
      IF lv_vgbel IS INITIAL.
        add( iv_type = 'W' iv_no = '061' iv_v1 = |{ ls_vbrp-aubel ALPHA = OUT }| ).
        CONTINUE.
      ENDIF.

      READ TABLE mt_vbkd INTO DATA(ls_vbkd) WITH TABLE KEY vbeln = lv_vgbel.
      IF sy-subrc <> 0.
        CONTINUE.
      ENDIF.

      DATA(lv_bstkd) = condense( CONV string( ls_vbkd-bstkd ) ).
      IF lv_bstkd IS NOT INITIAL AND NOT line_exists( lt_bstkd[ table_line = lv_bstkd ] ).
        APPEND lv_bstkd TO lt_bstkd.
      ENDIF.
      IF ls_vbkd-bstdk IS NOT INITIAL.
        DATA(lv_bstdk) = |{ ls_vbkd-bstdk+6(2) }.{ ls_vbkd-bstdk+4(2) }.{ ls_vbkd-bstdk(4) }|.
        IF NOT line_exists( lt_bstdk[ table_line = lv_bstdk ] ).
          APPEND lv_bstdk TO lt_bstdk.
        ENDIF.
      ENDIF.
    ENDLOOP.

    IF lt_bstkd IS NOT INITIAL OR lt_bstdk IS NOT INITIAL.
      APPEND |{ join( it_parts = lt_bstkd iv_sep = ` & ` ) }{
               COND string( WHEN lt_bstdk IS NOT INITIAL
                            THEN | DATE: { join( it_parts = lt_bstdk iv_sep = ` & ` ) }| ) }|
             TO rt_lines.
    ENDIF.

  ENDMETHOD.


  METHOD exporter_lines.

    " Logic B3-B5: NAME1, STREET, STR_SUPPL1, CITY1, POST_CODE1,
    " REGION (text), COUNTRY (text); B6 GST No. as in the samples
    DATA(ls_adr) = VALUE ty_adrc( mt_adrc[ addrnumber = iv_adrnr ] OPTIONAL ).
    IF ls_adr IS INITIAL.
      add( iv_type = 'W' iv_no = '049' iv_v1 = 'Exporter address' ).
      RETURN.
    ENDIF.

    rt_lines = VALUE #(
      ( CONV #( ls_adr-name1 ) )
      ( CONV #( ls_adr-street ) )
      ( CONV #( ls_adr-str_suppl1 ) )
      ( join( VALUE #( ( condense( |{ ls_adr-city1 }{ COND string( WHEN ls_adr-post_code1 IS NOT INITIAL
                                                                    THEN | - { ls_adr-post_code1 }| ) }| ) )
                       ( to_upper( VALUE t005u-bezei( mt_t005u[ land1 = ls_adr-country bland = ls_adr-region ]-bezei OPTIONAL ) ) )
                       ( country_text( ls_adr-country ) ) ) ) ) ).

    IF iv_gstin IS NOT INITIAL.
      APPEND |GST NO. { iv_gstin }| TO rt_lines.
    ENDIF.

    DELETE rt_lines WHERE table_line IS INITIAL.

  ENDMETHOD.


  METHOD party_lines.

    " Logic B20-B23 / B26-B29: NAME1, STREET, STR_SUPPL1, CITY1, CITY2,
    " POST_CODE1, COUNTRY (text); e-mail; telephone (R3_USER 1) and
    " mobile (R3_USER 3)
    DATA(ls_adr) = VALUE ty_adrc( mt_adrc[ addrnumber = is_partner-adrnr ] OPTIONAL ).
    IF ls_adr IS INITIAL.
      IF is_partner-kunnr IS NOT INITIAL.
        add( iv_type = 'W' iv_no = '049' iv_v1 = |Address of customer { is_partner-kunnr ALPHA = OUT }| ).
      ENDIF.
      RETURN.
    ENDIF.

    APPEND CONV string( ls_adr-name1 ) TO rt_lines.
    APPEND join( VALUE #( ( CONV #( ls_adr-street ) )
                          ( CONV #( ls_adr-str_suppl1 ) )
                          ( CONV #( ls_adr-city2 ) )
                          ( condense( |{ ls_adr-post_code1 } { ls_adr-city1 }| ) )
                          ( country_text( ls_adr-country ) ) ) ) TO rt_lines.

    " Default number first, then the lowest sequence number (A30)
    DATA lv_tel  TYPE string.
    DATA lv_mob  TYPE string.
    DATA lv_mail TYPE string.
    LOOP AT mt_adr2 INTO DATA(ls_adr2) WHERE addrnumber = is_partner-adrnr.
      CASE ls_adr2-r3_user.
        WHEN '1'.
          IF lv_tel IS INITIAL OR ls_adr2-flgdefault = abap_true.
            lv_tel = ls_adr2-telnr_long.
          ENDIF.
        WHEN '3'.
          IF lv_mob IS INITIAL OR ls_adr2-flgdefault = abap_true.
            lv_mob = ls_adr2-telnr_long.
          ENDIF.
      ENDCASE.
    ENDLOOP.
    LOOP AT mt_adr6 INTO DATA(ls_adr6) WHERE addrnumber = is_partner-adrnr.
      IF lv_mail IS INITIAL OR ls_adr6-flgdefault = abap_true.
        lv_mail = ls_adr6-smtp_addr.
      ENDIF.
    ENDLOOP.

    APPEND join( it_parts = VALUE #( ( COND #( WHEN lv_tel IS NOT INITIAL THEN |TEL.: { lv_tel }| ) )
                                     ( COND #( WHEN lv_mob IS NOT INITIAL THEN |MO : { lv_mob }| ) ) )
                 iv_sep   = ` ` ) TO rt_lines.
    IF lv_mail IS NOT INITIAL.
      APPEND |E-MAIL : { lv_mail }| TO rt_lines.
    ENDIF.

    DELETE rt_lines WHERE table_line IS INITIAL.

  ENDMETHOD.


  METHOD country_text.
    IF iv_land1 IS INITIAL.
      RETURN.
    ENDIF.
    DATA(lv_landx) = VALUE t005t-landx( mt_t005t[ land1 = iv_land1 ]-landx OPTIONAL ).
    rv_text = COND #( WHEN lv_landx IS NOT INITIAL THEN to_upper( lv_landx ) ELSE iv_land1 ).
  ENDMETHOD.


  METHOD matnr_of.

    " Part No. = SAP material number (logic B58 "from STEP no 53" = Part No.)
    CALL FUNCTION 'CONVERSION_EXIT_MATN1_INPUT'
      EXPORTING
        input        = iv_partno
      IMPORTING
        output       = rv_matnr
      EXCEPTIONS
        length_error = 1
        OTHERS       = 2.
    IF sy-subrc <> 0.
      CLEAR rv_matnr.
    ENDIF.

  ENDMETHOD.


  METHOD join.

    LOOP AT it_parts INTO DATA(lv_part).
      lv_part = condense( lv_part ).
      IF lv_part IS INITIAL.
        CONTINUE.
      ENDIF.
      rv_text = COND #( WHEN rv_text IS INITIAL THEN lv_part ELSE |{ rv_text }{ iv_sep }{ lv_part }| ).
    ENDLOOP.

  ENDMETHOD.


  METHOD append_block.

    " Lines longer than the text field are split; the form wraps text at
    " the window width anyway.
    DATA lv_line_no TYPE zsd_s_epack_prt_txt-line_no.
    DATA lv_rest    TYPE string.
    CONSTANTS lc_len TYPE i VALUE 255.

    LOOP AT it_lines INTO DATA(lv_line).
      lv_rest = lv_line.
      WHILE lv_rest IS NOT INITIAL.
        lv_line_no += 1.
        IF strlen( lv_rest ) > lc_len.
          APPEND VALUE #( block = iv_block line_no = lv_line_no text = lv_rest(lc_len) ) TO ct_texts.
          lv_rest = lv_rest+lc_len.
        ELSE.
          APPEND VALUE #( block = iv_block line_no = lv_line_no text = lv_rest ) TO ct_texts.
          CLEAR lv_rest.
        ENDIF.
      ENDWHILE.
    ENDLOOP.

  ENDMETHOD.


  METHOD steuc_of.
    rv_steuc = replace( val = condense( CONV string( iv_hsn ) ) sub = ` ` with = `` occ = 0 ).
  ENDMETHOD.

ENDCLASS.
