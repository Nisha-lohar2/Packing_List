"! <p class="shorttext synchronized">Export Packing List - data provider</p>
"!
"! WRICEF 102-B "Export consolidated packing list".
"!
"! Reads the packing list (Z tables) and the SAP data of its tax
"! invoices, validates them and returns everything the Smart Form prints.
"! The forms contain NO SELECT - they only format what this class passes.
"!
"! PERFORMANCE CONTRACT
"!   Every table is read once per run (array SELECT / FOR ALL ENTRIES);
"!   there is no SELECT inside a loop. Volumes are small (samples: 77-90
"!   lines, a handful of invoices per packing list).
"!
"! DEVIATIONS FROM THE FS LOGIC SHEET (Logic_sheet1.xlsx) - deliberate,
"! see docs/Packing_List_Requirement_Analysis_and_Development_Approach.md
"! section 6.4:
"!   B4   region text      T005S-BEZEI  -> T005U-BEZEI
"!   B5   country text     T005-LANDX   -> T005T-LANDX
"!   B54  material text    MARA-MAKTX   -> MAKT-MAKTX
"!   B20/B26 partners      KNVP SP/SH   -> VBPA AG/WE of the invoice
"!                         (the partners actually billed, not the
"!                          customer master default)
"!   B13  sales order      VBRK-AUBEL   -> VBRP-AUBEL (item field)
"!
"! ⚠ SAP-VERIFY: field names of J_1BBRANCH, T001Z, T604N, TVZBT, TINCT,
"!   ADR2/ADR3/ADR6 must be checked in the DEV system before activation.
CLASS zcl_sd_epack_data DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC.

  PUBLIC SECTION.

    "! Packing list for the selection screen input.
    "! Invoices -> link table ZSD_EPACK_INV -> exactly one packing list.
    "! @raising zcx_sd_epack | nothing / several / incomplete selection
    METHODS resolve_packno
      IMPORTING it_vbeln         TYPE zif_sd_epack=>tr_vbeln
                iv_packno        TYPE zsd_packno OPTIONAL
      RETURNING VALUE(rv_packno) TYPE zsd_packno
      RAISING   zcx_sd_epack.

    "! Read, validate and prepare everything for the form.
    "! @parameter iv_format | ZIF_SD_EPACK=>GC_FORMAT-*
    "! @raising zcx_sd_epack | any error - MT_MESSAGES holds the full list
    METHODS get_print_data
      IMPORTING iv_packno      TYPE zsd_packno
                iv_format      TYPE char1 DEFAULT zif_sd_epack=>gc_format-auto
      RETURNING VALUE(rs_data) TYPE zif_sd_epack=>ts_print_data
      RAISING   zcx_sd_epack.

    "! Warnings of the last GET_PRINT_DATA call (do not block output)
    METHODS get_warnings
      RETURNING VALUE(rt_messages) TYPE bapiret2_t.

    "! Record who printed the packing list and when (not for preview)
    CLASS-METHODS mark_printed
      IMPORTING iv_packno TYPE zsd_packno.

  PRIVATE SECTION.

    TYPES:
      BEGIN OF ty_vbrk,
        vbeln TYPE vbrk-vbeln,
        fkart TYPE vbrk-fkart,
        fksto TYPE vbrk-fksto,
        sfakn TYPE vbrk-sfakn,
        bukrs TYPE vbrk-bukrs,
        vkorg TYPE vbrk-vkorg,
        kunag TYPE vbrk-kunag,
        zterm TYPE vbrk-zterm,
        inco1 TYPE vbrk-inco1,
        inco2 TYPE vbrk-inco2,
        bupla TYPE vbrk-bupla,
      END OF ty_vbrk,
      tt_vbrk TYPE SORTED TABLE OF ty_vbrk WITH UNIQUE KEY vbeln,

      BEGIN OF ty_vbrp,
        vbeln TYPE vbrp-vbeln,
        posnr TYPE vbrp-posnr,
        matnr TYPE vbrp-matnr,
        werks TYPE vbrp-werks,
        aubel TYPE vbrp-aubel,
      END OF ty_vbrp,
      tt_vbrp TYPE STANDARD TABLE OF ty_vbrp WITH EMPTY KEY,

      BEGIN OF ty_partner,
        vbeln TYPE vbpa-vbeln,
        posnr TYPE vbpa-posnr,
        parvw TYPE vbpa-parvw,
        kunnr TYPE vbpa-kunnr,
        adrnr TYPE vbpa-adrnr,
      END OF ty_partner,
      tt_partner TYPE STANDARD TABLE OF ty_partner WITH EMPTY KEY,

      BEGIN OF ty_address,
        addrnumber TYPE adrc-addrnumber,
        date_from  TYPE adrc-date_from,
        name1      TYPE adrc-name1,
        name2      TYPE adrc-name2,
        name_co    TYPE adrc-name_co,
        street     TYPE adrc-street,
        str_suppl1 TYPE adrc-str_suppl1,
        str_suppl2 TYPE adrc-str_suppl2,
        city1      TYPE adrc-city1,
        city2      TYPE adrc-city2,
        post_code1 TYPE adrc-post_code1,
        region     TYPE adrc-region,
        country    TYPE adrc-country,
        building   TYPE adrc-building,
      END OF ty_address,
      tt_address TYPE SORTED TABLE OF ty_address WITH UNIQUE KEY addrnumber,

      BEGIN OF ty_comm,
        addrnumber TYPE adr2-addrnumber,
        consnumber TYPE adr2-consnumber,
        flgdefault TYPE adr2-flgdefault,
        kind       TYPE char1,     " T tel / M mobile / F fax / E e-mail
        r3_user    TYPE adr2-r3_user,
        value      TYPE string,
      END OF ty_comm,
      tt_comm TYPE STANDARD TABLE OF ty_comm WITH EMPTY KEY,

      BEGIN OF ty_party,
        kunnr TYPE kunnr,
        adrnr TYPE adrnr,
        stcd1 TYPE stcd1,
      END OF ty_party.

    DATA ms_pl        TYPE zif_sd_epack=>ts_packing_list.
    DATA mt_vbrk      TYPE tt_vbrk.
    DATA mt_vbrp      TYPE tt_vbrp.
    DATA ms_sold_to   TYPE ty_party.
    DATA ms_ship_to   TYPE ty_party.
    DATA mt_address   TYPE tt_address.
    DATA mt_comm      TYPE tt_comm.
    DATA mt_messages  TYPE bapiret2_t.
    DATA mv_company_adrnr TYPE adrnr.
    DATA mv_gstin         TYPE zsd_s_epack_prt_hdr-gstin.

    METHODS add
      IMPORTING iv_type TYPE bapi_mtype DEFAULT 'E'
                iv_no   TYPE symsgno
                iv_v1   TYPE simple OPTIONAL
                iv_v2   TYPE simple OPTIONAL
                iv_v3   TYPE simple OPTIONAL
                iv_v4   TYPE simple OPTIONAL.

    METHODS load_packing_list
      IMPORTING iv_packno TYPE zsd_packno
      RAISING   zcx_sd_epack.

    METHODS read_invoices.

    METHODS check_authority.

    METHODS check_consistency.

    METHODS read_partners.

    METHODS read_addresses.

    METHODS determine_form
      IMPORTING iv_format      TYPE char1
                iv_np_count    TYPE i
      RETURNING VALUE(rv_form) TYPE tdsfname.

    METHODS resolve_items
      EXPORTING et_items    TYPE zif_sd_epack=>tt_item
                et_headings TYPE zif_sd_epack=>tt_hsn_heading.

    METHODS build_header
      IMPORTING is_totals        TYPE zif_sd_epack=>ts_totals
      RETURNING VALUE(rs_header) TYPE zsd_s_epack_prt_hdr.

    METHODS build_texts
      IMPORTING it_headings     TYPE zif_sd_epack=>tt_hsn_heading
      RETURNING VALUE(rt_texts) TYPE zsd_tt_epack_prt_txt.

    METHODS exporter_lines
      RETURNING VALUE(rt_lines) TYPE zif_sd_epack=>tt_text_line.

    METHODS party_lines
      IMPORTING is_party        TYPE ty_party
                iv_with_tax_no  TYPE abap_bool DEFAULT abap_false
      RETURNING VALUE(rt_lines) TYPE zif_sd_epack=>tt_text_line.

    METHODS order_reference
      RETURNING VALUE(rt_lines) TYPE zif_sd_epack=>tt_text_line.

    METHODS payment_lines
      RETURNING VALUE(rt_lines) TYPE zif_sd_epack=>tt_text_line.

    METHODS comm_value
      IMPORTING iv_adrnr        TYPE adrnr
                iv_kind         TYPE char1
                iv_r3_user      TYPE adr2-r3_user OPTIONAL
      RETURNING VALUE(rv_value) TYPE string.

    METHODS country_text
      IMPORTING iv_land1       TYPE land1
      RETURNING VALUE(rv_text) TYPE string.

    METHODS region_text
      IMPORTING iv_land1       TYPE land1
                iv_region      TYPE regio
      RETURNING VALUE(rv_text) TYPE string.

    CLASS-METHODS join
      IMPORTING it_parts       TYPE string_table
                iv_sep         TYPE string DEFAULT `, `
      RETURNING VALUE(rv_text) TYPE string.

    CLASS-METHODS append_block
      IMPORTING iv_block TYPE zsd_epack_block
                it_lines TYPE zif_sd_epack=>tt_text_line
      CHANGING  ct_texts TYPE zsd_tt_epack_prt_txt.

    CLASS-METHODS to_matnr
      IMPORTING iv_partno       TYPE zsd_partno
      RETURNING VALUE(rv_matnr) TYPE matnr.

ENDCLASS.


CLASS zcl_sd_epack_data IMPLEMENTATION.

  METHOD resolve_packno.

    TYPES: BEGIN OF ty_link,
             zsd_packno TYPE zsd_packno,
             vbeln      TYPE vbeln_vf,
           END OF ty_link.
    DATA lt_link TYPE STANDARD TABLE OF ty_link WITH EMPTY KEY.

    CLEAR mt_messages.

    IF iv_packno IS INITIAL AND it_vbeln IS INITIAL.
      zcx_sd_epack=>raise( iv_msgno = '003' ).
    ENDIF.

    IF iv_packno IS NOT INITIAL.
      SELECT SINGLE zsd_packno FROM zsd_epack_hdr
        WHERE zsd_packno = @iv_packno
        INTO @rv_packno.
      IF sy-subrc <> 0.
        zcx_sd_epack=>raise( iv_msgno = '007' iv_v1 = iv_packno ).
      ENDIF.
    ELSE.
      SELECT zsd_packno, vbeln FROM zsd_epack_inv
        WHERE vbeln IN @it_vbeln
        INTO TABLE @lt_link.

      " Invoices entered one by one must all be linked. Invoices that
      " merely fall inside an entered From-To range are ignored when they
      " are not linked (ASSUMPTION A03).
      LOOP AT it_vbeln INTO DATA(ls_sel) WHERE sign = 'I' AND option = 'EQ'.
        IF NOT line_exists( lt_link[ vbeln = ls_sel-low ] ).
          add( iv_no = '004' iv_v1 = |{ ls_sel-low ALPHA = OUT }| ).
        ENDIF.
      ENDLOOP.

      DATA(lt_pl) = lt_link.
      SORT lt_pl BY zsd_packno.
      DELETE ADJACENT DUPLICATES FROM lt_pl COMPARING zsd_packno.

      CASE lines( lt_pl ).
        WHEN 0.
          IF mt_messages IS INITIAL.
            add( iv_no = '046' ).
          ENDIF.
        WHEN 1.
          rv_packno = lt_pl[ 1 ]-zsd_packno.
        WHEN OTHERS.
          add( iv_no = '005' iv_v1 = lt_pl[ 1 ]-zsd_packno iv_v2 = lt_pl[ 2 ]-zsd_packno ).
      ENDCASE.

      IF mt_messages IS NOT INITIAL.
        zcx_sd_epack=>raise_from_messages( mt_messages ).
      ENDIF.
    ENDIF.

    " The packing list is a unit: when invoices were entered, every
    " invoice of the packing list must be part of the selection
    " (ASSUMPTION A02).
    IF it_vbeln IS NOT INITIAL.
      SELECT zsd_packno, vbeln FROM zsd_epack_inv
        WHERE zsd_packno = @rv_packno
        INTO TABLE @lt_link.
      LOOP AT lt_link INTO DATA(ls_link).
        IF ls_link-vbeln NOT IN it_vbeln.
          add( iv_no = '006' iv_v1 = rv_packno iv_v2 = |{ ls_link-vbeln ALPHA = OUT }| ).
        ENDIF.
      ENDLOOP.
      LOOP AT it_vbeln INTO ls_sel WHERE sign = 'I' AND option = 'EQ'.
        IF NOT line_exists( lt_link[ vbeln = ls_sel-low ] ).
          add( iv_no = '045' iv_v1 = |{ ls_sel-low ALPHA = OUT }| iv_v2 = rv_packno ).
        ENDIF.
      ENDLOOP.
      IF mt_messages IS NOT INITIAL.
        zcx_sd_epack=>raise_from_messages( mt_messages ).
      ENDIF.
    ENDIF.

  ENDMETHOD.


  METHOD get_print_data.

    CLEAR: mt_messages, ms_pl, mt_vbrk, mt_vbrp, ms_sold_to, ms_ship_to,
           mt_address, mt_comm, mv_company_adrnr, mv_gstin.

    load_packing_list( iv_packno ).
    read_invoices( ).
    check_authority( ).
    check_consistency( ).
    read_partners( ).

    " FS 2.1 - checked again at output time although the upload already
    " rejects it: the table may have been changed by other means.
    APPEND LINES OF zcl_sd_epack_rules=>check_weights( ms_pl-items ) TO mt_messages.

    DATA(lv_np_count) = 0.
    DO zif_sd_epack=>gc_max_np TIMES.
      ASSIGN COMPONENT |ZSD_NP{ sy-index }| OF STRUCTURE ms_pl-header TO FIELD-SYMBOL(<lv_np>).
      IF sy-subrc = 0 AND <lv_np> IS NOT INITIAL.
        lv_np_count += 1.
      ENDIF.
    ENDDO.
    rs_data-form = determine_form( iv_format   = iv_format
                                   iv_np_count = lv_np_count ).

    IF zcl_sd_epack_rules=>has_errors( mt_messages ).
      zcx_sd_epack=>raise_from_messages( mt_messages ).
    ENDIF.

    read_addresses( ).

    resolve_items( IMPORTING et_items    = DATA(lt_items)
                             et_headings = DATA(lt_headings) ).

    DATA(ls_totals) = zcl_sd_epack_rules=>calc_totals( lt_items ).

    rs_data-header = build_header( ls_totals ).
    rs_data-header-np_count = lv_np_count.
    rs_data-texts  = build_texts( lt_headings ).
    rs_data-items  = zcl_sd_epack_rules=>build_print_items( it_items    = lt_items
                                                            it_headings = lt_headings ).

    IF zcl_sd_epack_rules=>has_errors( mt_messages ).
      zcx_sd_epack=>raise_from_messages( mt_messages ).
    ENDIF.

  ENDMETHOD.


  METHOD get_warnings.
    rt_messages = mt_messages.
    DELETE rt_messages WHERE type CA 'EAX'.
  ENDMETHOD.


  METHOD mark_printed.

    UPDATE zsd_epack_hdr
      SET zsd_prnam = @sy-uname,
          zsd_prdat = @sy-datum,
          zsd_przet = @sy-uzeit
      WHERE zsd_packno = @iv_packno.
    IF sy-subrc = 0.
      COMMIT WORK.
    ENDIF.

  ENDMETHOD.


  METHOD add.
    APPEND zcl_sd_epack_rules=>msg( iv_type = iv_type iv_no = iv_no
                                    iv_v1 = iv_v1 iv_v2 = iv_v2
                                    iv_v3 = iv_v3 iv_v4 = iv_v4 ) TO mt_messages.
  ENDMETHOD.


  METHOD load_packing_list.

    SELECT SINGLE * FROM zsd_epack_hdr
      WHERE zsd_packno = @iv_packno
      INTO @ms_pl-header.
    IF sy-subrc <> 0.
      zcx_sd_epack=>raise( iv_msgno = '007' iv_v1 = iv_packno ).
    ENDIF.

    SELECT vbeln FROM zsd_epack_inv
      WHERE zsd_packno = @iv_packno
      INTO TABLE @ms_pl-invoices.

    SELECT * FROM zsd_epack_data
      WHERE zsd_packno = @iv_packno
      ORDER BY zsd_sn
      INTO TABLE @ms_pl-items.

    IF ms_pl-invoices IS INITIAL OR ms_pl-items IS INITIAL.
      zcx_sd_epack=>raise( iv_msgno = '047' iv_v1 = iv_packno ).
    ENDIF.

  ENDMETHOD.


  METHOD read_invoices.

    DATA lr_fkart TYPE RANGE OF fkart.

    SELECT vbeln, fkart, fksto, sfakn, bukrs, vkorg, kunag, zterm,
           inco1, inco2, bupla
      FROM vbrk
      FOR ALL ENTRIES IN @ms_pl-invoices
      WHERE vbeln = @ms_pl-invoices-table_line
      INTO TABLE @mt_vbrk.

    SELECT vbeln, posnr, matnr, werks, aubel
      FROM vbrp
      FOR ALL ENTRIES IN @ms_pl-invoices
      WHERE vbeln = @ms_pl-invoices-table_line
      INTO TABLE @mt_vbrp.

    " Allowed billing types (ASSUMPTION A11)
    SELECT sign, opti AS option, low, high
      FROM tvarvc
      WHERE name = @zif_sd_epack=>gc_tvarv_fkart
        AND type = 'S'
      INTO CORRESPONDING FIELDS OF TABLE @lr_fkart.
    DELETE lr_fkart WHERE low IS INITIAL AND high IS INITIAL.

    LOOP AT ms_pl-invoices INTO DATA(lv_vbeln).
      DATA(lv_ext) = |{ lv_vbeln ALPHA = OUT }|.
      READ TABLE mt_vbrk INTO DATA(ls_vbrk) WITH TABLE KEY vbeln = lv_vbeln.
      IF sy-subrc <> 0.
        add( iv_no = '008' iv_v1 = lv_ext ).
        CONTINUE.
      ENDIF.
      IF ls_vbrk-fksto = abap_true OR ls_vbrk-sfakn IS NOT INITIAL.
        add( iv_no = '009' iv_v1 = lv_ext ).
      ENDIF.
      IF lr_fkart IS NOT INITIAL AND ls_vbrk-fkart NOT IN lr_fkart.
        add( iv_no = '010' iv_v1 = lv_ext iv_v2 = ls_vbrk-fkart ).
      ENDIF.
    ENDLOOP.

  ENDMETHOD.


  METHOD check_authority.

    " FS 2.3 "Security and Authorization" is empty - ASSUMPTION A24.
    DATA(lt_vbrk) = mt_vbrk.
    SORT lt_vbrk BY bukrs.
    DELETE ADJACENT DUPLICATES FROM lt_vbrk COMPARING bukrs.
    LOOP AT lt_vbrk INTO DATA(ls_vbrk).
      AUTHORITY-CHECK OBJECT zif_sd_epack=>gc_auth_object
        ID 'ACTVT' FIELD zif_sd_epack=>gc_actvt-print
        ID 'BUKRS' FIELD ls_vbrk-bukrs.
      IF sy-subrc <> 0.
        add( iv_no = '012' iv_v1 = 'Company code' iv_v2 = ls_vbrk-bukrs iv_v3 = zif_sd_epack=>gc_actvt-print ).
      ENDIF.
    ENDLOOP.

    lt_vbrk = mt_vbrk.
    SORT lt_vbrk BY vkorg.
    DELETE ADJACENT DUPLICATES FROM lt_vbrk COMPARING vkorg.
    LOOP AT lt_vbrk INTO ls_vbrk.
      AUTHORITY-CHECK OBJECT 'V_VBRK_VKO'
        ID 'VKORG' FIELD ls_vbrk-vkorg
        ID 'ACTVT' FIELD zif_sd_epack=>gc_actvt-print.
      IF sy-subrc <> 0.
        add( iv_no = '012' iv_v1 = 'Sales org.' iv_v2 = ls_vbrk-vkorg iv_v3 = zif_sd_epack=>gc_actvt-print ).
      ENDIF.
    ENDLOOP.

  ENDMETHOD.


  METHOD check_consistency.

    " One packing list prints ONE exporter, buyer, GSTIN, payment term and
    " Incoterm - all invoices must agree (ASSUMPTION A04).
    IF lines( mt_vbrk ) < 2.
      RETURN.
    ENDIF.

    DATA(ls_first) = mt_vbrk[ 1 ].
    LOOP AT mt_vbrk INTO DATA(ls_vbrk) FROM 2.
      IF ls_vbrk-bukrs <> ls_first-bukrs.
        add( iv_no = '011' iv_v1 = 'Company code' iv_v2 = ls_first-bukrs iv_v3 = ls_vbrk-bukrs ).
      ENDIF.
      IF ls_vbrk-kunag <> ls_first-kunag.
        add( iv_no = '011' iv_v1 = 'Sold-to party' iv_v2 = |{ ls_first-kunag ALPHA = OUT }| iv_v3 = |{ ls_vbrk-kunag ALPHA = OUT }| ).
      ENDIF.
      IF ls_vbrk-bupla <> ls_first-bupla.
        add( iv_no = '011' iv_v1 = 'Business place (GSTIN)' iv_v2 = ls_first-bupla iv_v3 = ls_vbrk-bupla ).
      ENDIF.
      IF ls_vbrk-zterm <> ls_first-zterm.
        add( iv_no = '011' iv_v1 = 'Payment terms' iv_v2 = ls_first-zterm iv_v3 = ls_vbrk-zterm ).
      ENDIF.
      IF ls_vbrk-inco1 <> ls_first-inco1 OR ls_vbrk-inco2 <> ls_first-inco2.
        add( iv_no = '011' iv_v1 = 'Incoterms' iv_v2 = |{ ls_first-inco1 } { ls_first-inco2 }|
                                               iv_v3 = |{ ls_vbrk-inco1 } { ls_vbrk-inco2 }| ).
      ENDIF.
    ENDLOOP.

  ENDMETHOD.


  METHOD read_partners.

    DATA lt_partner TYPE tt_partner.
    DATA lt_ship_to TYPE tt_partner.

    IF mt_vbrk IS INITIAL.
      RETURN.
    ENDIF.

    SELECT vbeln, posnr, parvw, kunnr, adrnr
      FROM vbpa
      FOR ALL ENTRIES IN @mt_vbrk
      WHERE vbeln = @mt_vbrk-vbeln
        AND ( parvw = @zif_sd_epack=>gc_parvw-sold_to OR parvw = @zif_sd_epack=>gc_parvw-ship_to )
      INTO TABLE @lt_partner.

    " Header partner (POSNR 000000) first, otherwise the first item partner
    SORT lt_partner BY vbeln parvw posnr.

    LOOP AT mt_vbrk INTO DATA(ls_vbrk).
      READ TABLE lt_partner INTO DATA(ls_we)
           WITH KEY vbeln = ls_vbrk-vbeln parvw = zif_sd_epack=>gc_parvw-ship_to.
      IF sy-subrc <> 0.
        add( iv_no = '042' iv_v1 = |{ ls_vbrk-vbeln ALPHA = OUT }| ).
        CONTINUE.
      ENDIF.
      APPEND ls_we TO lt_ship_to.
    ENDLOOP.

    " Ship-to must be identical as well (ASSUMPTION A04)
    DATA(lt_distinct) = lt_ship_to.
    SORT lt_distinct BY kunnr.
    DELETE ADJACENT DUPLICATES FROM lt_distinct COMPARING kunnr.
    IF lines( lt_distinct ) > 1.
      add( iv_no = '011' iv_v1 = 'Ship-to party'
                         iv_v2 = |{ lt_distinct[ 1 ]-kunnr ALPHA = OUT }|
                         iv_v3 = |{ lt_distinct[ 2 ]-kunnr ALPHA = OUT }| ).
    ENDIF.

    DATA(ls_first) = mt_vbrk[ 1 ].
    READ TABLE lt_partner INTO DATA(ls_ag)
         WITH KEY vbeln = ls_first-vbeln parvw = zif_sd_epack=>gc_parvw-sold_to.
    ms_sold_to = VALUE #( kunnr = COND #( WHEN sy-subrc = 0 THEN ls_ag-kunnr ELSE ls_first-kunag )
                          adrnr = COND #( WHEN sy-subrc = 0 THEN ls_ag-adrnr ) ).
    IF lt_ship_to IS NOT INITIAL.
      ms_ship_to = VALUE #( kunnr = lt_ship_to[ 1 ]-kunnr
                            adrnr = lt_ship_to[ 1 ]-adrnr ).
    ENDIF.

    " Customer master: address number fallback + tax number (ASSUMPTION A16)
    SELECT kunnr, adrnr, stcd1 FROM kna1
      WHERE kunnr = @ms_sold_to-kunnr OR kunnr = @ms_ship_to-kunnr
      INTO TABLE @DATA(lt_kna1).

    LOOP AT lt_kna1 INTO DATA(ls_kna1).
      IF ls_kna1-kunnr = ms_sold_to-kunnr.
        ms_sold_to-stcd1 = ls_kna1-stcd1.
        IF ms_sold_to-adrnr IS INITIAL.
          ms_sold_to-adrnr = ls_kna1-adrnr.
        ENDIF.
      ENDIF.
      IF ls_kna1-kunnr = ms_ship_to-kunnr.
        ms_ship_to-stcd1 = ls_kna1-stcd1.
        IF ms_ship_to-adrnr IS INITIAL.
          ms_ship_to-adrnr = ls_kna1-adrnr.
        ENDIF.
      ENDIF.
    ENDLOOP.

  ENDMETHOD.


  METHOD read_addresses.

    DATA lt_adrnr TYPE SORTED TABLE OF adrnr WITH UNIQUE KEY table_line.
    DATA lt_adrc  TYPE STANDARD TABLE OF ty_address WITH EMPTY KEY.

    DATA(ls_vbrk) = mt_vbrk[ 1 ].

    SELECT SINGLE adrnr FROM t001
      WHERE bukrs = @ls_vbrk-bukrs
      INTO @mv_company_adrnr.

    INSERT mv_company_adrnr  INTO TABLE lt_adrnr.
    INSERT ms_sold_to-adrnr  INTO TABLE lt_adrnr.
    INSERT ms_ship_to-adrnr  INTO TABLE lt_adrnr.
    DELETE lt_adrnr WHERE table_line IS INITIAL.

    IF lt_adrnr IS INITIAL.
      RETURN.
    ENDIF.

    SELECT addrnumber, date_from, name1, name2, name_co, street, str_suppl1,
           str_suppl2, city1, city2, post_code1, region, country, building
      FROM adrc
      FOR ALL ENTRIES IN @lt_adrnr
      WHERE addrnumber = @lt_adrnr-table_line
        AND nation     = @space
        AND date_from <= @sy-datum
      INTO TABLE @lt_adrc.

    " Latest valid version per address
    SORT lt_adrc BY addrnumber date_from DESCENDING.
    DELETE ADJACENT DUPLICATES FROM lt_adrc COMPARING addrnumber.
    mt_address = lt_adrc.

    " Communication data of the address itself (not of contact persons)
    SELECT addrnumber, consnumber, flgdefault, r3_user, tel_number, telnr_long
      FROM adr2
      FOR ALL ENTRIES IN @lt_adrnr
      WHERE addrnumber = @lt_adrnr-table_line
        AND persnumber = @space
      INTO TABLE @DATA(lt_adr2).

    LOOP AT lt_adr2 INTO DATA(ls_adr2).
      APPEND VALUE #( addrnumber = ls_adr2-addrnumber
                      consnumber = ls_adr2-consnumber
                      flgdefault = ls_adr2-flgdefault
                      r3_user    = ls_adr2-r3_user
                      kind       = COND #( WHEN ls_adr2-r3_user CA '23' THEN 'M' ELSE 'T' )
                      value      = COND #( WHEN ls_adr2-telnr_long IS NOT INITIAL
                                           THEN ls_adr2-telnr_long ELSE ls_adr2-tel_number ) )
             TO mt_comm.
    ENDLOOP.

    SELECT addrnumber, consnumber, flgdefault, fax_number, faxnr_long
      FROM adr3
      FOR ALL ENTRIES IN @lt_adrnr
      WHERE addrnumber = @lt_adrnr-table_line
        AND persnumber = @space
      INTO TABLE @DATA(lt_adr3).

    LOOP AT lt_adr3 INTO DATA(ls_adr3).
      APPEND VALUE #( addrnumber = ls_adr3-addrnumber
                      consnumber = ls_adr3-consnumber
                      flgdefault = ls_adr3-flgdefault
                      kind       = 'F'
                      value      = COND #( WHEN ls_adr3-faxnr_long IS NOT INITIAL
                                           THEN ls_adr3-faxnr_long ELSE ls_adr3-fax_number ) )
             TO mt_comm.
    ENDLOOP.

    SELECT addrnumber, consnumber, flgdefault, smtp_addr
      FROM adr6
      FOR ALL ENTRIES IN @lt_adrnr
      WHERE addrnumber = @lt_adrnr-table_line
        AND persnumber = @space
      INTO TABLE @DATA(lt_adr6).

    LOOP AT lt_adr6 INTO DATA(ls_adr6).
      APPEND VALUE #( addrnumber = ls_adr6-addrnumber
                      consnumber = ls_adr6-consnumber
                      flgdefault = ls_adr6-flgdefault
                      kind       = 'E'
                      value      = ls_adr6-smtp_addr )
             TO mt_comm.
    ENDLOOP.

    DELETE mt_comm WHERE value IS INITIAL.
    " Default entry first, then by sequence number
    SORT mt_comm BY addrnumber kind flgdefault DESCENDING consnumber.

  ENDMETHOD.


  METHOD determine_form.

    " FS: "If notifier party details are maintained in the Z table, then
    " user have to select ... Packing list with notifier party"
    " (ASSUMPTION A05: chosen automatically unless the user overrides).
    CASE iv_format.
      WHEN zif_sd_epack=>gc_format-with_np.
        rv_form = zif_sd_epack=>gc_form-with_np.
      WHEN zif_sd_epack=>gc_format-without_np.
        IF iv_np_count > 0.
          add( iv_no = '013' iv_v1 = ms_pl-header-zsd_packno ).
        ENDIF.
        rv_form = zif_sd_epack=>gc_form-without_np.
      WHEN OTHERS.
        rv_form = COND #( WHEN iv_np_count > 0
                          THEN zif_sd_epack=>gc_form-with_np
                          ELSE zif_sd_epack=>gc_form-without_np ).
    ENDCASE.

  ENDMETHOD.


  METHOD resolve_items.

    " FS #27 / logic B54, B58: description and HSN "from Z table OR
    " material master". Z value wins when filled (ASSUMPTION A12).
    TYPES: BEGIN OF ty_mat,
             matnr TYPE matnr,
           END OF ty_mat.
    DATA lt_mat  TYPE SORTED TABLE OF ty_mat WITH UNIQUE KEY matnr.
    DATA lt_steuc TYPE SORTED TABLE OF steuc WITH UNIQUE KEY table_line.
    DATA lt_goods TYPE zif_sd_epack=>tt_text_line.

    et_items = ms_pl-items.

    LOOP AT et_items INTO DATA(ls_item).
      INSERT VALUE #( matnr = to_matnr( ls_item-zsd_partno ) ) INTO TABLE lt_mat.
    ENDLOOP.
    DELETE lt_mat WHERE matnr IS INITIAL.

    IF lt_mat IS NOT INITIAL.
      SELECT matnr, maktx FROM makt
        FOR ALL ENTRIES IN @lt_mat
        WHERE matnr = @lt_mat-matnr
          AND spras = @zif_sd_epack=>gc_langu
        INTO TABLE @DATA(lt_makt).
      SORT lt_makt BY matnr.

      SELECT matnr, werks, steuc FROM marc
        FOR ALL ENTRIES IN @lt_mat
        WHERE matnr = @lt_mat-matnr
        INTO TABLE @DATA(lt_marc).
      SORT lt_marc BY matnr werks.
    ENDIF.

    LOOP AT et_items ASSIGNING FIELD-SYMBOL(<ls_item>).
      DATA(lv_matnr) = to_matnr( <ls_item>-zsd_partno ).

      IF lv_matnr IS NOT INITIAL AND NOT line_exists( mt_vbrp[ matnr = lv_matnr ] ).
        add( iv_type = 'W' iv_no = '030' iv_v1 = |{ CONV i( <ls_item>-zsd_sn ) }| iv_v2 = <ls_item>-zsd_partno ).
      ENDIF.

      IF <ls_item>-zsd_matdesc IS INITIAL.
        READ TABLE lt_makt INTO DATA(ls_makt) WITH KEY matnr = lv_matnr BINARY SEARCH.
        IF sy-subrc = 0.
          <ls_item>-zsd_matdesc = ls_makt-maktx.
        ELSE.
          add( iv_type = 'W' iv_no = '048' iv_v1 = |{ CONV i( <ls_item>-zsd_sn ) }| ).
        ENDIF.
      ENDIF.

      IF <ls_item>-zsd_hsn IS INITIAL.
        " HSN of the billing plant (FS logic B42: MARC-STEUC)
        DATA(lv_werks) = VALUE werks_d( mt_vbrp[ matnr = lv_matnr ]-werks OPTIONAL ).
        READ TABLE lt_marc INTO DATA(ls_marc) WITH KEY matnr = lv_matnr werks = lv_werks BINARY SEARCH.
        IF sy-subrc = 0 AND ls_marc-steuc IS NOT INITIAL.
          <ls_item>-zsd_hsn = ls_marc-steuc.
        ELSE.
          add( iv_type = 'W' iv_no = '044' iv_v1 = |{ CONV i( <ls_item>-zsd_sn ) }| ).
        ENDIF.
      ENDIF.

      <ls_item>-zsd_hsn = zcl_sd_epack_rules=>format_hsn( <ls_item>-zsd_hsn ).
      IF <ls_item>-zsd_hsn IS NOT INITIAL.
        INSERT CONV steuc( replace( val = <ls_item>-zsd_hsn sub = ` ` with = `` occ = 0 ) )
               INTO TABLE lt_steuc.
      ENDIF.
    ENDLOOP.

    " HSN texts (logic B42: T604N with SPRAS = EN, LAND1 = IN)
    IF lt_steuc IS NOT INITIAL.
      SELECT steuc, text1 FROM t604n
        FOR ALL ENTRIES IN @lt_steuc
        WHERE spras = @zif_sd_epack=>gc_langu
          AND land1 = @zif_sd_epack=>gc_hsn_country
          AND steuc = @lt_steuc-table_line
        INTO TABLE @DATA(lt_t604n).
    ENDIF.

    lt_goods = zcl_sd_epack_rules=>split_lines( ms_pl-header-zsd_hsn_desc ).

    " Section headings in order of first appearance; "club common HSN
    " code and description in all billing documents" (logic B42)
    LOOP AT et_items INTO ls_item.
      IF line_exists( et_headings[ hsn = ls_item-zsd_hsn ] ).
        CONTINUE.
      ENDIF.
      DATA(lv_steuc) = CONV steuc( replace( val = ls_item-zsd_hsn sub = ` ` with = `` occ = 0 ) ).
      APPEND VALUE #( hsn  = ls_item-zsd_hsn
                      text = zcl_sd_epack_rules=>heading_for_hsn(
                               iv_hsn         = ls_item-zsd_hsn
                               it_goods_lines = lt_goods
                               iv_hsn_text    = CONV #( VALUE t604n-text1( lt_t604n[ steuc = lv_steuc ]-text1 OPTIONAL ) ) ) )
             TO et_headings.
    ENDLOOP.

  ENDMETHOD.


  METHOD build_header.

    DATA(ls_hdr)  = ms_pl-header.
    DATA(ls_vbrk) = mt_vbrk[ 1 ].

    rs_header-packno     = ls_hdr-zsd_packno.
    rs_header-packdt_txt = zcl_sd_epack_rules=>format_date( ls_hdr-zsd_packdt ).
    rs_header-revno      = ls_hdr-zsd_revno.

    " Exporter's Ref. / IEC: Excel value if given, else company address
    " field BUILDING (logic B11) - ASSUMPTION A13
    rs_header-iec = COND #( WHEN ls_hdr-zsd_expref IS NOT INITIAL
                            THEN ls_hdr-zsd_expref
                            ELSE VALUE #( mt_address[ addrnumber = mv_company_adrnr ]-building OPTIONAL ) ).
    IF rs_header-iec IS INITIAL.
      add( iv_type = 'W' iv_no = '049' iv_v1 = 'IEC (Exporter''s Ref.)' ).
    ELSEIF rs_header-iec NS 'IEC'.
      rs_header-iec = |IEC:{ rs_header-iec }|.
    ENDIF.

    " GST No. (logic B6) - business place of the invoices
    SELECT SINGLE gstin FROM j_1bbranch
      WHERE bukrs  = @ls_vbrk-bukrs
        AND branch = @ls_vbrk-bupla
      INTO @rs_header-gstin.
    mv_gstin = rs_header-gstin.
    IF rs_header-gstin IS INITIAL.
      add( iv_type = 'W' iv_no = '049' iv_v1 = 'GST No.' ).
    ENDIF.

    " CIN (logic B7)
    SELECT SINGLE paval FROM t001z
      WHERE bukrs = @ls_vbrk-bukrs
        AND party = @zif_sd_epack=>gc_cin_party
      INTO @rs_header-cin.
    IF rs_header-cin IS INITIAL.
      add( iv_type = 'W' iv_no = '049' iv_v1 = 'CIN No.' ).
    ENDIF.

    rs_header-cnty_orgn = ls_hdr-zsd_cnty_orgn.
    rs_header-cnty_dest = country_text(
                            VALUE #( mt_address[ addrnumber = ms_ship_to-adrnr ]-country OPTIONAL ) ).
    rs_header-same_party = xsdbool( ms_sold_to-kunnr = ms_ship_to-kunnr ).

    rs_header-prec    = ls_hdr-zsd_prec.
    rs_header-rec_pc  = ls_hdr-zsd_rec_pc.
    rs_header-vsl_flt = ls_hdr-zsd_vsl_flt.
    rs_header-pol     = ls_hdr-zsd_pol.
    rs_header-pod     = ls_hdr-zsd_pod.
    rs_header-pld     = ls_hdr-zsd_pld.

    rs_header-tot_art_txt = |{ is_totals-articles }|.
    rs_header-tot_qty_txt = zcl_sd_epack_rules=>format_quantity( is_totals-quantity ).
    rs_header-tot_gwt_txt = zcl_sd_epack_rules=>format_weight( is_totals-gross ).
    rs_header-tot_nwt_txt = zcl_sd_epack_rules=>format_weight( is_totals-net ).
    rs_header-tot_text    = |({ zcl_sd_epack_rules=>total_text( is_totals ) })|.

  ENDMETHOD.


  METHOD build_texts.

    append_block( EXPORTING iv_block = zif_sd_epack=>gc_block-exporter
                            it_lines = exporter_lines( )
                  CHANGING  ct_texts = rt_texts ).

    append_block( EXPORTING iv_block = zif_sd_epack=>gc_block-sold_to
                            it_lines = party_lines( ms_sold_to )
                  CHANGING  ct_texts = rt_texts ).

    append_block( EXPORTING iv_block = zif_sd_epack=>gc_block-ship_to
                            it_lines = party_lines( is_party       = ms_ship_to
                                                    iv_with_tax_no = abap_true )
                  CHANGING  ct_texts = rt_texts ).

    DATA(lv_np) = 0.
    DO zif_sd_epack=>gc_max_np TIMES.
      lv_np = sy-index.
      ASSIGN COMPONENT |ZSD_NP{ lv_np }| OF STRUCTURE ms_pl-header TO FIELD-SYMBOL(<lv_np>).
      IF sy-subrc = 0.
        append_block( EXPORTING iv_block = CONV #( |NP{ lv_np }| )
                                it_lines = zcl_sd_epack_rules=>split_lines( <lv_np> )
                      CHANGING  ct_texts = rt_texts ).
      ENDIF.
    ENDDO.

    append_block( EXPORTING iv_block = zif_sd_epack=>gc_block-order_ref
                            it_lines = order_reference( )
                  CHANGING  ct_texts = rt_texts ).

    DATA(lv_buyer_ref) = condense( CONV string( ms_pl-header-zsd_party_ref ) ).
    IF ms_pl-header-zsd_ref_dt IS NOT INITIAL.
      lv_buyer_ref = |{ lv_buyer_ref } DATE: { zcl_sd_epack_rules=>format_date( ms_pl-header-zsd_ref_dt ) }|.
    ENDIF.
    append_block( EXPORTING iv_block = zif_sd_epack=>gc_block-buyer_ref
                            it_lines = zcl_sd_epack_rules=>split_lines( lv_buyer_ref )
                  CHANGING  ct_texts = rt_texts ).

    append_block( EXPORTING iv_block = zif_sd_epack=>gc_block-payment
                            it_lines = payment_lines( )
                  CHANGING  ct_texts = rt_texts ).

    append_block( EXPORTING iv_block = zif_sd_epack=>gc_block-marks
                            it_lines = zcl_sd_epack_rules=>split_lines( ms_pl-header-zsd_mark )
                  CHANGING  ct_texts = rt_texts ).

    append_block( EXPORTING iv_block = zif_sd_epack=>gc_block-container
                            it_lines = zcl_sd_epack_rules=>split_lines( ms_pl-header-zsd_cont )
                  CHANGING  ct_texts = rt_texts ).

    append_block( EXPORTING iv_block = zif_sd_epack=>gc_block-packages
                            it_lines = zcl_sd_epack_rules=>split_lines( ms_pl-header-zsd_nopack )
                  CHANGING  ct_texts = rt_texts ).

    " Description of Goods: uploaded text, otherwise the clubbed HSN
    " headings (ASSUMPTION A14)
    DATA(lt_goods) = zcl_sd_epack_rules=>split_lines( ms_pl-header-zsd_hsn_desc ).
    IF lt_goods IS INITIAL.
      lt_goods = VALUE #( FOR ls_heading IN it_headings ( ls_heading-text ) ).
    ENDIF.
    append_block( EXPORTING iv_block = zif_sd_epack=>gc_block-goods
                            it_lines = lt_goods
                  CHANGING  ct_texts = rt_texts ).

    append_block( EXPORTING iv_block = zif_sd_epack=>gc_block-declaration
                            it_lines = zcl_sd_epack_rules=>split_lines( ms_pl-header-zsd_pl_decl )
                  CHANGING  ct_texts = rt_texts ).

  ENDMETHOD.


  METHOD exporter_lines.

    " Layout of the samples:
    "   ASTRAL LIMITED.
    "   207/1, ASTRAL HOUSE, B/H RAJPATH CLUB,
    "   OFF. S. G. HIGHWAY,
    "   AHMEDABAD - 380 059, GUJARAT, INDIA.
    "   GST NO. 24AABCA2951N1ZO
    DATA(ls_adr) = VALUE ty_address( mt_address[ addrnumber = mv_company_adrnr ] OPTIONAL ).
    IF ls_adr IS INITIAL.
      add( iv_type = 'W' iv_no = '049' iv_v1 = 'Exporter address' ).
      RETURN.
    ENDIF.

    APPEND condense( |{ ls_adr-name1 } { ls_adr-name2 }| ) TO rt_lines.
    APPEND CONV string( ls_adr-street )     TO rt_lines.
    APPEND CONV string( ls_adr-str_suppl1 ) TO rt_lines.
    APPEND CONV string( ls_adr-str_suppl2 ) TO rt_lines.
    APPEND join( VALUE #( ( condense( |{ ls_adr-city1 }{ COND string( WHEN ls_adr-post_code1 IS NOT INITIAL
                                                                         THEN | - { ls_adr-post_code1 }| ) }| ) )
                          ( region_text( iv_land1 = ls_adr-country iv_region = ls_adr-region ) )
                          ( country_text( ls_adr-country ) ) ) ) TO rt_lines.

    " GSTIN was read in BUILD_HEADER (called before BUILD_TEXTS)
    IF mv_gstin IS NOT INITIAL.
      APPEND |GST NO. { mv_gstin }| TO rt_lines.
    ENDIF.

    DELETE rt_lines WHERE table_line IS INITIAL.

  ENDMETHOD.


  METHOD party_lines.

    " Layout of sample 90000081:
    "   NAME
    "   STREET, SUPPLEMENT, DISTRICT, POSTCODE CITY, REGION, COUNTRY
    "   K/A:- CONTACT                        (ADRC-NAME_CO, ASSUMPTION A15)
    "   TEL.: ... FAX : ... MO : ...
    "   E-MAIL : ...
    "   TAX NUMBER : ...                     (ship-to only, ASSUMPTION A16)
    DATA(ls_adr) = VALUE ty_address( mt_address[ addrnumber = is_party-adrnr ] OPTIONAL ).
    IF ls_adr IS INITIAL.
      add( iv_type = 'W' iv_no = '049' iv_v1 = |Address of customer { is_party-kunnr ALPHA = OUT }| ).
      RETURN.
    ENDIF.

    APPEND condense( |{ ls_adr-name1 } { ls_adr-name2 }| ) TO rt_lines.
    APPEND join( VALUE #( ( CONV #( ls_adr-street ) )
                          ( CONV #( ls_adr-str_suppl1 ) )
                          ( CONV #( ls_adr-str_suppl2 ) )
                          ( CONV #( ls_adr-city2 ) )
                          ( condense( |{ ls_adr-post_code1 } { ls_adr-city1 }| ) )
                          ( region_text( iv_land1 = ls_adr-country iv_region = ls_adr-region ) )
                          ( country_text( ls_adr-country ) ) ) ) TO rt_lines.

    IF ls_adr-name_co IS NOT INITIAL.
      APPEND |K/A:- { ls_adr-name_co }| TO rt_lines.
    ENDIF.

    " Logic B22/B23: R3_USER = 1 telephone, R3_USER = 3 mobile
    DATA(lv_tel) = comm_value( iv_adrnr = is_party-adrnr iv_kind = 'T' iv_r3_user = '1' ).
    DATA(lv_mob) = comm_value( iv_adrnr = is_party-adrnr iv_kind = 'M' iv_r3_user = '3' ).
    DATA(lv_fax) = comm_value( iv_adrnr = is_party-adrnr iv_kind = 'F' ).
    APPEND join( it_parts = VALUE #( ( COND #( WHEN lv_tel IS NOT INITIAL THEN |TEL.: { lv_tel }| ) )
                                     ( COND #( WHEN lv_fax IS NOT INITIAL THEN |FAX : { lv_fax }| ) )
                                     ( COND #( WHEN lv_mob IS NOT INITIAL THEN |MO : { lv_mob }| ) ) )
                 iv_sep   = ` ` ) TO rt_lines.

    DATA(lv_mail) = comm_value( iv_adrnr = is_party-adrnr iv_kind = 'E' ).
    IF lv_mail IS NOT INITIAL.
      APPEND |E-MAIL : { lv_mail }| TO rt_lines.
    ENDIF.

    IF iv_with_tax_no = abap_true AND is_party-stcd1 IS NOT INITIAL.
      APPEND |TAX NUMBER : { is_party-stcd1 }| TO rt_lines.
    ENDIF.

    DELETE rt_lines WHERE table_line IS INITIAL.

  ENDMETHOD.


  METHOD order_reference.

    " FS #8 / logic B13: customer reference + date of the QUOTATION the
    " sales orders were created from. Several orders -> joined with " & "
    " as on sample 90000014 (ASSUMPTION A17).
    TYPES: BEGIN OF ty_doc,
             vbeln TYPE vbeln,
           END OF ty_doc.
    DATA lt_so  TYPE SORTED TABLE OF ty_doc WITH UNIQUE KEY vbeln.
    DATA lt_ref TYPE SORTED TABLE OF ty_doc WITH UNIQUE KEY vbeln.
    DATA lt_bstkd TYPE string_table.
    DATA lt_bstdk TYPE string_table.

    LOOP AT mt_vbrp INTO DATA(ls_vbrp) WHERE aubel IS NOT INITIAL.
      INSERT VALUE #( vbeln = ls_vbrp-aubel ) INTO TABLE lt_so.
    ENDLOOP.
    IF lt_so IS INITIAL.
      RETURN.
    ENDIF.

    SELECT vbeln, vgbel, vgtyp FROM vbak
      FOR ALL ENTRIES IN @lt_so
      WHERE vbeln = @lt_so-vbeln
      INTO TABLE @DATA(lt_vbak).

    " Quotation (VBTYP 'B') if the order was created with reference to
    " one, otherwise the order itself
    LOOP AT lt_vbak INTO DATA(ls_vbak).
      INSERT VALUE #( vbeln = COND #( WHEN ls_vbak-vgbel IS NOT INITIAL AND ls_vbak-vgtyp = 'B'
                                      THEN ls_vbak-vgbel ELSE ls_vbak-vbeln ) )
             INTO TABLE lt_ref.
    ENDLOOP.

    SELECT vbeln, bstkd, bstdk FROM vbkd
      FOR ALL ENTRIES IN @lt_ref
      WHERE vbeln = @lt_ref-vbeln
        AND posnr = '000000'
      INTO TABLE @DATA(lt_vbkd).
    SORT lt_vbkd BY vbeln.

    LOOP AT lt_vbkd INTO DATA(ls_vbkd).
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


  METHOD payment_lines.

    DATA(ls_vbrk) = mt_vbrk[ 1 ].

    " Logic B43: TVZBT-VTEXT (first day-limit entry)
    SELECT vtext FROM tvzbt
      WHERE spras = @zif_sd_epack=>gc_langu
        AND zterm = @ls_vbrk-zterm
      ORDER BY ztagg
      INTO @DATA(lv_vtext)
      UP TO 1 ROWS.
    ENDSELECT.

    " Logic B44: Incoterms 1 + 2; description from TINCT (ASSUMPTION A18)
    SELECT SINGLE bezei FROM tinct
      WHERE spras = @zif_sd_epack=>gc_langu
        AND inco1 = @ls_vbrk-inco1
      INTO @DATA(lv_inco_text).

    APPEND |PAYMENT TERM: { COND string( WHEN lv_vtext IS NOT INITIAL THEN lv_vtext ELSE ls_vbrk-zterm ) }|
           TO rt_lines.
    APPEND condense( |TERMS OF SHIPMENT :- {
                       COND string( WHEN lv_inco_text IS NOT INITIAL THEN to_upper( lv_inco_text )
                                    ELSE ls_vbrk-inco1 ) } { ls_vbrk-inco2 }| )
           TO rt_lines.

  ENDMETHOD.


  METHOD comm_value.

    " Preferred: the requested R3_USER flag; otherwise default / first entry
    IF iv_r3_user IS NOT INITIAL.
      LOOP AT mt_comm INTO DATA(ls_comm)
           WHERE addrnumber = iv_adrnr AND kind = iv_kind AND r3_user = iv_r3_user.
        rv_value = ls_comm-value.
        RETURN.
      ENDLOOP.
    ENDIF.

    LOOP AT mt_comm INTO ls_comm WHERE addrnumber = iv_adrnr AND kind = iv_kind.
      rv_value = ls_comm-value.
      RETURN.
    ENDLOOP.

  ENDMETHOD.


  METHOD country_text.

    IF iv_land1 IS INITIAL.
      RETURN.
    ENDIF.
    SELECT SINGLE landx FROM t005t
      WHERE spras = @zif_sd_epack=>gc_langu
        AND land1 = @iv_land1
      INTO @DATA(lv_landx).
    rv_text = COND #( WHEN sy-subrc = 0 THEN to_upper( lv_landx ) ELSE iv_land1 ).

  ENDMETHOD.


  METHOD region_text.

    IF iv_region IS INITIAL.
      RETURN.
    ENDIF.
    SELECT SINGLE bezei FROM t005u
      WHERE spras = @zif_sd_epack=>gc_langu
        AND land1 = @iv_land1
        AND bland = @iv_region
      INTO @DATA(lv_bezei).
    rv_text = COND #( WHEN sy-subrc = 0 THEN to_upper( lv_bezei ) ELSE iv_region ).

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

    " Lines longer than the text field are split; the form wraps text
    " at the window width anyway.
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


  METHOD to_matnr.

    " Part No. = SAP material number (ASSUMPTION A12)
    IF iv_partno IS INITIAL.
      RETURN.
    ENDIF.
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

ENDCLASS.
