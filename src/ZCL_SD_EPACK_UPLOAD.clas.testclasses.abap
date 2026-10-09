*"* use this source file for your ABAP unit test classes
*&---------------------------------------------------------------------*
*& WRICEF 102-B "Export consolidated packing list" - Excel parser tests
*&
*& Covers the test-plan scenarios that do not need the database:
*& one / several invoices per packing list, several packing lists per
*& file, duplicate and invalid invoice numbers, missing mandatory data,
*& duplicate rows and the strict separation of items per packing list.
*& Rows are built in memory in the layout of "Z table Format1.xlsx".
*&---------------------------------------------------------------------*
CLASS ltc_parser DEFINITION FINAL
  FOR TESTING
  DURATION SHORT
  RISK LEVEL HARMLESS.

  PRIVATE SECTION.

    CONSTANTS gc_pl_a TYPE zsd_packno VALUE '0090000014'.
    CONSTANTS gc_pl_b TYPE zsd_packno VALUE '0090000081'.

    DATA mo_cut TYPE REF TO zcl_sd_epack_upload.

    METHODS setup.

    METHODS one_invoice_one_pl           FOR TESTING.
    METHODS several_invoices_one_pl      FOR TESTING.
    METHODS several_pls_in_one_file      FOR TESTING.
    METHODS items_never_cross_pls        FOR TESTING.
    METHODS duplicate_invoice_once       FOR TESTING.
    METHODS invalid_invoice_rejected     FOR TESTING.
    METHODS invoice_in_two_pls_rejected  FOR TESTING.
    METHODS missing_invoice_rejected     FOR TESTING.
    METHODS missing_pl_date_rejected     FOR TESTING.
    METHODS rows_before_first_pl         FOR TESTING.
    METHODS no_pl_number_in_file         FOR TESTING.
    METHODS header_conflict_in_pl        FOR TESTING.
    METHODS header_differs_between_pls   FOR TESTING.
    METHODS exact_duplicate_row_ignored  FOR TESTING.
    METHODS conflicting_sl_no_rejected   FOR TESTING.
    METHODS pl_without_items_rejected    FOR TESTING.
    METHODS wrong_template_rejected      FOR TESTING.
    METHODS total_row_is_not_an_item     FOR TESTING.
    METHODS weight_error_found_in_upload FOR TESTING.
    METHODS empty_file_rejected          FOR TESTING.

    METHODS heading
      RETURNING VALUE(rs_row) TYPE zcl_sd_epack_upload=>ty_row.

    METHODS line
      IMPORTING iv_row        TYPE i
                iv_sn         TYPE string
                iv_part       TYPE string DEFAULT `M071140316`
                iv_std        TYPE string DEFAULT `1`
                iv_art        TYPE string DEFAULT `60`
                iv_qty        TYPE string DEFAULT `60`
                iv_gwt        TYPE string DEFAULT `2121`
                iv_nwt        TYPE string DEFAULT `2107.2`
                iv_inv        TYPE string OPTIONAL
                iv_packno     TYPE string OPTIONAL
                iv_packdt     TYPE string OPTIONAL
                iv_pol        TYPE string OPTIONAL
      RETURNING VALUE(rs_row) TYPE zcl_sd_epack_upload=>ty_row.

    METHODS parse
      IMPORTING it_rows     TYPE zcl_sd_epack_upload=>tt_rows
      EXPORTING et_pl       TYPE zif_sd_epack=>tt_packing_list
                et_messages TYPE bapiret2_t.

    METHODS has_msg
      IMPORTING it_messages      TYPE bapiret2_t
                iv_no            TYPE symsgno
      RETURNING VALUE(rv_exists) TYPE abap_bool.

ENDCLASS.


CLASS ltc_parser IMPLEMENTATION.

  METHOD setup.
    mo_cut = NEW #( ).
  ENDMETHOD.


  METHOD heading.
    rs_row-row   = 1.
    rs_row-cells = VALUE #( FOR i = 1 UNTIL i > 38 ( `` ) ).
    rs_row-cells[ 1 ]  = `Sl.No.`.
    rs_row-cells[ 12 ] = `SAP Invoice Number`.
    rs_row-cells[ 13 ] = `Packing list no`.
  ENDMETHOD.


  METHOD line.
    rs_row-row   = iv_row.
    rs_row-cells = VALUE #( FOR i = 1 UNTIL i > 38 ( `` ) ).
    rs_row-cells[ 1 ]  = iv_sn.
    rs_row-cells[ 4 ]  = iv_std.
    rs_row-cells[ 5 ]  = iv_art.
    rs_row-cells[ 6 ]  = iv_part.
    rs_row-cells[ 8 ]  = iv_qty.
    rs_row-cells[ 9 ]  = iv_gwt.
    rs_row-cells[ 10 ] = iv_nwt.
    rs_row-cells[ 11 ] = `3917 2390`.
    rs_row-cells[ 12 ] = iv_inv.
    rs_row-cells[ 13 ] = iv_packno.
    rs_row-cells[ 14 ] = iv_packdt.
    rs_row-cells[ 26 ] = iv_pol.
  ENDMETHOD.


  METHOD parse.
    mo_cut->parse_rows( EXPORTING it_rows          = it_rows
                        IMPORTING et_packing_lists = et_pl
                                  et_messages      = et_messages ).
  ENDMETHOD.


  METHOD has_msg.
    rv_exists = xsdbool( line_exists( it_messages[ number = iv_no ] ) ).
  ENDMETHOD.


  METHOD one_invoice_one_pl.

    parse( EXPORTING it_rows = VALUE #( ( heading( ) )
                                        ( line( iv_row = 2 iv_sn = `1` iv_inv = `90000101`
                                                iv_packno = `90000014` iv_packdt = `27.05.2025` ) )
                                        ( line( iv_row = 3 iv_sn = `2` ) ) )
           IMPORTING et_pl       = DATA(lt_pl)
                     et_messages = DATA(lt_msg) ).

    cl_abap_unit_assert=>assert_initial( lt_msg ).
    cl_abap_unit_assert=>assert_equals( act = lines( lt_pl ) exp = 1 ).
    DATA(ls_pl) = lt_pl[ packno = gc_pl_a ].
    cl_abap_unit_assert=>assert_equals( act = lines( ls_pl-invoices ) exp = 1 ).
    cl_abap_unit_assert=>assert_equals( act = lines( ls_pl-items )    exp = 2 ).
    cl_abap_unit_assert=>assert_equals( act = ls_pl-header-zsd_packdt exp = '20250527' ).
    cl_abap_unit_assert=>assert_false( zcl_sd_epack_rules=>has_errors( ls_pl-messages ) ).

  ENDMETHOD.


  METHOD several_invoices_one_pl.

    parse( EXPORTING it_rows = VALUE #( ( heading( ) )
                                        ( line( iv_row = 2 iv_sn = `1` iv_inv = `90000101, 90000102`
                                                iv_packno = `90000014` iv_packdt = `27.05.2025` ) )
                                        ( line( iv_row = 3 iv_sn = `2` iv_inv = `90000103` ) ) )
           IMPORTING et_pl = DATA(lt_pl) ).

    DATA(ls_pl) = lt_pl[ packno = gc_pl_a ].
    cl_abap_unit_assert=>assert_equals( act = lines( ls_pl-invoices ) exp = 3 ).
    cl_abap_unit_assert=>assert_true( xsdbool( line_exists( ls_pl-invoices[ table_line = '0090000103' ] ) ) ).

  ENDMETHOD.


  METHOD several_pls_in_one_file.

    parse( EXPORTING it_rows = VALUE #( ( heading( ) )
                                        ( line( iv_row = 2 iv_sn = `1` iv_inv = `90000101`
                                                iv_packno = `90000014` iv_packdt = `27.05.2025` ) )
                                        ( line( iv_row = 3 iv_sn = `2` ) )
                                        ( line( iv_row = 4 iv_sn = `1` iv_inv = `90000201`
                                                iv_packno = `90000081` iv_packdt = `08.10.2025` ) ) )
           IMPORTING et_pl       = DATA(lt_pl)
                     et_messages = DATA(lt_msg) ).

    cl_abap_unit_assert=>assert_initial( lt_msg ).
    cl_abap_unit_assert=>assert_equals( act = lines( lt_pl ) exp = 2 ).
    cl_abap_unit_assert=>assert_equals( act = lines( lt_pl[ packno = gc_pl_a ]-items ) exp = 2 ).
    cl_abap_unit_assert=>assert_equals( act = lines( lt_pl[ packno = gc_pl_b ]-items ) exp = 1 ).
    cl_abap_unit_assert=>assert_equals( act = lt_pl[ packno = gc_pl_b ]-header-zsd_packdt exp = '20251008' ).

  ENDMETHOD.


  METHOD items_never_cross_pls.

    parse( EXPORTING it_rows = VALUE #( ( heading( ) )
                                        ( line( iv_row = 2 iv_sn = `1` iv_part = `PART-A1` iv_inv = `90000101`
                                                iv_packno = `90000014` iv_packdt = `27.05.2025` ) )
                                        ( line( iv_row = 3 iv_sn = `1` iv_part = `PART-B1` iv_inv = `90000201`
                                                iv_packno = `90000081` iv_packdt = `08.10.2025` ) )
                                        ( line( iv_row = 4 iv_sn = `2` iv_part = `PART-B2` ) ) )
           IMPORTING et_pl = DATA(lt_pl) ).

    DATA(ls_a) = lt_pl[ packno = gc_pl_a ].
    DATA(ls_b) = lt_pl[ packno = gc_pl_b ].

    LOOP AT ls_a-items INTO DATA(ls_item).
      cl_abap_unit_assert=>assert_equals( act = ls_item-zsd_packno exp = gc_pl_a ).
      cl_abap_unit_assert=>assert_char_cp( act = ls_item-zsd_partno exp = 'PART-A*' ).
    ENDLOOP.
    LOOP AT ls_b-items INTO ls_item.
      cl_abap_unit_assert=>assert_equals( act = ls_item-zsd_packno exp = gc_pl_b ).
      cl_abap_unit_assert=>assert_char_cp( act = ls_item-zsd_partno exp = 'PART-B*' ).
    ENDLOOP.
    cl_abap_unit_assert=>assert_equals( act = ls_a-invoices exp = VALUE zif_sd_epack=>tt_vbeln( ( '0090000101' ) ) ).
    cl_abap_unit_assert=>assert_equals( act = ls_b-invoices exp = VALUE zif_sd_epack=>tt_vbeln( ( '0090000201' ) ) ).

  ENDMETHOD.


  METHOD duplicate_invoice_once.

    parse( EXPORTING it_rows = VALUE #( ( heading( ) )
                                        ( line( iv_row = 2 iv_sn = `1` iv_inv = `90000101`
                                                iv_packno = `90000014` iv_packdt = `27.05.2025` ) )
                                        ( line( iv_row = 3 iv_sn = `2` iv_inv = `90000101` ) ) )
           IMPORTING et_pl = DATA(lt_pl) ).

    DATA(ls_pl) = lt_pl[ packno = gc_pl_a ].
    cl_abap_unit_assert=>assert_equals( act = lines( ls_pl-invoices ) exp = 1 ).
    cl_abap_unit_assert=>assert_true( has_msg( it_messages = ls_pl-messages iv_no = '056' ) ).
    cl_abap_unit_assert=>assert_false( zcl_sd_epack_rules=>has_errors( ls_pl-messages ) ).

  ENDMETHOD.


  METHOD invalid_invoice_rejected.

    parse( EXPORTING it_rows = VALUE #( ( heading( ) )
                                        ( line( iv_row = 2 iv_sn = `1` iv_inv = `90000101 / INV-X`
                                                iv_packno = `90000014` iv_packdt = `27.05.2025` ) ) )
           IMPORTING et_pl = DATA(lt_pl) ).

    DATA(ls_pl) = lt_pl[ packno = gc_pl_a ].
    cl_abap_unit_assert=>assert_true( has_msg( it_messages = ls_pl-messages iv_no = '018' ) ).
    cl_abap_unit_assert=>assert_equals( act = lines( ls_pl-invoices ) exp = 1 ).

  ENDMETHOD.


  METHOD invoice_in_two_pls_rejected.

    parse( EXPORTING it_rows = VALUE #( ( heading( ) )
                                        ( line( iv_row = 2 iv_sn = `1` iv_inv = `90000101`
                                                iv_packno = `90000014` iv_packdt = `27.05.2025` ) )
                                        ( line( iv_row = 3 iv_sn = `1` iv_inv = `90000101`
                                                iv_packno = `90000081` iv_packdt = `08.10.2025` ) ) )
           IMPORTING et_pl = DATA(lt_pl) ).

    cl_abap_unit_assert=>assert_true( has_msg( it_messages = lt_pl[ packno = gc_pl_a ]-messages iv_no = '058' ) ).
    cl_abap_unit_assert=>assert_true( has_msg( it_messages = lt_pl[ packno = gc_pl_b ]-messages iv_no = '058' ) ).

  ENDMETHOD.


  METHOD missing_invoice_rejected.

    parse( EXPORTING it_rows = VALUE #( ( heading( ) )
                                        ( line( iv_row = 2 iv_sn = `1`
                                                iv_packno = `90000014` iv_packdt = `27.05.2025` ) ) )
           IMPORTING et_pl = DATA(lt_pl) ).

    DATA(lt_msg) = lt_pl[ packno = gc_pl_a ]-messages.
    cl_abap_unit_assert=>assert_true( has_msg( it_messages = lt_msg iv_no = '017' ) ).
    cl_abap_unit_assert=>assert_true( zcl_sd_epack_rules=>has_errors( lt_msg ) ).

  ENDMETHOD.


  METHOD missing_pl_date_rejected.

    parse( EXPORTING it_rows = VALUE #( ( heading( ) )
                                        ( line( iv_row = 2 iv_sn = `1` iv_inv = `90000101`
                                                iv_packno = `90000014` ) ) )
           IMPORTING et_pl = DATA(lt_pl) ).

    DATA(lt_msg) = lt_pl[ packno = gc_pl_a ]-messages.
    cl_abap_unit_assert=>assert_equals( act = lt_msg[ number = '017' ]-message_v2 exp = 'Packing list date' ).

  ENDMETHOD.


  METHOD rows_before_first_pl.

    parse( EXPORTING it_rows = VALUE #( ( heading( ) )
                                        ( line( iv_row = 2 iv_sn = `1` ) )
                                        ( line( iv_row = 3 iv_sn = `1` iv_inv = `90000101`
                                                iv_packno = `90000014` iv_packdt = `27.05.2025` ) ) )
           IMPORTING et_pl       = DATA(lt_pl)
                     et_messages = DATA(lt_msg) ).

    cl_abap_unit_assert=>assert_true( has_msg( it_messages = lt_msg iv_no = '055' ) ).
    cl_abap_unit_assert=>assert_equals( act = lines( lt_pl[ packno = gc_pl_a ]-items ) exp = 1 ).

  ENDMETHOD.


  METHOD no_pl_number_in_file.

    parse( EXPORTING it_rows = VALUE #( ( heading( ) )
                                        ( line( iv_row = 2 iv_sn = `1` iv_inv = `90000101` ) ) )
           IMPORTING et_pl       = DATA(lt_pl)
                     et_messages = DATA(lt_msg) ).

    cl_abap_unit_assert=>assert_initial( lt_pl ).
    cl_abap_unit_assert=>assert_true( has_msg( it_messages = lt_msg iv_no = '059' ) ).

  ENDMETHOD.


  METHOD header_conflict_in_pl.

    parse( EXPORTING it_rows = VALUE #( ( heading( ) )
                                        ( line( iv_row = 2 iv_sn = `1` iv_inv = `90000101` iv_pol = `MUNDRA`
                                                iv_packno = `90000014` iv_packdt = `27.05.2025` ) )
                                        ( line( iv_row = 3 iv_sn = `2` iv_pol = `NHAVA SHEVA` ) ) )
           IMPORTING et_pl = DATA(lt_pl) ).

    cl_abap_unit_assert=>assert_true( has_msg( it_messages = lt_pl[ packno = gc_pl_a ]-messages iv_no = '020' ) ).

  ENDMETHOD.


  METHOD header_differs_between_pls.

    " Different header values in DIFFERENT packing lists are not a conflict
    parse( EXPORTING it_rows = VALUE #( ( heading( ) )
                                        ( line( iv_row = 2 iv_sn = `1` iv_inv = `90000101` iv_pol = `MUNDRA`
                                                iv_packno = `90000014` iv_packdt = `27.05.2025` ) )
                                        ( line( iv_row = 3 iv_sn = `1` iv_inv = `90000201` iv_pol = `NHAVA SHEVA`
                                                iv_packno = `90000081` iv_packdt = `08.10.2025` ) ) )
           IMPORTING et_pl = DATA(lt_pl) ).

    cl_abap_unit_assert=>assert_false( has_msg( it_messages = lt_pl[ packno = gc_pl_a ]-messages iv_no = '020' ) ).
    cl_abap_unit_assert=>assert_equals( act = lt_pl[ packno = gc_pl_a ]-header-zsd_pol exp = 'MUNDRA' ).
    cl_abap_unit_assert=>assert_equals( act = lt_pl[ packno = gc_pl_b ]-header-zsd_pol exp = 'NHAVA SHEVA' ).

  ENDMETHOD.


  METHOD exact_duplicate_row_ignored.

    parse( EXPORTING it_rows = VALUE #( ( heading( ) )
                                        ( line( iv_row = 2 iv_sn = `1` iv_inv = `90000101`
                                                iv_packno = `90000014` iv_packdt = `27.05.2025` ) )
                                        ( line( iv_row = 3 iv_sn = `2` ) )
                                        ( line( iv_row = 4 iv_sn = `2` ) ) )
           IMPORTING et_pl = DATA(lt_pl) ).

    DATA(ls_pl) = lt_pl[ packno = gc_pl_a ].
    cl_abap_unit_assert=>assert_equals( act = lines( ls_pl-items ) exp = 2 ).
    cl_abap_unit_assert=>assert_true( has_msg( it_messages = ls_pl-messages iv_no = '057' ) ).
    cl_abap_unit_assert=>assert_false( zcl_sd_epack_rules=>has_errors( ls_pl-messages ) ).

  ENDMETHOD.


  METHOD conflicting_sl_no_rejected.

    parse( EXPORTING it_rows = VALUE #( ( heading( ) )
                                        ( line( iv_row = 2 iv_sn = `1` iv_inv = `90000101`
                                                iv_packno = `90000014` iv_packdt = `27.05.2025` ) )
                                        ( line( iv_row = 3 iv_sn = `1` iv_qty = `61` ) ) )
           IMPORTING et_pl = DATA(lt_pl) ).

    cl_abap_unit_assert=>assert_true( has_msg( it_messages = lt_pl[ packno = gc_pl_a ]-messages iv_no = '021' ) ).

  ENDMETHOD.


  METHOD pl_without_items_rejected.

    " Header row of packing list B carries no Sl.No. -> no line items
    DATA(ls_row) = line( iv_row = 3 iv_sn = `` iv_part = `` iv_qty = `` iv_inv = `90000201`
                         iv_packno = `90000081` iv_packdt = `08.10.2025` ).

    parse( EXPORTING it_rows = VALUE #( ( heading( ) )
                                        ( line( iv_row = 2 iv_sn = `1` iv_inv = `90000101`
                                                iv_packno = `90000014` iv_packdt = `27.05.2025` ) )
                                        ( ls_row ) )
           IMPORTING et_pl = DATA(lt_pl) ).

    cl_abap_unit_assert=>assert_true( has_msg( it_messages = lt_pl[ packno = gc_pl_b ]-messages iv_no = '036' ) ).
    cl_abap_unit_assert=>assert_false( zcl_sd_epack_rules=>has_errors( lt_pl[ packno = gc_pl_a ]-messages ) ).

  ENDMETHOD.


  METHOD wrong_template_rejected.

    DATA(ls_heading) = heading( ).
    ls_heading-cells[ 13 ] = `Invoice date`.

    parse( EXPORTING it_rows = VALUE #( ( ls_heading )
                                        ( line( iv_row = 2 iv_sn = `1` iv_inv = `90000101`
                                                iv_packno = `90000014` iv_packdt = `27.05.2025` ) ) )
           IMPORTING et_pl       = DATA(lt_pl)
                     et_messages = DATA(lt_msg) ).

    cl_abap_unit_assert=>assert_initial( lt_pl ).
    cl_abap_unit_assert=>assert_true( has_msg( it_messages = lt_msg iv_no = '016' ) ).

  ENDMETHOD.


  METHOD total_row_is_not_an_item.

    " Like row 83 of Z table Format1.xlsx: no Sl.No., totals in E-J
    DATA(ls_total) = line( iv_row = 4 iv_sn = `` iv_part = `` iv_art = `120` iv_qty = `120` ).

    parse( EXPORTING it_rows = VALUE #( ( heading( ) )
                                        ( line( iv_row = 2 iv_sn = `1` iv_inv = `90000101`
                                                iv_packno = `90000014` iv_packdt = `27.05.2025` ) )
                                        ( line( iv_row = 3 iv_sn = `2` ) )
                                        ( ls_total ) )
           IMPORTING et_pl = DATA(lt_pl) ).

    DATA(ls_pl) = lt_pl[ packno = gc_pl_a ].
    cl_abap_unit_assert=>assert_equals( act = lines( ls_pl-items ) exp = 2 ).
    cl_abap_unit_assert=>assert_true( has_msg( it_messages = ls_pl-messages iv_no = '050' ) ).

  ENDMETHOD.


  METHOD weight_error_found_in_upload.

    parse( EXPORTING it_rows = VALUE #( ( heading( ) )
                                        ( line( iv_row = 2 iv_sn = `1` iv_inv = `90000101` iv_gwt = `100` iv_nwt = `101`
                                                iv_packno = `90000014` iv_packdt = `27.05.2025` ) ) )
           IMPORTING et_pl = DATA(lt_pl) ).

    cl_abap_unit_assert=>assert_true( has_msg( it_messages = lt_pl[ packno = gc_pl_a ]-messages iv_no = '001' ) ).

  ENDMETHOD.


  METHOD empty_file_rejected.

    parse( EXPORTING it_rows = VALUE #( ( heading( ) ) )
           IMPORTING et_pl       = DATA(lt_pl)
                     et_messages = DATA(lt_msg) ).

    cl_abap_unit_assert=>assert_initial( lt_pl ).
    cl_abap_unit_assert=>assert_true( has_msg( it_messages = lt_msg iv_no = '036' ) ).

  ENDMETHOD.

ENDCLASS.
