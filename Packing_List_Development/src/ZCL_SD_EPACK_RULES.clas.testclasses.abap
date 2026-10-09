*"* use this source file for your ABAP unit test classes
*&---------------------------------------------------------------------*
*& WRICEF 102-B "Export consolidated packing list" - rule unit tests
*&
*& Test data is taken from the repository samples:
*&   Z table Format1.xlsx / PackingList_90000081 (Nabil, merged groups)
*&   PackingList_90000014 (Tarmac, one weight for three article lines,
*&                         two package types)
*&---------------------------------------------------------------------*
CLASS ltc_rules DEFINITION FINAL
  FOR TESTING
  DURATION SHORT
  RISK LEVEL HARMLESS.

  PRIVATE SECTION.

    " Package groups
    METHODS merged_rows_share_groups      FOR TESTING.
    METHODS tarmac_weight_group           FOR TESTING.
    METHODS first_line_needs_articles     FOR TESTING.
    METHODS gross_without_net_rejected    FOR TESTING.
    METHODS duplicate_sl_no_rejected      FOR TESTING.
    METHODS descending_sl_no_rejected     FOR TESTING.
    METHODS package_type_carried_forward  FOR TESTING.

    " FS 2.1 weight validation
    METHODS gross_below_net_is_error      FOR TESTING.
    METHODS gross_equal_net_is_accepted   FOR TESTING.

    " Quantity = STD x Articles
    METHODS qty_rule_uses_group_articles  FOR TESTING.
    METHODS qty_mismatch_is_warning       FOR TESTING.

    " Totals
    METHODS totals_of_nabil_rows          FOR TESTING.
    METHODS total_text_single_type        FOR TESTING.
    METHODS total_text_two_types          FOR TESTING.
    METHODS total_text_no_type            FOR TESTING.

    " Printed rows
    METHODS print_rows_heading_subtotal   FOR TESTING.
    METHODS print_rows_blank_in_group     FOR TESTING.

    " Formatting / parsing
    METHODS hsn_is_formatted              FOR TESTING.
    METHODS heading_from_goods_text       FOR TESTING.
    METHODS heading_fallback_t604n        FOR TESTING.
    METHODS weight_three_decimals         FOR TESTING.
    METHODS date_formats_are_parsed       FOR TESTING.
    METHODS invalid_date_is_rejected      FOR TESTING.
    METHODS numbers_are_parsed            FOR TESTING.
    METHODS invoice_list_is_split         FOR TESTING.
    METHODS multi_line_cell_is_split      FOR TESTING.

    "! Nabil sample rows 1-4 (Z table Format1.xlsx rows 2-5)
    METHODS nabil_lines
      RETURNING VALUE(rt_lines) TYPE zif_sd_epack=>tt_upload_line.

    "! Tarmac sample lines 1-3 + 14 (pipes share one weight; then a box)
    METHODS tarmac_lines
      RETURNING VALUE(rt_lines) TYPE zif_sd_epack=>tt_upload_line.

    METHODS items
      IMPORTING it_lines        TYPE zif_sd_epack=>tt_upload_line
      RETURNING VALUE(rt_items) TYPE zif_sd_epack=>tt_item.

    METHODS errors
      IMPORTING it_lines           TYPE zif_sd_epack=>tt_upload_line
      RETURNING VALUE(rt_messages) TYPE bapiret2_t.

ENDCLASS.


CLASS ltc_rules IMPLEMENTATION.

  METHOD nabil_lines.
    rt_lines = VALUE #(
      ( row = 2 sn = 1 inv_sn = 10 artno = '1 TO 60'  std = 1 art = 60 art_filled = abap_true
        partno = 'M071140316' qty = 60  gwt = '2121.000' gwt_filled = abap_true
        nwt = '2107.200' nwt_filled = abap_true hsn = '3917 2390' pkg_type = 'Bundles' )
      ( row = 3 sn = 2 inv_sn = 5  std = 1
        partno = 'M071340307' qty = 60  hsn = '3917 2390' )
      ( row = 4 sn = 3 inv_sn = 9 artno = '61 TO 130' std = 2 art = 70 art_filled = abap_true
        partno = 'M071340309' qty = 140 gwt = '3446.100' gwt_filled = abap_true
        nwt = '3430.000' nwt_filled = abap_true hsn = '3917 2390' )
      ( row = 5 sn = 4 inv_sn = 5  std = 2
        partno = 'M071340307' qty = 140 hsn = '3917 2390' ) ).
  ENDMETHOD.


  METHOD tarmac_lines.
    rt_lines = VALUE #(
      ( row = 2 sn = 1 inv_sn = 3 std = 1 art = 171 art_filled = abap_true
        partno = 'M241270309' qty = 171 gwt = '3398.310' gwt_filled = abap_true
        nwt = '3353.310' nwt_filled = abap_true hsn = '3917 2390' pkg_type = 'Loose Pipes' )
      ( row = 3 sn = 2 inv_sn = 2 std = 1 art = 171 art_filled = abap_true
        partno = 'M241270307' qty = 171 hsn = '3917 2390' )
      ( row = 4 sn = 3 inv_sn = 1 std = 1 art = 171 art_filled = abap_true
        partno = 'M241270305' qty = 171 hsn = '3917 2390' )
      ( row = 5 sn = 14 inv_sn = 18 artno = '1 TO 24' std = 15 art = 24 art_filled = abap_true
        partno = 'M242002129' qty = 360 gwt = '459.840' gwt_filled = abap_true
        nwt = '424.320' nwt_filled = abap_true hsn = '3917 4000' pkg_type = 'Boxes' ) ).
  ENDMETHOD.


  METHOD items.
    DATA lt_messages TYPE bapiret2_t.
    zcl_sd_epack_rules=>build_items( EXPORTING iv_packno   = CONV zsd_packno( '90000081' )
                                               it_lines    = it_lines
                                     IMPORTING et_items    = rt_items
                                     CHANGING  ct_messages = lt_messages ).
  ENDMETHOD.


  METHOD errors.
    DATA lt_items TYPE zif_sd_epack=>tt_item.
    zcl_sd_epack_rules=>build_items( EXPORTING iv_packno   = CONV zsd_packno( '90000081' )
                                               it_lines    = it_lines
                                     IMPORTING et_items    = lt_items
                                     CHANGING  ct_messages = rt_messages ).
  ENDMETHOD.


  METHOD merged_rows_share_groups.

    DATA(lt_items) = items( nabil_lines( ) ).

    cl_abap_unit_assert=>assert_equals( act = lt_items[ 2 ]-zsd_art_grp exp = lt_items[ 1 ]-zsd_art_grp ).
    cl_abap_unit_assert=>assert_equals( act = lt_items[ 2 ]-zsd_wt_grp  exp = lt_items[ 1 ]-zsd_wt_grp ).
    cl_abap_unit_assert=>assert_differs( act = lt_items[ 3 ]-zsd_art_grp exp = lt_items[ 2 ]-zsd_art_grp ).
    " Weights stay on the first line of the group only
    cl_abap_unit_assert=>assert_initial( lt_items[ 2 ]-zsd_gwt ).
    cl_abap_unit_assert=>assert_initial( errors( nabil_lines( ) ) ).

  ENDMETHOD.


  METHOD tarmac_weight_group.

    DATA(lt_items) = items( tarmac_lines( ) ).

    " Three article groups, one weight group (lines 1-3 of 90000014)
    cl_abap_unit_assert=>assert_equals( act = lt_items[ 3 ]-zsd_art_grp exp = 3 ).
    cl_abap_unit_assert=>assert_equals( act = lt_items[ 3 ]-zsd_wt_grp  exp = 1 ).
    cl_abap_unit_assert=>assert_equals( act = lt_items[ 4 ]-zsd_wt_grp  exp = 2 ).

  ENDMETHOD.


  METHOD first_line_needs_articles.

    DATA(lt_lines) = nabil_lines( ).
    lt_lines[ 1 ]-art_filled = abap_false.

    DATA(lt_messages) = errors( lt_lines ).

    cl_abap_unit_assert=>assert_true( xsdbool( line_exists( lt_messages[ number = '027' ] ) ) ).

  ENDMETHOD.


  METHOD gross_without_net_rejected.

    DATA(lt_lines) = nabil_lines( ).
    lt_lines[ 3 ]-nwt_filled = abap_false.

    cl_abap_unit_assert=>assert_true( xsdbool( line_exists( errors( lt_lines )[ number = '028' ] ) ) ).

  ENDMETHOD.


  METHOD duplicate_sl_no_rejected.

    DATA(lt_lines) = nabil_lines( ).
    lt_lines[ 2 ]-sn = 1.

    cl_abap_unit_assert=>assert_true( xsdbool( line_exists( errors( lt_lines )[ number = '021' ] ) ) ).

  ENDMETHOD.


  METHOD descending_sl_no_rejected.

    DATA(lt_lines) = nabil_lines( ).
    lt_lines[ 3 ]-sn = 9.

    cl_abap_unit_assert=>assert_true( xsdbool( line_exists( errors( lt_lines )[ number = '041' ] ) ) ).

  ENDMETHOD.


  METHOD package_type_carried_forward.

    DATA(lt_items) = items( tarmac_lines( ) ).

    cl_abap_unit_assert=>assert_equals( act = lt_items[ 3 ]-zsd_pkg_type exp = 'Loose Pipes' ).
    cl_abap_unit_assert=>assert_equals( act = lt_items[ 4 ]-zsd_pkg_type exp = 'Boxes' ).

  ENDMETHOD.


  METHOD gross_below_net_is_error.

    DATA(lt_lines) = nabil_lines( ).
    lt_lines[ 3 ]-gwt = '3000.000'.

    DATA(lt_messages) = zcl_sd_epack_rules=>check_weights( items( lt_lines ) ).

    " FS text first (001), then the detail of the failing group (002)
    cl_abap_unit_assert=>assert_equals( act = lt_messages[ 1 ]-number exp = '001' ).
    cl_abap_unit_assert=>assert_equals( act = lt_messages[ 1 ]-type   exp = 'E' ).
    cl_abap_unit_assert=>assert_equals( act = lt_messages[ 2 ]-message_v1 exp = '3' ).
    cl_abap_unit_assert=>assert_equals( act = lines( lt_messages ) exp = 2 ).

  ENDMETHOD.


  METHOD gross_equal_net_is_accepted.

    " FS: "Gross Weight is always equal to or greater than the Net Weight"
    DATA(lt_lines) = nabil_lines( ).
    lt_lines[ 1 ]-gwt = lt_lines[ 1 ]-nwt.

    cl_abap_unit_assert=>assert_initial( zcl_sd_epack_rules=>check_weights( items( lt_lines ) ) ).

  ENDMETHOD.


  METHOD qty_rule_uses_group_articles.

    " Line 2 has no Articles of its own: 1 x 60 (merged) = 60
    cl_abap_unit_assert=>assert_initial( zcl_sd_epack_rules=>check_quantities( items( nabil_lines( ) ) ) ).

  ENDMETHOD.


  METHOD qty_mismatch_is_warning.

    DATA(lt_lines) = nabil_lines( ).
    lt_lines[ 4 ]-qty = 139.

    DATA(lt_messages) = zcl_sd_epack_rules=>check_quantities( items( lt_lines ) ).

    cl_abap_unit_assert=>assert_equals( act = lines( lt_messages ) exp = 1 ).
    cl_abap_unit_assert=>assert_equals( act = lt_messages[ 1 ]-type exp = 'W' ).
    cl_abap_unit_assert=>assert_false( zcl_sd_epack_rules=>has_errors( lt_messages ) ).

  ENDMETHOD.


  METHOD totals_of_nabil_rows.

    DATA(ls_totals) = zcl_sd_epack_rules=>calc_totals( items( nabil_lines( ) ) ).

    cl_abap_unit_assert=>assert_equals( act = ls_totals-articles exp = 130 ).
    cl_abap_unit_assert=>assert_equals( act = ls_totals-quantity exp = 400 ).
    cl_abap_unit_assert=>assert_equals( act = ls_totals-gross    exp = '5567.100' ).
    cl_abap_unit_assert=>assert_equals( act = ls_totals-net      exp = '5537.200' ).

  ENDMETHOD.


  METHOD total_text_single_type.

    DATA(lv_text) = zcl_sd_epack_rules=>total_text(
                      zcl_sd_epack_rules=>calc_totals( items( nabil_lines( ) ) ) ).

    cl_abap_unit_assert=>assert_equals( act = lv_text exp = `Total 130 Bundles` ).

  ENDMETHOD.


  METHOD total_text_two_types.

    DATA(lv_text) = zcl_sd_epack_rules=>total_text(
                      zcl_sd_epack_rules=>calc_totals( items( tarmac_lines( ) ) ) ).

    cl_abap_unit_assert=>assert_equals( act = lv_text
                                        exp = `Total 513 Loose Pipes + 24 Boxes = 537 Packages` ).

  ENDMETHOD.


  METHOD total_text_no_type.

    DATA(lt_lines) = nabil_lines( ).
    CLEAR lt_lines[ 1 ]-pkg_type.

    DATA(lv_text) = zcl_sd_epack_rules=>total_text(
                      zcl_sd_epack_rules=>calc_totals( items( lt_lines ) ) ).

    cl_abap_unit_assert=>assert_equals( act = lv_text exp = `Total 130 Packages` ).

  ENDMETHOD.


  METHOD print_rows_heading_subtotal.

    DATA(lt_print) = zcl_sd_epack_rules=>build_print_items(
      it_items    = items( tarmac_lines( ) )
      it_headings = VALUE #( ( hsn = '3917 2390' text = `PP LOW NOISE PIPES (PP PIPES) - HS CODE : 3917 2390` )
                             ( hsn = '3917 4000' text = `PP LOW NOISE FITTINGS (PP FITTINGS) - HS CODE : 3917 4000` ) ) ).

    " H, I, I, I, S(513 Loose Pipes), H, I, S(24 Boxes)
    cl_abap_unit_assert=>assert_equals( act = lines( lt_print ) exp = 8 ).
    cl_abap_unit_assert=>assert_equals( act = lt_print[ 1 ]-row_type exp = zif_sd_epack=>gc_row_type-heading ).
    cl_abap_unit_assert=>assert_equals( act = lt_print[ 5 ]-row_type exp = zif_sd_epack=>gc_row_type-subtotal ).
    cl_abap_unit_assert=>assert_equals( act = lt_print[ 5 ]-art_txt  exp = '513' ).
    cl_abap_unit_assert=>assert_equals( act = lt_print[ 5 ]-row_text exp = '<=Loose Pipes' ).
    cl_abap_unit_assert=>assert_equals( act = lt_print[ 6 ]-row_type exp = zif_sd_epack=>gc_row_type-heading ).
    cl_abap_unit_assert=>assert_equals( act = lt_print[ 8 ]-row_text exp = '<=Boxes' ).

  ENDMETHOD.


  METHOD print_rows_blank_in_group.

    DATA(lt_print) = zcl_sd_epack_rules=>build_print_items(
                       it_items    = items( nabil_lines( ) )
                       it_headings = VALUE #( ) ).

    " Row 1 = heading; rows 2/3 = first merged group
    cl_abap_unit_assert=>assert_equals( act = lt_print[ 2 ]-art_txt exp = '60' ).
    cl_abap_unit_assert=>assert_equals( act = lt_print[ 2 ]-gwt_txt exp = '2121.000' ).
    cl_abap_unit_assert=>assert_initial( lt_print[ 3 ]-art_txt ).
    cl_abap_unit_assert=>assert_initial( lt_print[ 3 ]-gwt_txt ).
    cl_abap_unit_assert=>assert_false( lt_print[ 2 ]-wt_last ).
    cl_abap_unit_assert=>assert_true( lt_print[ 3 ]-wt_last ).
    " Only one package type -> no subtotal rows
    cl_abap_unit_assert=>assert_false( xsdbool( line_exists(
      lt_print[ row_type = zif_sd_epack=>gc_row_type-subtotal ] ) ) ).

  ENDMETHOD.


  METHOD hsn_is_formatted.
    cl_abap_unit_assert=>assert_equals( act = zcl_sd_epack_rules=>format_hsn( '39172390' )  exp = '3917 2390' ).
    cl_abap_unit_assert=>assert_equals( act = zcl_sd_epack_rules=>format_hsn( '3917 2390' ) exp = '3917 2390' ).
    cl_abap_unit_assert=>assert_equals( act = zcl_sd_epack_rules=>format_hsn( '3917' )      exp = '3917' ).
  ENDMETHOD.


  METHOD heading_from_goods_text.

    DATA(lv_text) = zcl_sd_epack_rules=>heading_for_hsn(
      iv_hsn         = CONV zsd_hsn( '3917 4000' )
      it_goods_lines = VALUE #( ( `PP LOW NOISE PIPES (PP PIPES) - HS CODE : 3917 2390` )
                                ( `PP LOW NOISE FITTINGS (PP FITTINGS) - HS CODE : 3917 4000` ) ) ).

    cl_abap_unit_assert=>assert_equals( act = lv_text
                                        exp = `PP LOW NOISE FITTINGS (PP FITTINGS) - HS CODE : 3917 4000` ).

  ENDMETHOD.


  METHOD heading_fallback_t604n.

    DATA(lv_text) = zcl_sd_epack_rules=>heading_for_hsn( iv_hsn         = CONV zsd_hsn( '39172390' )
                                                         it_goods_lines = VALUE #( )
                                                         iv_hsn_text    = `Tubes, pipes and hoses` ).

    cl_abap_unit_assert=>assert_equals( act = lv_text exp = `Tubes, pipes and hoses - HS CODE : 3917 2390` ).

  ENDMETHOD.


  METHOD weight_three_decimals.
    cl_abap_unit_assert=>assert_equals( act = zcl_sd_epack_rules=>format_weight( CONV zsd_gwt( '115386.78' ) ) exp = `115386.780` ).
    cl_abap_unit_assert=>assert_equals( act = zcl_sd_epack_rules=>format_quantity( CONV zsd_qty( 12146 ) )     exp = `12146` ).
    cl_abap_unit_assert=>assert_equals( act = zcl_sd_epack_rules=>format_date( CONV d( '20250527' ) )   exp = `27-05-2025` ).
  ENDMETHOD.


  METHOD date_formats_are_parsed.

    DATA lv_date TYPE d.

    cl_abap_unit_assert=>assert_true( zcl_sd_epack_rules=>try_parse_date( EXPORTING iv_text = `45804`
                                                                          IMPORTING ev_date = lv_date ) ).
    cl_abap_unit_assert=>assert_equals( act = lv_date exp = '20250527' ).

    zcl_sd_epack_rules=>try_parse_date( EXPORTING iv_text = `8.10.2025` IMPORTING ev_date = lv_date ).
    cl_abap_unit_assert=>assert_equals( act = lv_date exp = '20251008' ).

    zcl_sd_epack_rules=>try_parse_date( EXPORTING iv_text = `27-05-2025` IMPORTING ev_date = lv_date ).
    cl_abap_unit_assert=>assert_equals( act = lv_date exp = '20250527' ).

    zcl_sd_epack_rules=>try_parse_date( EXPORTING iv_text = `2025-08-18` IMPORTING ev_date = lv_date ).
    cl_abap_unit_assert=>assert_equals( act = lv_date exp = '20250818' ).

    zcl_sd_epack_rules=>try_parse_date( EXPORTING iv_text = `20250818` IMPORTING ev_date = lv_date ).
    cl_abap_unit_assert=>assert_equals( act = lv_date exp = '20250818' ).

  ENDMETHOD.


  METHOD invalid_date_is_rejected.
    DATA lv_date TYPE d.
    cl_abap_unit_assert=>assert_false( zcl_sd_epack_rules=>try_parse_date( EXPORTING iv_text = `31.02.2025`
                                                                           IMPORTING ev_date = lv_date ) ).
    cl_abap_unit_assert=>assert_false( zcl_sd_epack_rules=>try_parse_date( EXPORTING iv_text = `next week`
                                                                           IMPORTING ev_date = lv_date ) ).
  ENDMETHOD.


  METHOD numbers_are_parsed.

    DATA lv_value TYPE decfloat34.

    cl_abap_unit_assert=>assert_true( zcl_sd_epack_rules=>try_parse_number( EXPORTING iv_text  = `2107.2000000000003`
                                                                            IMPORTING ev_value = lv_value ) ).
    cl_abap_unit_assert=>assert_equals( act = round( val = lv_value dec = 3 ) exp = CONV decfloat34( '2107.2' ) ).

    zcl_sd_epack_rules=>try_parse_number( EXPORTING iv_text = `1,234.5` IMPORTING ev_value = lv_value ).
    cl_abap_unit_assert=>assert_equals( act = lv_value exp = CONV decfloat34( '1234.5' ) ).

    cl_abap_unit_assert=>assert_false( zcl_sd_epack_rules=>try_parse_number( EXPORTING iv_text  = `12 kg`
                                                                             IMPORTING ev_value = lv_value ) ).

  ENDMETHOD.


  METHOD invoice_list_is_split.

    zcl_sd_epack_rules=>split_invoices( EXPORTING iv_text    = `90000014, 90000015 & 90000020 / ABC`
                                        IMPORTING et_vbeln   = DATA(lt_vbeln)
                                                  et_invalid = DATA(lt_invalid) ).

    cl_abap_unit_assert=>assert_equals( act = lines( lt_vbeln ) exp = 3 ).
    cl_abap_unit_assert=>assert_true( xsdbool( line_exists( lt_vbeln[ table_line = '0090000015' ] ) ) ).
    cl_abap_unit_assert=>assert_equals( act = lt_invalid exp = VALUE string_table( ( `ABC` ) ) ).

  ENDMETHOD.


  METHOD multi_line_cell_is_split.

    DATA(lt_lines) = zcl_sd_epack_rules=>split_lines(
      |CSNU8464923-1X40' HC{ cl_abap_char_utilities=>cr_lf }TCLU1466959-1X40' HC{ cl_abap_char_utilities=>newline }| ).

    cl_abap_unit_assert=>assert_equals( act = lt_lines
                                        exp = VALUE zif_sd_epack=>tt_text_line( ( `CSNU8464923-1X40' HC` )
                                                                                ( `TCLU1466959-1X40' HC` ) ) ).

  ENDMETHOD.

ENDCLASS.
