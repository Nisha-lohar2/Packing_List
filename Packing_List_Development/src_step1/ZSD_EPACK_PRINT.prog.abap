*&---------------------------------------------------------------------*
*& Report  ZSD_EPACK_PRINT  -  STEP 1 only
*&---------------------------------------------------------------------*
*& WRICEF 102-B Export packing list
*& 1. Selection screen: Excel file
*& 2. Read the file in binary form
*& 3. Read the first worksheet of the Excel
*& 4. Data of the worksheet -> internal table, as it is
*&    (columns A, B, C ... of the sheet, all values as text)
*&---------------------------------------------------------------------*
REPORT zsd_epack_print.

*----------------------------------------------------------------------*
* Type of one upload row - columns of the Excel template, in order
*----------------------------------------------------------------------*
TYPES:
  BEGIN OF ty_upload,
    zsd_packno    TYPE zsd_vbeln,       " Packing list Number
    vbrk_vbeln    TYPE z_string,        " SAP Invoice Number (one or "A, B")
    zsd_sn        TYPE zserno,          " Sl.No.
    zsd_inv_sn    TYPE zserno,          " Sr.No. INV
    zsd_artno     TYPE zchar15,         " Article No.
    zsd_std       TYPE zsqno,           " STD
    zsd_art       TYPE zsqno,           " Articles
    zsd_partno    TYPE zsd_debezei40,   " Part No.
    zsd_matdesc   TYPE zsd_debezei40,   " Product Description (fetch from material master)
    zsd_qty       TYPE zqty,            " Quantity Pcs
    zsd_gwt       TYPE zde_grwt,        " Gross Wt. KGS
    zsd_nwt       TYPE zde_net_wt,      " Net Wt. KGS
    zsd_hsn       TYPE zna_segment,     " H.S Code
    zsd_packdt    TYPE zsd_date,        " Packing list date
    zsd_expref    TYPE zsd_uid,         " Exporter Ref
    zsd_party_ref TYPE zref,            " Buyer's Reference No. & Date
    zsd_ref_dt    TYPE zsd_date,        " Buyer's Reference date
    zsd_cnty_orgn TYPE zsd_uid,         " Country of origin of goods
    zsd_np1       TYPE zsd_np,          " Notifier Party1
    zsd_np2       TYPE zsd_np,          " Notifier Party2
    zsd_np3       TYPE zsd_np,          " Notifier Party3
    zsd_np4       TYPE zsd_np,          " Notifier Party4
    zsd_np5       TYPE zsd_np,          " Notifier Party5
    zsd_prec      TYPE zchar30,         " Pre-Carriage
    zsd_rec_pc    TYPE zchar30,         " Place of Receipt by Pre-carrier
    zsd_pol       TYPE zchar30,         " Port of Loading
    zsd_pod       TYPE zchar30,         " Port of Discharge
    zsd_pld       TYPE zchar30,         " Place of Delivery
    zsd_mark      TYPE zsd_mark,        " Marks & Nos.
    zsd_vsl_flt   TYPE zchar30,         " Vessel / Flight No.
    zsd_cont      TYPE zsd_cont,        " Container No
    zsd_nopack    TYPE zsd_nopack,      " No. of Packages
    zsd_hsn_desc  TYPE zsd_hsn_desc,    " Description of goods
    zsd_pl_decl   TYPE z_string,        " Declaration of Packing list
    zsd_adv_amt   TYPE zde_amt1,        " Advance amount received
    zsd_cmmt      TYPE zsd_cmmt,        " Commercial terms of invoice
    zsd_inv_decl  TYPE z_string,        " Declaration of Invoice
  END OF ty_upload,
  ty_upload_tab TYPE STANDARD TABLE OF ty_upload WITH EMPTY KEY.

*----------------------------------------------------------------------*
* Data
*----------------------------------------------------------------------*
DATA gt_bin     TYPE solix_tab.                      " file, binary
DATA gv_size    TYPE i.                              " file length
DATA gv_xstring TYPE xstring.                        " file, xstring
DATA go_excel   TYPE REF TO cl_fdt_xl_spreadsheet.   " Excel document
DATA gt_sheets  TYPE if_fdt_doc_spreadsheet=>t_worksheet_names.
DATA gv_sheet   TYPE string.                         " worksheet name
DATA gr_data    TYPE REF TO data.                    " worksheet data

FIELD-SYMBOLS <gt_excel> TYPE STANDARD TABLE.        " Excel data, as it is
DATA gt_upload  TYPE ty_upload_tab.                  " upload rows, typed

*----------------------------------------------------------------------*
* Selection screen
*----------------------------------------------------------------------*
SELECTION-SCREEN BEGIN OF BLOCK b1 WITH FRAME TITLE TEXT-b01.
PARAMETERS p_file TYPE rlgrap-filename LOWER CASE OBLIGATORY.
SELECTION-SCREEN END OF BLOCK b1.

*----------------------------------------------------------------------*
AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_file.
*----------------------------------------------------------------------*
  PERFORM f4_file CHANGING p_file.

*----------------------------------------------------------------------*
START-OF-SELECTION.
*----------------------------------------------------------------------*
  PERFORM read_file.
  PERFORM read_worksheet.
  PERFORM display.

*&---------------------------------------------------------------------*
*& 1. Excel file from the PC in binary form -> xstring
*&---------------------------------------------------------------------*
FORM read_file.

  DATA lv_filename TYPE string.

  lv_filename = p_file.

  cl_gui_frontend_services=>gui_upload(
    EXPORTING
      filename                = lv_filename
      filetype                = 'BIN'
    IMPORTING
      filelength              = gv_size
    CHANGING
      data_tab                = gt_bin
    EXCEPTIONS
      file_open_error         = 1
      file_read_error         = 2
      OTHERS                  = 3 ).
  IF sy-subrc <> 0.
    MESSAGE 'File could not be read' TYPE 'S' DISPLAY LIKE 'E'.
    LEAVE LIST-PROCESSING.
  ENDIF.

  " Binary table -> xstring
  gv_xstring = cl_bcs_convert=>solix_to_xstring( it_solix = gt_bin
                                                 iv_size  = gv_size ).

ENDFORM.

*&---------------------------------------------------------------------*
*& 2. Excel -> first worksheet -> internal table, as it is
*&---------------------------------------------------------------------*
FORM read_worksheet.

  DATA lv_filename TYPE string.

  lv_filename = p_file.

  TRY.
      CREATE OBJECT go_excel
        EXPORTING
          document_name = lv_filename
          xdocument     = gv_xstring.
    CATCH cx_fdt_excel_core.
      MESSAGE 'File is not a valid Excel (.xlsx) file' TYPE 'S' DISPLAY LIKE 'E'.
      LEAVE LIST-PROCESSING.
  ENDTRY.

  " Worksheets of the file - the first one is read
  go_excel->if_fdt_doc_spreadsheet~get_worksheet_names(
    IMPORTING
      worksheet_names = gt_sheets ).

  READ TABLE gt_sheets INTO gv_sheet INDEX 1.
  IF sy-subrc <> 0.
    MESSAGE 'The Excel file has no worksheet' TYPE 'S' DISPLAY LIKE 'E'.
    LEAVE LIST-PROCESSING.
  ENDIF.

  " Data of the worksheet -> internal table (one row per Excel row,
  " one column per Excel column, values as in the sheet)
  gr_data = go_excel->if_fdt_doc_spreadsheet~get_itab_from_worksheet( gv_sheet ).
  ASSIGN gr_data->* TO <gt_excel>.
  IF <gt_excel> IS NOT ASSIGNED OR <gt_excel> IS INITIAL.
    MESSAGE 'The worksheet is empty' TYPE 'S' DISPLAY LIKE 'E'.
    LEAVE LIST-PROCESSING.
  ENDIF.

ENDFORM.

*&---------------------------------------------------------------------*
*& 3. Show the internal table (to check what was read)
*&---------------------------------------------------------------------*
FORM display.

  DATA lo_alv TYPE REF TO cl_salv_table.

  TRY.
      cl_salv_table=>factory(
        IMPORTING
          r_salv_table = lo_alv
        CHANGING
          t_table      = <gt_excel> ).
      lo_alv->get_functions( )->set_all( abap_true ).
      lo_alv->display( ).
    CATCH cx_salv_msg.
      MESSAGE 'Data could not be displayed' TYPE 'S' DISPLAY LIKE 'E'.
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
*&   Text symbol    B01     Excel file
*&   Selection text P_FILE  Excel file (.xlsx)
*&---------------------------------------------------------------------*
