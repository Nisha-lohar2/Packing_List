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
