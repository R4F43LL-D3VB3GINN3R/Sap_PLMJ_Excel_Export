*&---------------------------------------------------------------------*
*& Include          ZPS_RPT_HORAS_DESP_V4_XLS
*&---------------------------------------------------------------------*

CLASS lcl_excel_form DEFINITION.
  PUBLIC SECTION.

    METHODS get_excel_formulary.

  PROTECTED SECTION.
  PRIVATE SECTION.

    METHODS get_data.
    METHODS set_style.
    METHODS set_sheets.
    METHODS set_sheet_title.
    METHODS download_xls.

    "table to get data from report
    DATA: gt_data TYPE TABLE OF ty_s_report,
          gs_data TYPE ty_s_report.

    DATA: "objects and variables to excel file construction
      o_xl           TYPE REF TO zcl_excel,
      lo_worksheet   TYPE REF TO zcl_excel_worksheet,
      lo_column      TYPE REF TO zcl_excel_worksheet_columndime,
      lo_style       TYPE REF TO zcl_excel_style,
      o_border_dark  TYPE REF TO zcl_excel_style_border,
      o_border_light TYPE REF TO zcl_excel_style_border.

    DATA: "cell styles
      tp_style_bold_center_guid  TYPE zexcel_cell_style,
      tp_style_bold_center_guid2 TYPE zexcel_cell_style,
      tp_style_bold_center_guid3 TYPE zexcel_cell_style,
      tp_style_bold_center_guid4 TYPE zexcel_cell_style,
      tp_style_bold_center_guid5 TYPE zexcel_cell_style,
      tp_style_bold_center_guid6 TYPE zexcel_cell_style,
      tp_style_bold_center_guid7 TYPE zexcel_cell_style.

    DATA: "excel file variables
      gv_title    TYPE zexcel_sheet_title,
      lt_bin_data TYPE TABLE OF x255,
      lv_xstr     TYPE xstring.

ENDCLASS.

CLASS lcl_excel_form IMPLEMENTATION.

  "--------------------------------------------------------------------------------------------
  " DATA       : 12.08.2025
  " CREATED BY : RLA@SBX
  "
  " FUNCTIONALITY:
  "   - fills excel file with data
  "--------------------------------------------------------------------------------------------

  METHOD set_sheets.

    set_sheet_title( ).

    DATA: lv_old_assunto TYPE char24.

    "counters sets with initial position.
    DATA: lv_countheader       TYPE i VALUE 1,
          lv_count_tableheader TYPE i VALUE 6,
          lv_count_tablelines  TYPE i VALUE 7.

    "hours and amounts aggregations
    DATA: lv_total_hours  TYPE p DECIMALS 2,
          lv_total_amount TYPE p DECIMALS 2.

    "hours and amounts aggregation for footer (total sums)
    DATA: lv_total_hours_footer  TYPE p DECIMALS 2,
          lv_total_amount_footer TYPE p DECIMALS 2.

    "max lens for columns width
    DATA: lv_maxlen_col_a_float TYPE float,
          lv_maxlen_col_a_int   TYPE i.

    SORT gt_data BY pep_mask ASCENDING.
    LOOP AT gt_data INTO gs_data.

      IF lv_old_assunto = gs_data-pep_mask.
        CONTINUE.
      ENDIF.

      "--------------------------------------------------------------------------------------------------------------------------------------------------
      "                                                      HEADER
      "--------------------------------------------------------------------------------------------------------------------------------------------------
      lo_worksheet->set_cell( ip_row = lv_countheader ip_column = 'A' ip_value = gs_data-type_txt  ip_style = tp_style_bold_center_guid2 ).
      ADD 1 TO lv_countheader.
      lo_worksheet->set_cell( ip_row = lv_countheader ip_column = 'A' ip_value = gs_data-kunnr     ip_style = tp_style_bold_center_guid2 ).
      ADD 1 TO lv_countheader.
      lo_worksheet->set_cell( ip_row = lv_countheader ip_column = 'A' ip_value = gs_data-prart_txt ip_style = tp_style_bold_center_guid2 ).
      ADD 1 TO lv_countheader.
      lo_worksheet->set_cell( ip_row = lv_countheader ip_column = 'A' ip_value = gs_data-pep_mask  ip_style = tp_style_bold_center_guid2 ).

      "get length for column A
      IF lv_maxlen_col_a_int LT strlen( gs_data-type_txt ).
        lv_maxlen_col_a_int = strlen( gs_data-type_txt ).
      ENDIF.

      "--------------------------------------------------------------------------------------------------------------------------------------------------
      "                                                       TABLE HEADER
      "--------------------------------------------------------------------------------------------------------------------------------------------------
      lo_worksheet->set_cell( ip_row = lv_count_tableheader ip_column = 'A' ip_value = 'Data'           ip_style = tp_style_bold_center_guid ).
      lo_worksheet->set_cell( ip_row = lv_count_tableheader ip_column = 'B' ip_value = 'Discriminativo' ip_style = tp_style_bold_center_guid ).
      lo_worksheet->set_cell( ip_row = lv_count_tableheader ip_column = 'C' ip_value = 'Horas'          ip_style = tp_style_bold_center_guid ).
      lo_worksheet->set_cell( ip_row = lv_count_tableheader ip_column = 'D' ip_value = 'Preço/Hora'     ip_style = tp_style_bold_center_guid ).
      lo_worksheet->set_cell( ip_row = lv_count_tableheader ip_column = 'E' ip_value = 'Montante'       ip_style = tp_style_bold_center_guid ).
      lo_worksheet->set_cell( ip_row = lv_count_tableheader ip_column = 'F' ip_value = 'Nome'           ip_style = tp_style_bold_center_guid ).

      "--------------------------------------------------------------------------------------------------------------------------------------------------
      "                                                       TABLE LINES
      "--------------------------------------------------------------------------------------------------------------------------------------------------
      LOOP AT gt_data INTO gs_data WHERE pep_mask = gs_data-pep_mask.

        lo_worksheet->set_cell( ip_row = lv_count_tablelines ip_column = 'A' ip_value = gs_data-data               ip_style = tp_style_bold_center_guid3 ).
        lo_worksheet->set_cell( ip_row = lv_count_tablelines ip_column = 'B' ip_value = gs_data-descriminativo_str ip_style = tp_style_bold_center_guid6 ).
        lo_worksheet->set_cell( ip_row = lv_count_tablelines ip_column = 'C' ip_value = gs_data-horas              ip_style = tp_style_bold_center_guid4 ).
        lo_worksheet->set_cell( ip_row = lv_count_tablelines ip_column = 'D' ip_value = gs_data-horas_advogado     ip_style = tp_style_bold_center_guid4 ).
        lo_worksheet->set_cell( ip_row = lv_count_tablelines ip_column = 'E' ip_value = gs_data-montante           ip_style = tp_style_bold_center_guid4 ).
        lo_worksheet->set_cell( ip_row = lv_count_tablelines ip_column = 'F' ip_value = gs_data-rufnm              ip_style = tp_style_bold_center_guid2 ).
        ADD 1 TO lv_count_tablelines.

        "sum hours and amount to pay
        lv_total_hours  = lv_total_hours  + gs_data-horas.
        lv_total_amount = lv_total_amount + gs_data-montante.

        "sum hours and amount to pay for footer
        lv_total_hours_footer  = lv_total_hours_footer  + gs_data-horas.
        lv_total_amount_footer = lv_total_amount_footer + gs_data-montante.

      ENDLOOP.

      "--------------------------------------------------------------------------------------------------------------------------------------------------
      "                                                        AMOUNT TO PAY HOURS / TOTAL
      "--------------------------------------------------------------------------------------------------------------------------------------------------
      lo_worksheet->set_cell( ip_row = lv_count_tablelines ip_column = 'C' ip_value = lv_total_hours  ip_style = tp_style_bold_center_guid5 ).
      lo_worksheet->set_cell( ip_row = lv_count_tablelines ip_column = 'E' ip_value = lv_total_amount ip_style = tp_style_bold_center_guid5 ).
      ADD 1 TO lv_count_tablelines.

      CLEAR: lv_total_hours, lv_total_amount.

      "--------------------------------------------------------------------------------------------------------------------------------------------------
      "                                                        ADJUST FOR COUNTERS
      "--------------------------------------------------------------------------------------------------------------------------------------------------
      lv_countheader       = lv_count_tablelines + 2.  "sets new position for header
      lv_count_tableheader = lv_countheader + 5.       "sets new position for table header
      lv_count_tablelines  = lv_count_tableheader + 1. "sets new position for table lines.
      "------------------------------------------------------------------------------------

      lv_old_assunto = gs_data-pep_mask.

    ENDLOOP.

    "--------------------------------------------------------------------------------------------------------------------------------------------------
    "                                                          TOTAL FOR FOOTER
    "--------------------------------------------------------------------------------------------------------------------------------------------------
    lo_worksheet->set_cell( ip_row = lv_count_tablelines ip_column = 'B' ip_value = 'TOTAL'                ip_style = tp_style_bold_center_guid7 ).
    lo_worksheet->set_cell( ip_row = lv_count_tablelines ip_column = 'C' ip_value = lv_total_hours_footer  ip_style = tp_style_bold_center_guid5 ).
    lo_worksheet->set_cell( ip_row = lv_count_tablelines ip_column = 'E' ip_value = lv_total_amount_footer ip_style = tp_style_bold_center_guid5 ).

    "--------------------------------------------------------------------------------------------------------------------------------------------------
    "                                                          ADJUST FOR COLUMNS
    "--------------------------------------------------------------------------------------------------------------------------------------------------

    "set width for col A -> Header
    lv_maxlen_col_a_float = lv_maxlen_col_a_int.
    lv_maxlen_col_a_float = lv_maxlen_col_a_float * '1.5'.
    lo_column = lo_worksheet->get_column_dimension( ip_column = 'A' ).
    lo_column->set_width( ip_width = lv_maxlen_col_a_float ).

    "set width for col B -> Decritive Text
    lo_column = lo_worksheet->get_column_dimension( ip_column = 'B' ).
    lo_column->set_width( ip_width = 100 ).

    "set width for col D -> Price/Hours
    lo_column = lo_worksheet->get_column_dimension( ip_column = 'D' ).
    lo_column->set_width( ip_width = 15 ).

    "set width for col E -> Amount
    lo_column = lo_worksheet->get_column_dimension( ip_column = 'E' ).
    lo_column->set_width( ip_width = 15 ).

    "set width for col F -> Lawyer's Name
    lo_column = lo_worksheet->get_column_dimension( ip_column = 'F' ).
    lo_column->set_width( ip_width = 30 ).

  ENDMETHOD.

  METHOD get_data.

    "--------------------------------------------------------------------------------------------
    " DATA       : 12.08.2025
    " CREATED BY : RLA@SBX
    "
    " FUNCTIONALITY:
    "   - receives internal table from report
    "--------------------------------------------------------------------------------------------

    CREATE OBJECT o_xl.

    gt_data = gt_dados.

  ENDMETHOD.

  METHOD get_excel_formulary.

    "--------------------------------------------------------------------------------------------
    " DATA       : 12.08.2025
    " CREATED BY : RLA@SBX
    "
    " FUNCTIONALITY:
    "   - send excel file to report
    "--------------------------------------------------------------------------------------------

    get_data( ).
    set_style( ).
    set_sheets( ).
    download_xls( ).

  ENDMETHOD.

  METHOD set_sheet_title.

    "--------------------------------------------------------------------------------------------
    " DATA       : 12.08.2025
    " CREATED BY : RLA@SBX
    "
    " FUNCTIONALITY:
    "   - set title to worksheet
    "--------------------------------------------------------------------------------------------

    DATA(o_xl_ws) = o_xl->get_active_worksheet( ).
    lo_worksheet = o_xl_ws.

    gv_title = | { sy-datum } |.
    lo_worksheet->set_title( ip_title = gv_title ).

  ENDMETHOD.

  METHOD download_xls.

    "--------------------------------------------------------------------------------------------
    " DATA       : 12.08.2025
    " CREATED BY : RLA@SBX
    "
    " FUNCTIONALITY:
    "   - donwnload excel file to pc local
    "--------------------------------------------------------------------------------------------

    DATA(o_xlwriter)  = CAST zif_excel_writer( NEW zcl_excel_writer_2007( ) ).
    DATA(lv_xl_xdata) = o_xlwriter->write_file( o_xl ).
    DATA(it_raw_data) = cl_bcs_convert=>xstring_to_solix( EXPORTING iv_xstring = lv_xl_xdata ).

    DATA: lv_filename TYPE string.
    lv_filename = |Relatorios_Hora_Despesas_{ sy-datum }|.
    DATA: lv_path TYPE string VALUE 'C:\temp\'.
    DATA: lv_fullpath TYPE string.
    CONCATENATE lv_path lv_filename INTO lv_fullpath.

    CALL METHOD cl_gui_frontend_services=>file_save_dialog
      EXPORTING
        default_extension = 'xlsx'
        default_file_name = lv_filename
        file_filter       = 'Excel (*.xlsx)|*.xlsx|All Files (*.*)|*.*'
      CHANGING
        filename          = lv_filename
        path              = lv_path
        fullpath          = lv_fullpath
      EXCEPTIONS
        OTHERS            = 1.

    cl_gui_frontend_services=>gui_download(
      EXPORTING
        filename     = lv_fullpath
        filetype     = 'BIN'
        bin_filesize = xstrlen( lv_xl_xdata )
      CHANGING
        data_tab     = it_raw_data
    ).

  ENDMETHOD.

  METHOD set_style.

    "--------------------------------------------------------------------------------------------
    " DATA       : 12.08.2025
    " CREATED BY : RLA@SBX
    "
    " FUNCTIONALITY:
    "   - set cell styles for all needs
    "--------------------------------------------------------------------------------------------

    "cria objetos das bordas
    CREATE OBJECT o_border_dark.
    o_border_dark->border_color-rgb  = zcl_excel_style_color=>c_black.
    o_border_dark->border_style      = zcl_excel_style_border=>c_border_thin.
    CREATE OBJECT o_border_light.
    o_border_light->border_color-rgb = zcl_excel_style_color=>c_gray.
    o_border_light->border_style     = zcl_excel_style_border=>c_border_thin.

    "monta o primeiro estilo para a coluna A da paginacao
    CREATE OBJECT me->lo_style.

    "default style
    lo_style                             = o_xl->add_new_style( ).
    lo_style->font->bold                 = abap_true.
    lo_style->font->italic               = abap_false.
    lo_style->font->name                 = zcl_excel_style_font=>c_name_arial.
    lo_style->font->scheme               = zcl_excel_style_font=>c_scheme_none.
    lo_style->font->size                 = 10.
    lo_style->font->color-rgb            = zcl_excel_style_color=>c_black.
    lo_style->alignment->horizontal      = zcl_excel_style_alignment=>c_horizontal_left.
    lo_style->alignment->vertical        = zcl_excel_style_alignment=>c_vertical_top.
    lo_style->borders->allborders        = o_border_light.
    lo_style->fill->filltype             = zcl_excel_style_fill=>c_fill_solid.
    lo_style->fill->fgcolor-rgb          = zcl_excel_style_color=>c_gray.
    tp_style_bold_center_guid            = lo_style->get_guid( ). "nao esquecer

    "header table style
    lo_style                             = o_xl->add_new_style( ).
    lo_style->font->bold                 = abap_false.
    lo_style->font->italic               = abap_false.
    lo_style->font->name                 = zcl_excel_style_font=>c_name_arial.
    lo_style->font->scheme               = zcl_excel_style_font=>c_scheme_none.
    lo_style->font->size                 = 10.
    lo_style->font->color-rgb            = zcl_excel_style_color=>c_black.
    lo_style->alignment->horizontal      = zcl_excel_style_alignment=>c_horizontal_left.
    lo_style->alignment->vertical        = zcl_excel_style_alignment=>c_vertical_top.
    lo_style->borders->allborders        = o_border_dark.
    tp_style_bold_center_guid2           = lo_style->get_guid( ). "nao esquecer

    "dates style
    lo_style                             = o_xl->add_new_style( ).
    lo_style->font->bold                 = abap_false.
    lo_style->font->italic               = abap_false.
    lo_style->font->name                 = zcl_excel_style_font=>c_name_arial.
    lo_style->font->scheme               = zcl_excel_style_font=>c_scheme_none.
    lo_style->font->size                 = 10.
    lo_style->font->color-rgb            = zcl_excel_style_color=>c_black.
    lo_style->alignment->horizontal      = zcl_excel_style_alignment=>c_horizontal_right.
    lo_style->alignment->vertical        = zcl_excel_style_alignment=>c_vertical_top.
    lo_style->borders->allborders        = o_border_dark.
    lo_style->fill->filltype             = zcl_excel_style_fill=>c_fill_solid.
    lo_style->fill->fgcolor-rgb          = zcl_excel_style_color=>c_white.
    lo_style->number_format->format_code = 'DD/MM/YYYY'.
    tp_style_bold_center_guid3           = lo_style->get_guid( ). "nao esquecer

    "hours and amount style
    lo_style                             = o_xl->add_new_style( ).
    lo_style->font->bold                 = abap_false.
    lo_style->font->italic               = abap_false.
    lo_style->font->name                 = zcl_excel_style_font=>c_name_arial.
    lo_style->font->scheme               = zcl_excel_style_font=>c_scheme_none.
    lo_style->font->size                 = 10.
    lo_style->font->color-rgb            = zcl_excel_style_color=>c_black.
    lo_style->alignment->horizontal      = zcl_excel_style_alignment=>c_horizontal_right.
    lo_style->alignment->vertical        = zcl_excel_style_alignment=>c_vertical_top.
    lo_style->borders->allborders        = o_border_dark.
    lo_style->fill->filltype             = zcl_excel_style_fill=>c_fill_solid.
    lo_style->fill->fgcolor-rgb          = zcl_excel_style_color=>c_white.
    lo_style->number_format->format_code = '#,##0.00'.
    tp_style_bold_center_guid4           = lo_style->get_guid( ). "nao esquecer

    "total amount and hours style
    lo_style                             = o_xl->add_new_style( ).
    lo_style->font->bold                 = abap_true.
    lo_style->font->italic               = abap_false.
    lo_style->font->name                 = zcl_excel_style_font=>c_name_arial.
    lo_style->font->scheme               = zcl_excel_style_font=>c_scheme_none.
    lo_style->font->size                 = 10.
    lo_style->font->color-rgb            = zcl_excel_style_color=>c_black.
    lo_style->alignment->horizontal      = zcl_excel_style_alignment=>c_horizontal_right.
    lo_style->borders->allborders        = o_border_dark.
    lo_style->fill->filltype             = zcl_excel_style_fill=>c_fill_solid.
    lo_style->fill->fgcolor-rgb          = zcl_excel_style_color=>c_white.
    lo_style->number_format->format_code = '#,##0.00'.
    tp_style_bold_center_guid5           = lo_style->get_guid( ). "nao esquecer

    "discriminative style
    lo_style                             = o_xl->add_new_style( ).
    lo_style->font->bold                 = abap_false.
    lo_style->font->italic               = abap_false.
    lo_style->font->name                 = zcl_excel_style_font=>c_name_arial.
    lo_style->font->scheme               = zcl_excel_style_font=>c_scheme_none.
    lo_style->font->size                 = 10.
    lo_style->font->color-rgb            = zcl_excel_style_color=>c_black.
    lo_style->alignment->horizontal      = zcl_excel_style_alignment=>c_horizontal_left.
    lo_style->alignment->wraptext        = abap_true.
    lo_style->borders->allborders        = o_border_dark.
    lo_style->fill->filltype             = zcl_excel_style_fill=>c_fill_solid.
    lo_style->fill->fgcolor-rgb          = zcl_excel_style_color=>c_white.
    tp_style_bold_center_guid6           = lo_style->get_guid( ). "nao esquecer

    "total name style
    lo_style                             = o_xl->add_new_style( ).
    lo_style->font->bold                 = abap_true.
    lo_style->font->italic               = abap_false.
    lo_style->font->name                 = zcl_excel_style_font=>c_name_arial.
    lo_style->font->scheme               = zcl_excel_style_font=>c_scheme_none.
    lo_style->font->size                 = 10.
    lo_style->font->color-rgb            = zcl_excel_style_color=>c_black.
    lo_style->alignment->horizontal      = zcl_excel_style_alignment=>c_horizontal_right.
    lo_style->alignment->vertical        = zcl_excel_style_alignment=>c_vertical_top.
    tp_style_bold_center_guid7           = lo_style->get_guid( ). "nao esquecer

    "é possível montar vários estilos guid e usar da forma como convém

  ENDMETHOD.

ENDCLASS.
