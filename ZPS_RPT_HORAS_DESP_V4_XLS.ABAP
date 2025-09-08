*&---------------------------------------------------------------------*
*& Include          ZPS_RPT_HORAS_DESP_V4_XLS
*&---------------------------------------------------------------------*

*$*$-------------------------------------------------------------------*
*$*  MODIFICAÇÕES
*$---------------------------------------------------------------------*
*& Autor        : Rafael Albuquerque (SBX)
*& Data         : 13.08.2025
*& Referência   : INC0138977
*& Transporte   : S4DK908489
*& Ch. pesquisa : SBX-ID-EXCEL_FORM
*& Objetivo     : Exportação Arquivo Excel
*&--------------------------------------------------------------------

*$*$-------------------------------------------------------------------*
*$*  MODIFICAÇÕES
*$---------------------------------------------------------------------*
*& Autor        : Rafael Albuquerque (SBX)
*& Data         : 21.08.2025
*& Referência   : INC0138977
*& Transporte   : S4DK908489
*& Ch. pesquisa : RLA@SBX-ID-EXCEL_FORM
*& Objetivo     : - Inserir a denominação dos clientes e assuntos na coluna A
*&                - Inserir a denominação dos clientes e assuntos na coluna A
*&                - Remover bordas da Coluna A
*&                - Diminuir tamanho da Coluna A
*&--------------------------------------------------------------------

*$*$-------------------------------------------------------------------*
*$*  MODIFICAÇÕES
*$---------------------------------------------------------------------*
*& Autor        : Rafael Albuquerque (SBX)
*& Data         : 25.08.2025
*& Referência   : INC0138977
*& Transporte   : S4DK908489
*& Ch. pesquisa : RLA@SBX-ID-EXCEL_FORM
*& Objetivo     : - Novo Estilo para célula - Discriminativo alinhado em vertical baixo
*&                - Correção no espaçamento de linhas para o subtotal ao final do ficheiro
*&                - Alterações na tela de seleção e obrigatoriedade de preenchimento de campos
*&--------------------------------------------------------------------

*$*$-------------------------------------------------------------------*
*$*  MODIFICAÇÕES
*$---------------------------------------------------------------------*
*& Autor        : Rafael Albuquerque (SBX)
*& Data         : 25.08.2025
*& Referência   : INC0138977
*& Transporte   : S4DK908489
*& Ch. pesquisa : RLA@SBX-ID-EXCEL_FORM_003
*& Objetivo     : - Ao selecionar as opções que possuam despesas, no excel deve aparecer...
"                   após a soma dos honorários uma secção das despesas com o mesmo cabeçalho do honorário.
"                   Saltar 1 linha e criar a secção das despesas com a somas
*&                - No cabeçalho, ajustar de forma que apareça "Código + cliente e Código + assunto" na mesma linha, reduzindo de 4 linhas para 2
*&                - Retirar do excel a coluna "Preço/Hora(ADV)"
*&                - Incluir no excel, após o nome do advogado o código.
*&--------------------------------------------------------------------

CLASS lcl_excel_form DEFINITION.
  PUBLIC SECTION.

    METHODS get_excel_formulary.

  PROTECTED SECTION.
  PRIVATE SECTION.

    METHODS get_data.
    METHODS set_style.
    METHODS get_print_setup.
    METHODS set_sheets.
    METHODS set_sheet_title.
    METHODS download_xls.

    "table to get data from report
    DATA: gt_honorarios TYPE TABLE OF ty_s_report,
          gt_despesas   TYPE TABLE OF ty_s_report,
          gt_data       TYPE TABLE OF ty_s_report,
          gs_data       TYPE ty_s_report.

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
      tp_style_bold_center_guid7 TYPE zexcel_cell_style,
      tp_style_bold_center_guid8 TYPE zexcel_cell_style.

    DATA: "excel file variables
      gv_title    TYPE zexcel_sheet_title,
      lt_bin_data TYPE TABLE OF x255,
      lv_xstr     TYPE xstring.

    DATA: lv_flag_honorarios          TYPE abap_bool VALUE 'false',
          lv_flag_despesas            TYPE abap_bool VALUE 'false',
          lv_flag_honorarios_despesas TYPE abap_bool VALUE 'false'.

ENDCLASS.

CLASS lcl_excel_form IMPLEMENTATION.

  METHOD get_print_setup.

  "--------------------------------------------------------------------------------------------
  " DATA       : 08.09.2025
  " CREATED BY : RLA@SBX
  "
  " FUNCTIONALITY:
  "   - receives program print behavior
  "--------------------------------------------------------------------------------------------

    "reset flags
    lv_flag_honorarios          = abap_false.
    lv_flag_despesas            = abap_false.
    lv_flag_honorarios_despesas = abap_false.

    "verifies radiobuttom pressed
    IF p_hr EQ 'X'.
      lv_flag_honorarios          = p_hr.   "honorarios
    ELSEIF p_des EQ 'X'.
      lv_flag_despesas            = p_des.  "despesas
    ELSE.
      lv_flag_honorarios_despesas = p_tudo. "honorarios + despesas
    ENDIF.

  ENDMETHOD.

  "--------------------------------------------------------------------------------------------
  " DATA       : 12.08.2025
  " CREATED BY : RLA@SBX
  "
  " FUNCTIONALITY:
  "   - fills excel file with data
  "
  " UPDATES
  "   - 21.08.2025 :: New length for col A
  "   - 21.08.2025 :: New style for Header Info applied
  "   - 21.08.2025 :: Code fix to display customer and subject descritive
  "   - 21.08.2025 :: New column for (E preço/hora)
  "   - 21.08.2025 :: Update descritive column for (preço/hora) -> (preço/hora(adv))
  "   - 21.08.2025 :: Adjusts columns for total hours and subtotal
  "   - 25.08.2025 :: New style applied for discriminative
  "   - 25.08.2025 :: New row for subtotalS
  "   - 08.09.2025 :: New behavior for print (fees / expenses / fees + expenses)
  "   - 08.09.2025 :: Removed column hours lawyer
  "   - 08.09.2025 :: Includes customer code and subject code to related to their names
  "   - 08.09.2025 :: Removed lines switched to 2 lines concatenated
  "--------------------------------------------------------------------------------------------

  METHOD set_sheets.

    set_sheet_title( ).

    DATA: lv_old_assunto TYPE char24.

    "counters sets with initial position for honorarios.
    DATA: lv_countheader       TYPE i VALUE 2,
          lv_count_tableheader TYPE i VALUE 5,
          lv_count_tablelines  TYPE i VALUE 6.

    "counters sets with initial position for despesas.
    DATA: lv_countheader2       TYPE i VALUE 2,
          lv_count_tableheader2 TYPE i VALUE 5,
          lv_count_tablelines2  TYPE i VALUE 6.

    "hours and amounts aggregations for
    DATA: lv_total_hours  TYPE p DECIMALS 2,
          lv_total_amount TYPE p DECIMALS 2.

    "hours and amounts aggregation for footer (total sums)
    DATA: lv_total_hours_footer  TYPE p DECIMALS 2,
          lv_total_amount_footer TYPE p DECIMALS 2.

    "--------------------------------------------------------------------------------------------------------------------------------------------------
    "--------------------------------------------------------------------------------------------------------------------------------------------------
    "--------------------------------------------------------------------------------------------------------------------------------------------------
    "                                                           IMPRESSÃO DOS HONORÁRIOS
    "--------------------------------------------------------------------------------------------------------------------------------------------------
    "--------------------------------------------------------------------------------------------------------------------------------------------------
    "--------------------------------------------------------------------------------------------------------------------------------------------------
    IF gt_honorarios IS NOT INITIAL.

      LOOP AT gt_honorarios INTO gs_data.

        IF lv_old_assunto = gs_data-pep_mask.
          CONTINUE.
        ENDIF.

        "--------------------------------------------------------------------------------------------------------------------------------------------------
        "                                                      HEADER
        "--------------------------------------------------------------------------------------------------------------------------------------------------
        lo_worksheet->set_cell( ip_row = lv_countheader ip_column = 'A' ip_value = |{ gs_data-kunnr } { gs_data-name1 }|    ip_style = tp_style_bold_center_guid8 ).
        ADD 1 TO lv_countheader.
        lo_worksheet->set_cell( ip_row = lv_countheader ip_column = 'A' ip_value = |{ gs_data-pep_mask } { gs_data-post1 }| ip_style = tp_style_bold_center_guid8 ).

        "--------------------------------------------------------------------------------------------------------------------------------------------------
        "                                                       TABLE HEADER
        "--------------------------------------------------------------------------------------------------------------------------------------------------
        lo_worksheet->set_cell( ip_row = lv_count_tableheader ip_column = 'A' ip_value = 'Data'            ip_style = tp_style_bold_center_guid ).
        lo_worksheet->set_cell( ip_row = lv_count_tableheader ip_column = 'B' ip_value = 'Discriminativo'  ip_style = tp_style_bold_center_guid ).
        lo_worksheet->set_cell( ip_row = lv_count_tableheader ip_column = 'C' ip_value = 'Horas'           ip_style = tp_style_bold_center_guid ).
        lo_worksheet->set_cell( ip_row = lv_count_tableheader ip_column = 'D' ip_value = 'Preço/Hora'      ip_style = tp_style_bold_center_guid ).
        lo_worksheet->set_cell( ip_row = lv_count_tableheader ip_column = 'E' ip_value = 'Montante'        ip_style = tp_style_bold_center_guid ).
        lo_worksheet->set_cell( ip_row = lv_count_tableheader ip_column = 'F' ip_value = 'Nome'            ip_style = tp_style_bold_center_guid ).

        "--------------------------------------------------------------------------------------------------------------------------------------------------
        "                                                       TABLE LINES
        "--------------------------------------------------------------------------------------------------------------------------------------------------
        LOOP AT gt_honorarios INTO gs_data WHERE pep_mask = gs_data-pep_mask.

          lo_worksheet->set_cell( ip_row = lv_count_tablelines ip_column = 'A' ip_value = gs_data-data                          ip_style = tp_style_bold_center_guid3 ).
          lo_worksheet->set_cell( ip_row = lv_count_tablelines ip_column = 'B' ip_value = gs_data-descriminativo_str            ip_style = tp_style_bold_center_guid6 ).
          lo_worksheet->set_cell( ip_row = lv_count_tablelines ip_column = 'C' ip_value = gs_data-horas                         ip_style = tp_style_bold_center_guid4 ).
          lo_worksheet->set_cell( ip_row = lv_count_tablelines ip_column = 'D' ip_value = gs_data-pvp                           ip_style = tp_style_bold_center_guid4 ).
          lo_worksheet->set_cell( ip_row = lv_count_tablelines ip_column = 'E' ip_value = gs_data-montante                      ip_style = tp_style_bold_center_guid2 ).
          lo_worksheet->set_cell( ip_row = lv_count_tablelines ip_column = 'F' ip_value = |{ gs_data-rufnm } { gs_data-pernr }| ip_style = tp_style_bold_center_guid2 ).
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
        lv_count_tableheader = lv_countheader + 3.       "sets new position for table header
        lv_count_tablelines  = lv_count_tableheader.     "sets new position for table lines.
        "------------------------------------------------------------------------------------

        lv_old_assunto = gs_data-pep_mask.

      ENDLOOP.

      "--------------------------------------------------------------------------------------------------------------------------------------------------
      "                                                          TOTAL FOR FOOTER
      "--------------------------------------------------------------------------------------------------------------------------------------------------
      lv_count_tablelines = lv_count_tablelines - 3.

      lo_worksheet->set_cell( ip_row = lv_count_tablelines ip_column = 'B' ip_value = 'TOTAL HONORÁRIOS'     ip_style = tp_style_bold_center_guid7 ).
      lo_worksheet->set_cell( ip_row = lv_count_tablelines ip_column = 'C' ip_value = lv_total_hours_footer  ip_style = tp_style_bold_center_guid5 ).
      lo_worksheet->set_cell( ip_row = lv_count_tablelines ip_column = 'E' ip_value = lv_total_amount_footer ip_style = tp_style_bold_center_guid5 ).

      CLEAR gs_data.
      CLEAR lv_old_assunto.
      CLEAR lv_total_hours.
      CLEAR lv_total_amount.
      CLEAR lv_total_hours_footer.
      CLEAR lv_total_amount_footer.

    ENDIF.

    "in case of honorarios + despesas
    IF lv_flag_honorarios_despesas EQ 'X'.
      lv_countheader2       = lv_countheader + 3.
      lv_count_tableheader2 = lv_count_tableheader + 4.
      lv_count_tablelines2  = lv_count_tablelines  + 8.
    ENDIF.

    "--------------------------------------------------------------------------------------------------------------------------------------------------
    "--------------------------------------------------------------------------------------------------------------------------------------------------
    "--------------------------------------------------------------------------------------------------------------------------------------------------
    "                                                        IMPRESSÃO DAS DESPESAS
    "--------------------------------------------------------------------------------------------------------------------------------------------------
    "--------------------------------------------------------------------------------------------------------------------------------------------------
    "--------------------------------------------------------------------------------------------------------------------------------------------------

    IF gt_despesas IS NOT INITIAL.

      LOOP AT gt_despesas INTO gs_data.

        IF lv_old_assunto = gs_data-pep_mask.
          CONTINUE.
        ENDIF.

        "--------------------------------------------------------------------------------------------------------------------------------------------------
        "                                                      HEADER
        "--------------------------------------------------------------------------------------------------------------------------------------------------
        lo_worksheet->set_cell( ip_row = lv_countheader2 ip_column = 'A' ip_value = |{ gs_data-kunnr } { gs_data-name1 }|    ip_style = tp_style_bold_center_guid8 ).
        ADD 1 TO lv_countheader2.
        lo_worksheet->set_cell( ip_row = lv_countheader2 ip_column = 'A' ip_value = |{ gs_data-pep_mask } { gs_data-post1 }| ip_style = tp_style_bold_center_guid8 ).

        "--------------------------------------------------------------------------------------------------------------------------------------------------
        "                                                       TABLE HEADER
        "--------------------------------------------------------------------------------------------------------------------------------------------------
        lo_worksheet->set_cell( ip_row = lv_count_tableheader2 ip_column = 'A' ip_value = 'Data'            ip_style = tp_style_bold_center_guid ).
        lo_worksheet->set_cell( ip_row = lv_count_tableheader2 ip_column = 'B' ip_value = 'Discriminativo'  ip_style = tp_style_bold_center_guid ).
        lo_worksheet->set_cell( ip_row = lv_count_tableheader2 ip_column = 'C' ip_value = 'Horas'           ip_style = tp_style_bold_center_guid ).
        lo_worksheet->set_cell( ip_row = lv_count_tableheader2 ip_column = 'D' ip_value = 'Preço/Hora'      ip_style = tp_style_bold_center_guid ).
        lo_worksheet->set_cell( ip_row = lv_count_tableheader2 ip_column = 'E' ip_value = 'Montante'        ip_style = tp_style_bold_center_guid ).
        lo_worksheet->set_cell( ip_row = lv_count_tableheader2 ip_column = 'F' ip_value = 'Nome'            ip_style = tp_style_bold_center_guid ).

        "--------------------------------------------------------------------------------------------------------------------------------------------------
        "                                                       TABLE LINES
        "--------------------------------------------------------------------------------------------------------------------------------------------------
        LOOP AT gt_despesas INTO gs_data WHERE pep_mask = gs_data-pep_mask.

          lo_worksheet->set_cell( ip_row = lv_count_tablelines2 ip_column = 'A' ip_value = gs_data-data                          ip_style = tp_style_bold_center_guid3 ).
          lo_worksheet->set_cell( ip_row = lv_count_tablelines2 ip_column = 'B' ip_value = gs_data-descriminativo_str            ip_style = tp_style_bold_center_guid6 ).
          lo_worksheet->set_cell( ip_row = lv_count_tablelines2 ip_column = 'C' ip_value = gs_data-horas                         ip_style = tp_style_bold_center_guid4 ).
          lo_worksheet->set_cell( ip_row = lv_count_tablelines2 ip_column = 'D' ip_value = gs_data-pvp                           ip_style = tp_style_bold_center_guid4 ).
          lo_worksheet->set_cell( ip_row = lv_count_tablelines2 ip_column = 'E' ip_value = gs_data-montante                      ip_style = tp_style_bold_center_guid2 ).
          lo_worksheet->set_cell( ip_row = lv_count_tablelines2 ip_column = 'F' ip_value = |{ gs_data-rufnm } { gs_data-pernr }| ip_style = tp_style_bold_center_guid2 ).
          ADD 1 TO lv_count_tablelines2.

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
        lo_worksheet->set_cell( ip_row = lv_count_tablelines2 ip_column = 'C' ip_value = lv_total_hours  ip_style = tp_style_bold_center_guid5 ).
        lo_worksheet->set_cell( ip_row = lv_count_tablelines2 ip_column = 'E' ip_value = lv_total_amount ip_style = tp_style_bold_center_guid5 ).
        ADD 1 TO lv_count_tablelines2.

        CLEAR: lv_total_hours, lv_total_amount.

        "--------------------------------------------------------------------------------------------------------------------------------------------------
        "                                                        ADJUST FOR COUNTERS
        "--------------------------------------------------------------------------------------------------------------------------------------------------
        lv_countheader2       = lv_count_tablelines2 + 2.  "sets new position for header
        lv_count_tableheader2 = lv_countheader2 + 3.       "sets new position for table header
        lv_count_tablelines2  = lv_count_tableheader2.     "sets new position for table lines.
        "------------------------------------------------------------------------------------

        lv_old_assunto = gs_data-pep_mask.

      ENDLOOP.

      "--------------------------------------------------------------------------------------------------------------------------------------------------
      "                                                          TOTAL FOR FOOTER
      "--------------------------------------------------------------------------------------------------------------------------------------------------
      lv_count_tablelines2 = lv_count_tablelines2 - 3.

      lo_worksheet->set_cell( ip_row = lv_count_tablelines2 ip_column = 'B' ip_value = 'TOTAL DESPESAS'       ip_style = tp_style_bold_center_guid7 ).
      lo_worksheet->set_cell( ip_row = lv_count_tablelines2 ip_column = 'C' ip_value = lv_total_hours_footer  ip_style = tp_style_bold_center_guid5 ).
      lo_worksheet->set_cell( ip_row = lv_count_tablelines2 ip_column = 'E' ip_value = lv_total_amount_footer ip_style = tp_style_bold_center_guid5 ).

    ENDIF.


    "--------------------------------------------------------------------------------------------------------------------------------------------------
    "                                                          ADJUST FOR COLUMNS
    "--------------------------------------------------------------------------------------------------------------------------------------------------

    "set width for col A -> Header Info
    lo_column = lo_worksheet->get_column_dimension( ip_column = 'A' ).
    lo_column->set_width( ip_width = 10 ).

    "set width for col B -> Decritive Text
    lo_column = lo_worksheet->get_column_dimension( ip_column = 'B' ).
    lo_column->set_width( ip_width = 100 ).

    "set width for col D -> Price/Hours
    lo_column = lo_worksheet->get_column_dimension( ip_column = 'D' ).
    lo_column->set_width( ip_width = 20 ).

    "set width for col E -> Amount
    lo_column = lo_worksheet->get_column_dimension( ip_column = 'E' ).
    lo_column->set_width( ip_width = 20 ).

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
    "
    " UPDATES
    "   - 08.09.2025 :: new way of distributing data
    "--------------------------------------------------------------------------------------------

    CREATE OBJECT o_xl.

    get_print_setup( ).

    IF lv_flag_honorarios EQ 'X'.
      LOOP AT gt_dados INTO DATA(ls_honorarios) WHERE lstar EQ 'A001'.
        MOVE-CORRESPONDING ls_honorarios TO gs_data.
        APPEND gs_data TO gt_honorarios.
      ENDLOOP.
      SORT gt_honorarios BY pep_mask ASCENDING.
    ELSEIF lv_flag_despesas EQ 'X'.
      LOOP AT gt_dados INTO DATA(ls_despesas) WHERE lstar NE 'A001'.
        MOVE-CORRESPONDING ls_despesas TO gs_data.
        APPEND gs_data TO gt_despesas.
      ENDLOOP.
    ELSE.
      LOOP AT gt_dados INTO DATA(ls_honorarios2) WHERE lstar EQ 'A001'.
        MOVE-CORRESPONDING ls_honorarios2 TO gs_data.
        APPEND gs_data TO gt_honorarios.
      ENDLOOP.
      SORT gt_honorarios BY pep_mask ASCENDING.
      LOOP AT gt_dados INTO DATA(ls_despesas2) WHERE lstar NE 'A001'.
        MOVE-CORRESPONDING ls_despesas2 TO gs_data.
        APPEND gs_data TO gt_despesas.
      ENDLOOP.
    ENDIF.

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
    "
    " UPDATES
    "   - 21.08.2025 :: New Style tp_style_bold_center_guid8
    "   - 25.08.2025 :: New Style tp_style_bold_center_guid9
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
*    lo_style->borders->allborders        = o_border_dark.
    tp_style_bold_center_guid8           = lo_style->get_guid( ). "nao esquecer

    "las column style
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
    lo_style->alignment->vertical        = zcl_excel_style_alignment=>c_vertical_bottom.
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
