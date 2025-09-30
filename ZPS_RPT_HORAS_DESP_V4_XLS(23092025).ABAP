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

*$*$-------------------------------------------------------------------*
*$*  MODIFICAÇÕES
*$---------------------------------------------------------------------*
*& Autor        : Rafael Albuquerque (SBX)
*& Data         : 23.09.2025
*& Referência   : INC0138977
*& Transporte   : S4DK908489
*& Ch. pesquisa : RLA@SBX-ID-EXCEL_FORM_004
*& Objetivo     : - new excel layout customization
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
    METHODS print_honorarios_despesas .
    METHODS print_despesas.

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
      tp_style_bold_center_guid   TYPE zexcel_cell_style,
      tp_style_bold_center_guid2  TYPE zexcel_cell_style,
      tp_style_bold_center_guid3  TYPE zexcel_cell_style,
      tp_style_bold_center_guid4  TYPE zexcel_cell_style,
      tp_style_bold_center_guid5  TYPE zexcel_cell_style,
      tp_style_bold_center_guid6  TYPE zexcel_cell_style,
      tp_style_bold_center_guid7  TYPE zexcel_cell_style,
      tp_style_bold_center_guid8  TYPE zexcel_cell_style,
*      RLA@SBX-ID-EXCEL_FORM_004
      tp_style_bold_center_guid9  TYPE zexcel_cell_style,
      tp_style_bold_center_guid10 TYPE zexcel_cell_style.
*      RLA@SBX-ID-EXCEL_FORM_004

    DATA: "excel file variables
      gv_title    TYPE zexcel_sheet_title,
      lt_bin_data TYPE TABLE OF x255,
      lv_xstr     TYPE xstring.

    DATA: lv_flag_honorarios          TYPE abap_bool VALUE 'false',
          lv_flag_despesas            TYPE abap_bool VALUE 'false',
          lv_flag_honorarios_despesas TYPE abap_bool VALUE 'false'.

    TYPES: BEGIN OF ty_somatorios,
             assunto    TYPE string,
             horatotal  TYPE i,
             valortotal TYPE i,
           END OF ty_somatorios.

*    RLA@SBX-ID-EXCEL_FORM_004
    DATA: gt_somatorios TYPE TABLE OF ty_somatorios,
          gs_somatorios TYPE ty_somatorios.

    DATA: gv_assunto_pai   TYPE string,
          gv_assunto_filho TYPE string.
    DATA: gv_strlen_assunto_pai   TYPE i,
          gv_strlen_assunto_filho TYPE i.
*    RLA@SBX-ID-EXCEL_FORM_004

ENDCLASS.

CLASS lcl_excel_form IMPLEMENTATION.

  METHOD print_despesas.

    "--------------------------------------------------------------------------------------------
    " DATA       : 23.09.2025
    " CREATED BY : RLA@SBX
    "
    " FUNCTIONALITY:
    "   - print expenses
    "--------------------------------------------------------------------------------------------

*    RLA@SBX-ID-EXCEL_FORM_004
    set_sheet_title( ).

    DATA: lv_old_assunto            TYPE char24,
          lv_old_subassunto         TYPE char24,
          lv_old_despesa            TYPE char24,
          lv_old_subassunto_despesa TYPE char24.

    DATA: strlen_discriminativo TYPE i.
    DATA: lv_index_assuntos TYPE i VALUE 0.

    "counters sets with initial position for honorarios.
    DATA: lv_countlines TYPE i VALUE 2.

    "hours and amounts aggregations for
    DATA: lv_total_hours  TYPE p DECIMALS 2,
          lv_total_amount TYPE p DECIMALS 2.

    SORT gt_honorarios BY pep_mask.
    SORT gt_despesas   BY pep_mask.

    DATA: lv_tothoras  TYPE i,
          lv_totamount TYPE i.

    LOOP AT gt_despesas INTO DATA(gs_despesas3).

      IF lv_old_subassunto_despesa EQ gs_despesas3-pep_mask.
        CONTINUE.
      ENDIF.

      "Tabelas de Montantes por Projetos e Subprojetos
      "--------------------------------------------------------------------------------------------------------------------------------------------
      IF lv_index_assuntos EQ 0.
        gv_strlen_assunto_pai    = strlen( gs_despesas3-pep_mask ) - 1.
        gv_assunto_pai           = substring( val = gs_despesas3-pep_mask off = 0 len = gv_strlen_assunto_pai ).
        gs_somatorios-assunto    = gs_despesas3-pep_mask.
        gs_somatorios-horatotal  = lv_total_hours.
        gs_somatorios-valortotal = lv_total_amount.
        APPEND gs_somatorios TO gt_somatorios.
      ELSE.
        gv_strlen_assunto_filho  = strlen( gs_despesas3-pep_mask ) - 1.
        gv_assunto_filho         = substring( val = gs_despesas3-pep_mask off = 0 len = gv_strlen_assunto_filho ).
        IF gv_assunto_pai EQ gv_assunto_filho.
          gs_somatorios-assunto    = gs_despesas3-pep_mask.
          gs_somatorios-horatotal  = lv_total_hours.
          gs_somatorios-valortotal = lv_total_amount.
          APPEND gs_somatorios TO gt_somatorios.
          gv_assunto_pai = gv_assunto_filho.
        ELSE.
          IF lines( gt_somatorios ) > 1.
            lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'C' ip_value = 'Resumo Despesas' ip_style = tp_style_bold_center_guid ).
            lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'D' ip_value = 'Montante'         ip_style = tp_style_bold_center_guid ).
            ADD 1 TO lv_countlines.

            LOOP AT gt_somatorios INTO gs_somatorios.
              lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'C' ip_value = gs_somatorios-assunto    ip_style = tp_style_bold_center_guid10 ).
              lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'D' ip_value = gs_somatorios-valortotal ip_style = tp_style_bold_center_guid4 ).
              ADD 1 TO lv_countlines.

              lv_tothoras  = lv_tothoras  + gs_somatorios-horatotal.
              lv_totamount = lv_totamount + gs_somatorios-valortotal.
            ENDLOOP.
            lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'C' ip_value = 'Total'      ip_style = tp_style_bold_center_guid9 ).
            lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'D' ip_value = lv_totamount ip_style = tp_style_bold_center_guid5 ).

            ADD 2 TO lv_countlines.
            CLEAR gs_somatorios.
            CLEAR: lv_tothoras, lv_totamount.
          ENDIF.
          CLEAR gt_somatorios.
          gs_somatorios-assunto    = gs_data-pep_mask.
          gs_somatorios-horatotal  = lv_total_hours.
          gs_somatorios-valortotal = lv_total_amount.
          APPEND gs_somatorios TO gt_somatorios.
          gv_assunto_pai           = gv_assunto_filho.
        ENDIF.
      ENDIF.

      "DESPESAS
      "--------------------------------------------------------------------------------------------------------------------------------------------------------------
      lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'A' ip_value = |{ gs_despesas3-kunnr } { gs_despesas3-name1 }|    ip_style = tp_style_bold_center_guid8 ).
      ADD 1 TO lv_countlines.
      lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'A' ip_value = |{ gs_despesas3-pep_mask } { gs_despesas3-post1 }| ip_style = tp_style_bold_center_guid8 ).
      ADD 1 TO lv_countlines.

      lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'A' ip_value = 'Data'                 ip_style = tp_style_bold_center_guid ).
      lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'B' ip_value = 'Despesas'             ip_style = tp_style_bold_center_guid ).
      lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'C' ip_value = 'Designação'           ip_style = tp_style_bold_center_guid ).
      lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'D' ip_value = 'Relatório da Despesa' ip_style = tp_style_bold_center_guid ).
      lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'E' ip_value = 'Montante'             ip_style = tp_style_bold_center_guid ).
      lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'F' ip_value = 'Nome'                 ip_style = tp_style_bold_center_guid ).
      ADD 1 TO lv_countlines.

      LOOP AT gt_despesas INTO DATA(gs_despesas4) WHERE pep_mask EQ gs_despesas3-pep_mask.

        lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'A' ip_value = gs_despesas4-data               ip_style = tp_style_bold_center_guid3 ).
        lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'B' ip_value = gs_despesas4-descriminativo_str ip_style = tp_style_bold_center_guid6 ).
        lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'C' ip_value = gs_despesas4-type_txt           ip_style = tp_style_bold_center_guid4 ).
        lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'D' ip_value = gs_despesas4-relexp             ip_style = tp_style_bold_center_guid4 ).
        lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'E' ip_value = gs_despesas4-montante           ip_style = tp_style_bold_center_guid2 ).
        lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'F' ip_value = gs_despesas4-rufnm              ip_style = tp_style_bold_center_guid2 ).
        ADD 1 TO lv_countlines.

        "sum hours and amount to pay
        lv_total_hours  = lv_total_hours  + gs_despesas4-horas.
        lv_total_amount = lv_total_amount + gs_despesas4-montante.

      ENDLOOP.

      lv_old_subassunto_despesa = gs_despesas3-pep_mask.

      lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'D' ip_value = 'Totais'        ip_style = tp_style_bold_center_guid9 ).
      lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'E' ip_value = lv_total_amount ip_style = tp_style_bold_center_guid5 ).
      ADD 2 TO lv_countlines.
      ADD 1 TO lv_index_assuntos.

      "insert amounts into sums by projects
      LOOP AT gt_somatorios INTO gs_somatorios WHERE assunto EQ gs_despesas3-pep_mask.
        gs_somatorios-valortotal = lv_total_amount.
        MODIFY gt_somatorios FROM gs_somatorios.
      ENDLOOP.
      CLEAR: lv_total_hours, lv_total_amount.

    ENDLOOP.

    "Tabelas de Montantes por Projetos e Subprojetos
    "--------------------------------------------------------------------------------------------------------------------------------------------
    IF lines( gt_somatorios ) > 1.
      lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'C' ip_value = 'Resumo Discriminativo' ip_style = tp_style_bold_center_guid ).
      lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'D' ip_value = 'Montante'              ip_style = tp_style_bold_center_guid ).
      ADD 1 TO lv_countlines.

      LOOP AT gt_somatorios INTO gs_somatorios.
        lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'C' ip_value = gs_somatorios-assunto    ip_style = tp_style_bold_center_guid10 ).
        lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'D' ip_value = gs_somatorios-valortotal ip_style = tp_style_bold_center_guid4 ).
        ADD 1 TO lv_countlines.

        lv_tothoras  = lv_tothoras  + gs_somatorios-horatotal.
        lv_totamount = lv_totamount + gs_somatorios-valortotal.
      ENDLOOP.
      lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'C' ip_value = 'Total'      ip_style = tp_style_bold_center_guid9 ).
      lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'D' ip_value = lv_totamount ip_style = tp_style_bold_center_guid5 ).

      ADD 2 TO lv_countlines.
    ENDIF.

    CLEAR gs_somatorios.
    CLEAR gt_somatorios.
    CLEAR: lv_tothoras, lv_totamount.
    CLEAR lv_index_assuntos.

    "COLUMNS
    "-----------------------------------------------------------------
    lo_column = lo_worksheet->get_column_dimension( ip_column = 'A' ).
    lo_column->set_width( ip_width = 10 ).

    lo_column = lo_worksheet->get_column_dimension( ip_column = 'B' ).
    lo_column->set_width( ip_width = 100 ).

    lo_column = lo_worksheet->get_column_dimension( ip_column = 'C' ).
    lo_column->set_width( ip_width = 28 ).

    lo_column->set_auto_size( ip_auto_size = abap_false ).
    lo_column = lo_worksheet->get_column_dimension( ip_column = 'D' ).
    lo_column->set_width( ip_width = 12 ).

    lo_column->set_auto_size( ip_auto_size = abap_false ).
    lo_column = lo_worksheet->get_column_dimension( ip_column = 'E' ).
    lo_column->set_width( ip_width = 12 ).

    lo_column->set_auto_size( ip_auto_size = abap_false ).
    lo_column = lo_worksheet->get_column_dimension( ip_column = 'F' ).
    lo_column->set_width( ip_width = 20 ).
*    RLA@SBX-ID-EXCEL_FORM_004

  ENDMETHOD.

  METHOD print_honorarios_despesas.

    "--------------------------------------------------------------------------------------------
    " DATA       : 23.09.2025
    " CREATED BY : RLA@SBX
    "
    " FUNCTIONALITY:
    "   - print honoraries and expenses
    "--------------------------------------------------------------------------------------------

*    RLA@SBX-ID-EXCEL_FORM_004
    set_sheet_title( ).

    DATA: lv_old_assunto            TYPE char24,
          lv_old_subassunto         TYPE char24,
          lv_old_despesa            TYPE char24,
          lv_old_subassunto_despesa TYPE char24.

    DATA: strlen_discriminativo TYPE i.
    DATA: lv_index_assuntos TYPE i VALUE 0.

    "counters sets with initial position for honorarios.
    DATA: lv_countlines TYPE i VALUE 2.

    "hours and amounts aggregations for
    DATA: lv_total_hours  TYPE p DECIMALS 2,
          lv_total_amount TYPE p DECIMALS 2.

    SORT gt_honorarios BY pep_mask.
    SORT gt_despesas   BY pep_mask.

    DATA: lv_tothoras  TYPE i,
          lv_totamount TYPE i.

    "HONORARIOS
    "--------------------------------------------------------------------------------------------------------------------------------------------------
    LOOP AT gt_honorarios INTO gs_data.

      IF lv_old_assunto EQ gs_data-pep_mask.
        CONTINUE.
      ENDIF.

      "Tabelas de Montantes por Projetos e Subprojetos
      "--------------------------------------------------------------------------------------------------------------------------------------------
      IF lv_index_assuntos EQ 0.
        gv_strlen_assunto_pai    = strlen( gs_data-pep_mask ) - 1.
        gv_assunto_pai           = substring( val = gs_data-pep_mask off = 0 len = gv_strlen_assunto_pai ).
        gs_somatorios-assunto    = gs_data-pep_mask.
        gs_somatorios-horatotal  = lv_total_hours.
        gs_somatorios-valortotal = lv_total_amount.
        APPEND gs_somatorios TO gt_somatorios.
      ELSE.
        gv_strlen_assunto_filho  = strlen( gs_data-pep_mask ) - 1.
        gv_assunto_filho         = substring( val = gs_data-pep_mask off = 0 len = gv_strlen_assunto_filho ).
        IF gv_assunto_pai EQ gv_assunto_filho.
          gs_somatorios-assunto    = gs_data-pep_mask.
          gs_somatorios-horatotal  = lv_total_hours.
          gs_somatorios-valortotal = lv_total_amount.
          APPEND gs_somatorios TO gt_somatorios.
          gv_assunto_pai = gv_assunto_filho.
        ELSE.
          IF lines( gt_somatorios ) > 1.
            lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'C' ip_value = 'Resumo Discriminativo' ip_style = tp_style_bold_center_guid ).
            lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'D' ip_value = 'Horas'                 ip_style = tp_style_bold_center_guid ).
            lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'E' ip_value = 'Montante'              ip_style = tp_style_bold_center_guid ).
            ADD 1 TO lv_countlines.

            LOOP AT gt_somatorios INTO gs_somatorios.
              lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'C' ip_value = gs_somatorios-assunto    ip_style = tp_style_bold_center_guid10 ).
              lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'D' ip_value = gs_somatorios-horatotal  ip_style = tp_style_bold_center_guid4 ).
              lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'E' ip_value = gs_somatorios-valortotal ip_style = tp_style_bold_center_guid4 ).
              ADD 1 TO lv_countlines.

              lv_tothoras  = lv_tothoras  + gs_somatorios-horatotal.
              lv_totamount = lv_totamount + gs_somatorios-valortotal.
            ENDLOOP.
            lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'C' ip_value = 'Total'      ip_style = tp_style_bold_center_guid9 ).
            lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'D' ip_value = lv_tothoras  ip_style = tp_style_bold_center_guid5 ).
            lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'E' ip_value = lv_totamount ip_style = tp_style_bold_center_guid5 ).

            ADD 2 TO lv_countlines.
            CLEAR gs_somatorios.
            CLEAR: lv_tothoras, lv_totamount.
          ENDIF.
          CLEAR gt_somatorios.
          gs_somatorios-assunto    = gs_data-pep_mask.
          gs_somatorios-horatotal  = lv_total_hours.
          gs_somatorios-valortotal = lv_total_amount.
          APPEND gs_somatorios TO gt_somatorios.
          gv_assunto_pai           = gv_assunto_filho.
        ENDIF.
      ENDIF.

      "ASSUNTO
      "----------------------------------------------------------------------------------------------------------------------------------------------------------
      lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'A' ip_value = |{ gs_data-kunnr } { gs_data-name1 }|    ip_style = tp_style_bold_center_guid8 ).
      ADD 1 TO lv_countlines.
      lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'A' ip_value = |{ gs_data-pep_mask } { gs_data-post1 }| ip_style = tp_style_bold_center_guid8 ).
      ADD 1 TO lv_countlines.

      lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'A' ip_value = 'Data'           ip_style = tp_style_bold_center_guid ).
      lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'B' ip_value = 'Discriminativo' ip_style = tp_style_bold_center_guid ).
      lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'C' ip_value = 'Horas'          ip_style = tp_style_bold_center_guid ).
      lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'D' ip_value = 'Preço/Hora'     ip_style = tp_style_bold_center_guid ).
      lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'E' ip_value = 'Montante'       ip_style = tp_style_bold_center_guid ).
      lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'F' ip_value = 'Nome'           ip_style = tp_style_bold_center_guid ).
      ADD 1 TO lv_countlines.

      LOOP AT gt_honorarios INTO DATA(gs_data2) WHERE pep_mask EQ gs_data-pep_mask.

        lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'A' ip_value = gs_data2-data               ip_style = tp_style_bold_center_guid3 ).
        lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'B' ip_value = gs_data2-descriminativo_str ip_style = tp_style_bold_center_guid6 ).
        lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'C' ip_value = gs_data2-horas              ip_style = tp_style_bold_center_guid4 ).
        lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'D' ip_value = gs_data2-pvp                ip_style = tp_style_bold_center_guid4 ).
        lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'E' ip_value = gs_data2-montante           ip_style = tp_style_bold_center_guid2 ).
        lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'F' ip_value = gs_data2-rufnm              ip_style = tp_style_bold_center_guid2 ).
        ADD 1 TO lv_countlines.

        "sum hours and amount to pay
        lv_total_hours  = lv_total_hours  + gs_data2-horas.
        lv_total_amount = lv_total_amount + gs_data2-montante.

      ENDLOOP.

      lv_old_assunto = gs_data-pep_mask.

      lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'B' ip_value = 'Totais'        ip_style = tp_style_bold_center_guid9 ).
      lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'C' ip_value = lv_total_hours  ip_style = tp_style_bold_center_guid5 ).
      lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'E' ip_value = lv_total_amount ip_style = tp_style_bold_center_guid5 ).
      ADD 2 TO lv_countlines.
      ADD 1 TO lv_index_assuntos.

      "insert amounts into sums by projects
      LOOP AT gt_somatorios INTO gs_somatorios WHERE assunto EQ gs_data-pep_mask.
        gs_somatorios-horatotal  = lv_total_hours.
        gs_somatorios-valortotal = lv_total_amount.
        MODIFY gt_somatorios FROM gs_somatorios.
      ENDLOOP.
      CLEAR: lv_total_hours, lv_total_amount.

    ENDLOOP.

    "Tabelas de Montantes por Projetos
    "--------------------------------------------------------------------------------------------------------------------------------------------
    IF lines( gt_somatorios ) > 1.
      lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'C' ip_value = 'Resumo Discriminativo' ip_style = tp_style_bold_center_guid ).
      lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'D' ip_value = 'Horas'                 ip_style = tp_style_bold_center_guid ).
      lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'E' ip_value = 'Montante'              ip_style = tp_style_bold_center_guid ).
      ADD 1 TO lv_countlines.

      LOOP AT gt_somatorios INTO gs_somatorios.
        lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'C' ip_value = gs_somatorios-assunto    ip_style = tp_style_bold_center_guid10 ).
        lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'D' ip_value = gs_somatorios-horatotal  ip_style = tp_style_bold_center_guid4 ).
        lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'E' ip_value = gs_somatorios-valortotal ip_style = tp_style_bold_center_guid4 ).
        ADD 1 TO lv_countlines.

        lv_tothoras  = lv_tothoras  + gs_somatorios-horatotal.
        lv_totamount = lv_totamount + gs_somatorios-valortotal.
      ENDLOOP.
      lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'C' ip_value = 'Total'      ip_style = tp_style_bold_center_guid9 ).
      lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'D' ip_value = lv_tothoras  ip_style = tp_style_bold_center_guid5 ).
      lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'E' ip_value = lv_totamount ip_style = tp_style_bold_center_guid5 ).

      ADD 2 TO lv_countlines.
    ENDIF.

    CLEAR gs_somatorios.
    CLEAR gt_somatorios.
    CLEAR: lv_tothoras, lv_totamount.
    CLEAR lv_index_assuntos.

    "DESPESAS
    "--------------------------------------------------------------------------------------------------------------------------------------------------
    IF gt_despesas IS NOT INITIAL.

      LOOP AT gt_despesas INTO DATA(gs_despesas3).

        IF lv_old_subassunto_despesa EQ gs_despesas3-pep_mask.
          CONTINUE.
        ENDIF.

        "Tabelas de Montantes por Projetos e Subprojetos
        "--------------------------------------------------------------------------------------------------------------------------------------------
        IF lv_index_assuntos EQ 0.
          gv_strlen_assunto_pai    = strlen( gs_despesas3-pep_mask ) - 1.
          gv_assunto_pai           = substring( val = gs_despesas3-pep_mask off = 0 len = gv_strlen_assunto_pai ).
          gs_somatorios-assunto    = gs_despesas3-pep_mask.
          gs_somatorios-horatotal  = lv_total_hours.
          gs_somatorios-valortotal = lv_total_amount.
          APPEND gs_somatorios TO gt_somatorios.
        ELSE.
          gv_strlen_assunto_filho  = strlen( gs_despesas3-pep_mask ) - 1.
          gv_assunto_filho         = substring( val = gs_despesas3-pep_mask off = 0 len = gv_strlen_assunto_filho ).
          IF gv_assunto_pai EQ gv_assunto_filho.
            gs_somatorios-assunto    = gs_despesas3-pep_mask.
            gs_somatorios-horatotal  = lv_total_hours.
            gs_somatorios-valortotal = lv_total_amount.
            APPEND gs_somatorios TO gt_somatorios.
            gv_assunto_pai = gv_assunto_filho.
          ELSE.
            IF lines( gt_somatorios ) > 1.
              lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'C' ip_value = 'Resumo Despesas' ip_style = tp_style_bold_center_guid ).
              lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'D' ip_value = 'Montante'        ip_style = tp_style_bold_center_guid ).
              ADD 1 TO lv_countlines.

              LOOP AT gt_somatorios INTO gs_somatorios.
                lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'C' ip_value = gs_somatorios-assunto    ip_style = tp_style_bold_center_guid10 ).
                lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'D' ip_value = gs_somatorios-valortotal ip_style = tp_style_bold_center_guid4 ).
                ADD 1 TO lv_countlines.

                lv_tothoras  = lv_tothoras  + gs_somatorios-horatotal.
                lv_totamount = lv_totamount + gs_somatorios-valortotal.
              ENDLOOP.
              lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'C' ip_value = 'Total'      ip_style = tp_style_bold_center_guid9 ).
              lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'D' ip_value = lv_totamount ip_style = tp_style_bold_center_guid5 ).

              ADD 2 TO lv_countlines.
              CLEAR gs_somatorios.
              CLEAR: lv_tothoras, lv_totamount.
            ENDIF.
            CLEAR gt_somatorios.
            gs_somatorios-assunto    = gs_data-pep_mask.
            gs_somatorios-horatotal  = lv_total_hours.
            gs_somatorios-valortotal = lv_total_amount.
            APPEND gs_somatorios TO gt_somatorios.
            gv_assunto_pai           = gv_assunto_filho.
          ENDIF.
        ENDIF.

        "DESPESAS
        "--------------------------------------------------------------------------------------------------------------------------------------------------------------
        lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'A' ip_value = |{ gs_despesas3-kunnr } { gs_despesas3-name1 }|    ip_style = tp_style_bold_center_guid8 ).
        ADD 1 TO lv_countlines.
        lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'A' ip_value = |{ gs_despesas3-pep_mask } { gs_despesas3-post1 }| ip_style = tp_style_bold_center_guid8 ).
        ADD 1 TO lv_countlines.

        lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'A' ip_value = 'Data'                 ip_style = tp_style_bold_center_guid ).
        lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'B' ip_value = 'Despesas'             ip_style = tp_style_bold_center_guid ).
        lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'C' ip_value = 'Designação'           ip_style = tp_style_bold_center_guid ).
        lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'D' ip_value = 'Relatório da Despesa' ip_style = tp_style_bold_center_guid ).
        lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'E' ip_value = 'Montante'             ip_style = tp_style_bold_center_guid ).
        lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'F' ip_value = 'Nome'                 ip_style = tp_style_bold_center_guid ).
        ADD 1 TO lv_countlines.

        LOOP AT gt_despesas INTO DATA(gs_despesas4) WHERE pep_mask EQ gs_despesas3-pep_mask.

          lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'A' ip_value = gs_despesas4-data               ip_style = tp_style_bold_center_guid3 ).
          lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'B' ip_value = gs_despesas4-descriminativo_str ip_style = tp_style_bold_center_guid6 ).
          lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'C' ip_value = gs_despesas4-type_txt           ip_style = tp_style_bold_center_guid4 ).
          lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'D' ip_value = gs_despesas4-relexp             ip_style = tp_style_bold_center_guid4 ).
          lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'E' ip_value = gs_despesas4-montante           ip_style = tp_style_bold_center_guid2 ).
          lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'F' ip_value = gs_despesas4-rufnm              ip_style = tp_style_bold_center_guid2 ).
          ADD 1 TO lv_countlines.

          "sum hours and amount to pay
          lv_total_hours  = lv_total_hours  + gs_despesas4-horas.
          lv_total_amount = lv_total_amount + gs_despesas4-montante.

        ENDLOOP.

        lv_old_subassunto_despesa = gs_despesas3-pep_mask.

        lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'D' ip_value = 'Totais'        ip_style = tp_style_bold_center_guid9 ).
        lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'E' ip_value = lv_total_amount ip_style = tp_style_bold_center_guid5 ).
        ADD 2 TO lv_countlines.
        ADD 1 TO lv_index_assuntos.

        "insert amounts into sums by projects
        LOOP AT gt_somatorios INTO gs_somatorios WHERE assunto EQ gs_despesas3-pep_mask.
          gs_somatorios-valortotal = lv_total_amount.
          MODIFY gt_somatorios FROM gs_somatorios.
        ENDLOOP.
        CLEAR: lv_total_hours, lv_total_amount.

      ENDLOOP.

      "Tabelas de Montantes por Projetos e Subprojetos de Despesas
      "--------------------------------------------------------------------------------------------------------------------------------------------
      IF lines( gt_somatorios ) > 1.
        lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'C' ip_value = 'Resumo Despesas' ip_style = tp_style_bold_center_guid ).
        lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'D' ip_value = 'Montante'        ip_style = tp_style_bold_center_guid ).
        ADD 1 TO lv_countlines.

        LOOP AT gt_somatorios INTO gs_somatorios.
          lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'C' ip_value = gs_somatorios-assunto    ip_style = tp_style_bold_center_guid10 ).
          lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'D' ip_value = gs_somatorios-valortotal ip_style = tp_style_bold_center_guid4 ).
          ADD 1 TO lv_countlines.

          lv_tothoras  = lv_tothoras  + gs_somatorios-horatotal.
          lv_totamount = lv_totamount + gs_somatorios-valortotal.
        ENDLOOP.
        lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'C' ip_value = 'Total'      ip_style = tp_style_bold_center_guid9 ).
        lo_worksheet->set_cell( ip_row = lv_countlines ip_column = 'D' ip_value = lv_totamount ip_style = tp_style_bold_center_guid5 ).

        ADD 2 TO lv_countlines.
      ENDIF.
    ENDIF.

    CLEAR gs_somatorios.
    CLEAR gt_somatorios.
    CLEAR: lv_tothoras, lv_totamount.
    CLEAR lv_index_assuntos.

    "COLUMNS
    "-----------------------------------------------------------------
    lo_column = lo_worksheet->get_column_dimension( ip_column = 'A' ).
    lo_column->set_width( ip_width = 10 ).

    lo_column = lo_worksheet->get_column_dimension( ip_column = 'B' ).
    lo_column->set_width( ip_width = 100 ).

    lo_column = lo_worksheet->get_column_dimension( ip_column = 'C' ).
    lo_column->set_width( ip_width = 28 ).

    lo_column->set_auto_size( ip_auto_size = abap_false ).
    lo_column = lo_worksheet->get_column_dimension( ip_column = 'D' ).
    lo_column->set_width( ip_width = 12 ).

    lo_column->set_auto_size( ip_auto_size = abap_false ).
    lo_column = lo_worksheet->get_column_dimension( ip_column = 'E' ).
    lo_column->set_width( ip_width = 12 ).

    lo_column->set_auto_size( ip_auto_size = abap_false ).
    lo_column = lo_worksheet->get_column_dimension( ip_column = 'F' ).
    lo_column->set_width( ip_width = 20 ).
*    RLA@SBX-ID-EXCEL_FORM_004

  ENDMETHOD.

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
  "   - 23.09.2023 :: Removed blank lines between header and table header
  "--------------------------------------------------------------------------------------------

  METHOD set_sheets.
    "--------------------------------------------------------------------------------------------
    " DATA       : 12.08.2025
    " CREATED BY : RLA@SBX
    "
    " FUNCTIONALITY:
    "   - two ways for print options (expensies or fees)
    "--------------------------------------------------------------------------------------------

    "this method prints expenses or not if exists
    IF gt_honorarios IS NOT INITIAL.
      print_honorarios_despesas( ).
    ELSEIF gt_honorarios IS INITIAL AND gt_despesas IS NOT INITIAL.
      print_despesas( ).
    ENDIF.

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
    "   - 23.09.2025 :: set customers into new table for kunnr
    "   - 23.09.2025 :: set subprojects into new table for projects
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
    "
    " UPDATES
    "   - 25.06.2025 :: New method set_sheets
    "--------------------------------------------------------------------------------------------

    get_data( ).
    set_style( ).

*    RLA@SBX-ID-EXCEL_FORM_004
    set_sheets( ).
*    RLA@SBX-ID-EXCEL_FORM_004

    download_xls( ).

  ENDMETHOD.

  METHOD set_sheet_title.

    "--------------------------------------------------------------------------------------------
    " DATA       : 12.08.2025
    " CREATED BY : RLA@SBX
    "
    " FUNCTIONALITY:
    "   - set title to worksheet
    "
    " UPDATES
    "   - 25.06.2025 :: gridlines removed from template
    "--------------------------------------------------------------------------------------------

    DATA(o_xl_ws) = o_xl->get_active_worksheet( ).
    lo_worksheet = o_xl_ws.

    gv_title = | { sy-datum } |.
    lo_worksheet->set_title( ip_title = gv_title ).

*    RLA@SBX-ID-EXCEL_FORM_004
    lo_worksheet->set_show_gridlines( i_show_gridlines = abap_false ).
*    RLA@SBX-ID-EXCEL_FORM_004

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
    "   - 23.09.2023 :: All styles had their horizontal aligments justified by center
    "                   Two decimals places for amount column style
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

    "header table style
    lo_style                             = o_xl->add_new_style( ).
    lo_style->font->bold                 = abap_true.
    lo_style->font->italic               = abap_false.
    lo_style->font->name                 = zcl_excel_style_font=>c_name_arial.
    lo_style->font->scheme               = zcl_excel_style_font=>c_scheme_none.
    lo_style->font->size                 = 10.
    lo_style->font->color-rgb            = zcl_excel_style_color=>c_black.
    lo_style->alignment->horizontal      = zcl_excel_style_alignment=>c_horizontal_center.
    lo_style->alignment->vertical        = zcl_excel_style_alignment=>c_vertical_center.
    lo_style->alignment->wraptext        = abap_true.
    lo_style->borders->allborders        = o_border_light.
    lo_style->fill->filltype             = zcl_excel_style_fill=>c_fill_solid.
    lo_style->fill->fgcolor-rgb          = zcl_excel_style_color=>c_gray.
    tp_style_bold_center_guid            = lo_style->get_guid( ). "nao esquecer

    "total do assunto e total despesas style
    lo_style                             = o_xl->add_new_style( ).
    lo_style->font->bold                 = abap_true.
    lo_style->font->italic               = abap_false.
    lo_style->font->name                 = zcl_excel_style_font=>c_name_arial.
    lo_style->font->scheme               = zcl_excel_style_font=>c_scheme_none.
    lo_style->font->size                 = 10.
    lo_style->font->color-rgb            = zcl_excel_style_color=>c_black.
    lo_style->alignment->horizontal      = zcl_excel_style_alignment=>c_horizontal_right.
    lo_style->alignment->vertical        = zcl_excel_style_alignment=>c_vertical_top.
*    lo_style->borders->allborders        = o_border_light.
*    lo_style->fill->filltype             = zcl_excel_style_fill=>c_fill_solid.
*    lo_style->fill->fgcolor-rgb          = zcl_excel_style_color=>c_white.
    tp_style_bold_center_guid9           = lo_style->get_guid( ). "nao esquecer

    "default
    lo_style                             = o_xl->add_new_style( ).
    lo_style->font->bold                 = abap_true.
    lo_style->font->italic               = abap_false.
    lo_style->font->name                 = zcl_excel_style_font=>c_name_arial.
    lo_style->font->scheme               = zcl_excel_style_font=>c_scheme_none.
    lo_style->font->size                 = 10.
    lo_style->font->color-rgb            = zcl_excel_style_color=>c_black.
    lo_style->alignment->horizontal      = zcl_excel_style_alignment=>c_horizontal_left.
    lo_style->alignment->vertical        = zcl_excel_style_alignment=>c_vertical_top.
*    lo_style->borders->allborders        = o_border_dark.
    tp_style_bold_center_guid8           = lo_style->get_guid( ). "nao esquecer

    "amount column style
    lo_style                             = o_xl->add_new_style( ).
    lo_style->font->bold                 = abap_false.
    lo_style->font->italic               = abap_false.
    lo_style->font->name                 = zcl_excel_style_font=>c_name_arial.
    lo_style->font->scheme               = zcl_excel_style_font=>c_scheme_none.
    lo_style->font->size                 = 10.
    lo_style->font->color-rgb            = zcl_excel_style_color=>c_black.
    lo_style->alignment->horizontal      = zcl_excel_style_alignment=>c_horizontal_center.
    lo_style->alignment->vertical        = zcl_excel_style_alignment=>c_vertical_top.
    lo_style->borders->allborders        = o_border_dark.
    lo_style->number_format->format_code = '#,##0.00'.
    tp_style_bold_center_guid2           = lo_style->get_guid( ). "nao esquecer

    "dates style
    lo_style                             = o_xl->add_new_style( ).
    lo_style->font->bold                 = abap_false.
    lo_style->font->italic               = abap_false.
    lo_style->font->name                 = zcl_excel_style_font=>c_name_arial.
    lo_style->font->scheme               = zcl_excel_style_font=>c_scheme_none.
    lo_style->font->size                 = 10.
    lo_style->font->color-rgb            = zcl_excel_style_color=>c_black.
    lo_style->alignment->horizontal      = zcl_excel_style_alignment=>c_horizontal_center.
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
    lo_style->alignment->horizontal      = zcl_excel_style_alignment=>c_horizontal_center.
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
    lo_style->alignment->horizontal      = zcl_excel_style_alignment=>c_horizontal_center.
*    lo_style->borders->allborders        = o_border_dark.
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
    lo_style->alignment->vertical        = zcl_excel_style_alignment=>c_vertical_top.
    lo_style->alignment->wraptext        = abap_true.
    lo_style->borders->allborders        = o_border_dark.
    lo_style->fill->filltype             = zcl_excel_style_fill=>c_fill_solid.
    lo_style->fill->fgcolor-rgb          = zcl_excel_style_color=>c_white.
    tp_style_bold_center_guid6           = lo_style->get_guid( ). "nao esquecer

*    RLA@SBX-ID-EXCEL_FORM_004
    "total name style
    lo_style                             = o_xl->add_new_style( ).
    lo_style->font->bold                 = abap_true.
    lo_style->font->italic               = abap_false.
    lo_style->font->name                 = zcl_excel_style_font=>c_name_arial.
    lo_style->font->scheme               = zcl_excel_style_font=>c_scheme_none.
    lo_style->font->size                 = 10.
    lo_style->font->color-rgb            = zcl_excel_style_color=>c_black.
    lo_style->alignment->horizontal      = zcl_excel_style_alignment=>c_horizontal_center.
    lo_style->alignment->vertical        = zcl_excel_style_alignment=>c_vertical_top.
    tp_style_bold_center_guid7           = lo_style->get_guid( ). "nao esquecer

    "total amount and hours style assunto text for sup table
    lo_style                             = o_xl->add_new_style( ).
    lo_style->font->bold                 = abap_true.
    lo_style->font->italic               = abap_false.
    lo_style->font->name                 = zcl_excel_style_font=>c_name_arial.
    lo_style->font->scheme               = zcl_excel_style_font=>c_scheme_none.
    lo_style->font->size                 = 10.
    lo_style->font->color-rgb            = zcl_excel_style_color=>c_black.
    lo_style->alignment->horizontal      = zcl_excel_style_alignment=>c_horizontal_center.
    lo_style->borders->allborders        = o_border_dark.
    lo_style->fill->filltype             = zcl_excel_style_fill=>c_fill_solid.
    lo_style->fill->fgcolor-rgb          = zcl_excel_style_color=>c_white.
*    lo_style->number_format->format_code = '#,##0.00'.
    tp_style_bold_center_guid10           = lo_style->get_guid( ). "nao esquecer
*    RLA@SBX-ID-EXCEL_FORM_004

    "é possível montar vários estilos guid e usar da forma como convém

  ENDMETHOD.

ENDCLASS.
