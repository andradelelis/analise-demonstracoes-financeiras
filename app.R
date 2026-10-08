# app.R
# Shiny: 3 abas
# (1) Demonstrações: baixar/visualizar (wide) por tipo_doc escolhido
# (2) Indicadores: calcula por CD_CONTA (DRE + BPA + BPP) e gera tabela + gráficos
# (3) Fórmulas: lista todas as fórmulas usadas no app

library(shiny)
library(GetDFPData2)
library(dplyr)
library(stringr)
library(openxlsx)
library(tidyr)
library(ggplot2)

# =========================================================
# 1) MAPA DE CONTAS (EDITE AQUI)
# =========================================================
MAPA_CONTAS <- list(
  # --- DRE ---
  receita_liquida       = c("3.01"),
  lucro_bruto           = c("3.03"),
  resultado_operacional = c("3.05"),   # EBIT / Resultado Operacional
  lucro_liquido         = c("3.11"),
  
  # --- BPA (Ativo) ---
  ativo_total        = c("1"),
  ativo_circulante   = c("1.01"),
  estoques           = c("1.01.04"),
  caixa_equivalentes = c("1.01.01"),
  
  # --- BPP (Passivo/PL) ---
  passivo_total      = c("2"),
  passivo_circulante = c("2.01"),
  patrimonio_liquido = c("2.03"),
  
  # --- Dívida (BPP) ---
  divida_curto_prazo = c("2.01.04"),
  divida_longo_prazo = c("2.02.01")
)

# Contas essenciais para prévia do DFC_MI (na tela)
CONTAS_DFC_MI_PREVIA <- c(
  "6.01","6.01.01","6.01.01.01","6.01.02","6.01.02.01","6.01.02.02",
  "6.02","6.02.01",
  "6.03","6.03.01","6.03.02","6.03.03","6.03.04","6.03.05",
  "6.04","6.05","6.05.01","6.05.02"
)

# =========================================================
# Fórmulas (aba dedicada)
# =========================================================
FORMULAS <- tibble::tribble(
  ~Indicador, ~Formula,
  "Vendas (Receita Líquida)", "receita = soma(VL_CONTA) das contas em MAPA_CONTAS$receita_liquida",
  "EBIT (Resultado Operacional)", "ebit = soma(VL_CONTA) das contas em MAPA_CONTAS$resultado_operacional",
  "Margem Bruta", "margem_bruta = lucro_bruto / receita",
  "Margem Operacional", "margem_operacional = ebit / receita",
  "Margem Líquida", "margem_liquida = lucro_liquido / receita",
  "ROA", "roa = lucro_liquido / ativo_total",
  "ROE", "roe = lucro_liquido / patrimonio_liquido",
  "Dívida Bruta", "divida_bruta = divida_cp + divida_lp",
  "Dívida Líquida", "divida_liquida = divida_bruta - caixa_equivalentes",
  "Dívida Líq / PL", "divida_liq_pl = divida_liquida / patrimonio_liquido",
  "Endividamento Total", "endividamento_total = (divida_cp + divida_lp) / ativo_total",
  "Endividamento CP", "endividamento_curto_prazo = passivo_circulante / ativo_total",
  "Liquidez Corrente", "liquidez_corrente = ativo_circulante / passivo_circulante",
  "CCL", "ccl = ativo_circulante - passivo_circulante",
  "Liquidez Seca", "liquidez_seca = (ativo_circulante - estoques) / passivo_circulante",
  "Liquidez Imediata", "liquidez_imediata = caixa_equivalentes / passivo_circulante"
)

# =========================================================
# UI
# =========================================================
ui <- fluidPage(
  titlePanel("Análise de Demonstrações Financeiras - DFP-CVM"),
  
  sidebarLayout(
    sidebarPanel(
      h4("1) Encontrar empresa"),
      textInput("busca_nome", "Buscar por nome:", value = "petrobras"),
      actionButton("btn_buscar", "Buscar empresas", class = "btn-primary"),
      br(), br(),
      
      h4("2) Selecionar empresa (1)"),
      uiOutput("ui_empresas"),
      br(),
      
      h4("3) Configurações gerais"),
      selectInput(
        "tipo_format",
        "Formato:",
        choices = c("Consolidado (con)" = "con", "Individual (ind)" = "ind"),
        selected = "con"
      ),
      numericInput("ano_ini", "Ano inicial (máx. 5 anos):", value = 2021, min = 1995, max = 2025),
      numericInput("ano_fim", "Ano final:", value = 2025, min = 1995, max = 2025),
      checkboxInput("usar_memoise", "Usar cache (memoise)", value = TRUE),
      br(),
      
      h4("4) Unidade do valor"),
      selectInput(
        "unidade",
        "Escala para VL_CONTA:",
        choices = c("R$ (unidade)" = "1", "R$ mil" = "1000", "R$ mi" = "1000000"),
        selected = "1000"
      ),
      br(),
      
      h4("Aba 1: Demonstrações Financeiras"),
      selectInput(
        "tipo_doc",
        "Demonstração para visualizar/baixar:",
        choices = c(
          "DRE (Resultado)" = "DRE",
          "DFC – Método Indireto (DFC_MI)" = "DFC_MI",
          "BPA (Ativo)" = "BPA",
          "BPP (Passivo)" = "BPP",
          "DMPL (Mutações do PL)" = "DMPL",
          "DVA (Valor Adicionado)" = "DVA"
        ),
        selected = "DRE"
      ),
      selectInput(
        "modo_previa",
        "Prévia na tela:",
        choices = c("Automático (DFC_MI=essenciais)" = "auto", "Tudo" = "all"),
        selected = "auto"
      ),
      downloadButton("baixar_xlsx", "Baixar Excel (Completo)", class = "btn-success")
    ),
    
    mainPanel(
      tabsetPanel(
        tabPanel(
          "Demonstrações",
          h4("Empresas encontradas"),
          tableOutput("tbl_encontradas"),
          br(),
          h4("Status"),
          verbatimTextOutput("status"),
          br(),
          h4("Prévia (Wide)"),
          tableOutput("tbl_preview")
        ),
        tabPanel(
          "Indicadores",
          br(),
          h4("Tabela de indicadores"),
          tableOutput("tbl_indicadores"),
          br(),
          h4("Gráficos"),
          plotOutput("plot_vendas_ebit"),
          plotOutput("plot_margens"),
          plotOutput("plot_rentab"),
          plotOutput("plot_liq"),
          plotOutput("plot_endiv")
        ),
        tabPanel(
          "Fórmulas",
          br(),
          h4("Fórmulas utilizadas no app"),
          tableOutput("tbl_formulas"),
          br(),
          helpText("Observação: VL_CONTA é escalonado conforme a unidade escolhida no sidebar.")
        )
      )
    )
  )
)

# =========================================================
# Server
# =========================================================
server <- function(input, output, session) {
  
  # ---------- Helpers ----------
  unidade_label <- reactive({
    esc <- as.numeric(input$unidade)
    dplyr::case_when(
      esc == 1 ~ "R$ (unidade)",
      esc == 1000 ~ "R$ mil",
      esc == 1000000 ~ "R$ mi",
      TRUE ~ paste0("R$ / ", esc)
    )
  })
  
  doc_label <- reactive({
    switch(input$tipo_doc,
           "DFC_MD" = "DFC (Método Direto)",
           "DFC_MI" = "DFC (Método Indireto)",
           input$tipo_doc
    )
  })
  
  achar_col_ano <- function(df) {
    if ("ano" %in% names(df)) return("ano")
    candidatos <- c("DT_REFER", "DT_FIM_EXERC", "DT_INI_EXERC")
    candidatos <- candidatos[candidatos %in% names(df)]
    if (length(candidatos) > 0) return(candidatos[1])
    dtc <- names(df)[str_detect(names(df), "^DT_")]
    if (length(dtc) > 0) return(dtc[1])
    NA_character_
  }
  
  to_year <- function(x) {
    if (inherits(x, "Date")) return(as.integer(format(x, "%Y")))
    if (is.numeric(x)) return(as.integer(x))
    suppressWarnings({
      dx <- as.Date(x)
      if (!all(is.na(dx))) return(as.integer(format(dx, "%Y")))
    })
    y <- suppressWarnings(as.integer(str_extract(as.character(x), "\\b\\d{4}\\b")))
    y
  }
  
  # Baixa uma demonstração e devolve em LONG
  baixar_demo_long <- function(codigo_num, ano_ini, ano_fim, tipo_doc, tipo_format, escala, usar_memoise) {
    res_list <- get_dfp_data(
      companies_cvm_codes = codigo_num,
      first_year = ano_ini,
      last_year  = ano_fim,
      type_docs  = tipo_doc,
      type_format = tipo_format,
      clean_data  = TRUE,
      use_memoise = isTRUE(usar_memoise)
    )
    
    demo_raw <- NULL
    if (is.data.frame(res_list)) demo_raw <- res_list
    if (is.list(res_list) && length(res_list) >= 1 && is.data.frame(res_list[[1]])) demo_raw <- res_list[[1]]
    
    validate(need(!is.null(demo_raw) && nrow(demo_raw) > 0, paste("Sem dados para", tipo_doc)))
    
    ano_col <- achar_col_ano(demo_raw)
    validate(need(!is.na(ano_col), paste("Sem coluna de ano/data em", tipo_doc)))
    
    demo_raw %>%
      select(any_of(c("CD_CVM","DENOM_CIA","CD_CONTA","DS_CONTA","VL_CONTA", ano_col))) %>%
      mutate(
        ANO = if (ano_col == "ano") as.integer(.data[[ano_col]]) else to_year(.data[[ano_col]]),
        VL_CONTA = as.numeric(VL_CONTA) / as.numeric(escala)
      ) %>%
      filter(!is.na(ANO)) %>%
      group_by(CD_CVM, DENOM_CIA, CD_CONTA, DS_CONTA, ANO) %>%
      summarise(VL_CONTA = sum(VL_CONTA, na.rm = TRUE), .groups = "drop")
  }
  
  construir_wide_puro <- function(df_long) {
    anos <- sort(unique(df_long$ANO))
    validate(need(length(anos) > 0, "Não encontrei anos no retorno para montar colunas."))
    df_long %>%
      select(CD_CONTA, DS_CONTA, ANO, VL_CONTA) %>%
      group_by(CD_CONTA, DS_CONTA, ANO) %>%
      summarise(VL_CONTA = sum(VL_CONTA, na.rm = TRUE), .groups = "drop") %>%
      pivot_wider(names_from = ANO, values_from = VL_CONTA) %>%
      arrange(CD_CONTA)
  }
  
  pegar_serie_por_cd <- function(df_long, cds) {
    if (length(cds) == 0 || all(is.na(cds))) {
      return(tibble(ANO = integer(), valor = numeric()))
    }
    df_long %>%
      filter(CD_CONTA %in% cds) %>%
      group_by(ANO) %>%
      summarise(valor = sum(VL_CONTA, na.rm = TRUE), .groups = "drop")
  }
  
  # ---------- Busca empresas ----------
  empresas_encontradas <- reactiveVal(NULL)
  
  observeEvent(input$btn_buscar, {
    req(input$busca_nome)
    termo <- str_trim(input$busca_nome)
    
    res <- tryCatch(search_company(termo), error = function(e) NULL)
    empresas_encontradas(res)
    
    if (!is.null(res) && nrow(res) > 0) {
      col_code <- names(res)[str_detect(names(res), regex("cvm|code", ignore_case = TRUE))][1]
      col_name <- names(res)[str_detect(names(res), regex("name|denom|company", ignore_case = TRUE))][1]
      if (is.na(col_code) || is.null(col_code)) col_code <- names(res)[1]
      if (is.na(col_name) || is.null(col_name)) col_name <- names(res)[2]
      
      choices <- setNames(res[[col_code]], paste0(res[[col_name]], " (", res[[col_code]], ")"))
      updateSelectizeInput(session, "codigo", choices = choices, server = TRUE)
    } else {
      updateSelectizeInput(session, "codigo", choices = character(0), server = TRUE)
    }
  })
  
  output$ui_empresas <- renderUI({
    selectizeInput("codigo", "Código CVM (selecione 1):", choices = character(0), multiple = FALSE)
  })
  
  output$tbl_encontradas <- renderTable({
    empresas_encontradas()
  }, striped = TRUE, hover = TRUE, spacing = "s")
  
  # ---------- Validação de anos ----------
  validar_periodo <- reactive({
    validate(
      need(!is.null(input$ano_ini) && !is.null(input$ano_fim), "Informe o período."),
      need(input$ano_ini <= input$ano_fim, "Ano inicial deve ser <= ano final."),
      need((input$ano_fim - input$ano_ini) <= 4, "Limite: no máximo 5 anos (ex.: 2020 a 2024).")
    )
    TRUE
  })
  
  # =========================================================
  # ABA 1: Demonstrações Financeiras
  # =========================================================
  dados_demo_completo <- reactive({
    req(input$codigo)
    validar_periodo()
    
    codigo_num <- as.integer(input$codigo)
    escala <- as.numeric(input$unidade)
    
    withProgress(message = "Carregando demonstração...", value = 0, {
      incProgress(0.2)
      df_long <- baixar_demo_long(
        codigo_num = codigo_num,
        ano_ini = input$ano_ini,
        ano_fim = input$ano_fim,
        tipo_doc = input$tipo_doc,
        tipo_format = input$tipo_format,
        escala = escala,
        usar_memoise = input$usar_memoise
      )
      incProgress(0.8)
      
      wide <- construir_wide_puro(df_long)
      
      list(
        cd_cvm = unique(df_long$CD_CVM)[1],
        denom  = unique(df_long$DENOM_CIA)[1],
        wide   = wide
      )
    })
  })
  
  dados_demo_previa <- reactive({
    x <- dados_demo_completo()
    req(x)
    df <- x$wide
    if (input$modo_previa == "all") return(df)
    if (input$tipo_doc == "DFC_MI") df %>% filter(CD_CONTA %in% CONTAS_DFC_MI_PREVIA) else df
  })
  
  output$status <- renderPrint({
    if (is.null(input$codigo) || identical(input$codigo, "")) {
      cat("Selecione a empresa.\n")
      cat("Demonstração:", doc_label(), "\n")
      cat("Formato:", input$tipo_format, "\n")
      cat("Período:", input$ano_ini, "a", input$ano_fim, " (máx. 5 anos)\n")
      cat("Unidade:", unidade_label(), "\n")
      return()
    }
    
    x <- tryCatch(dados_demo_completo(), error = function(e) e)
    
    cat("Empresa (CVM):", input$codigo, "\n")
    cat("Demonstração:", doc_label(), "\n")
    cat("Formato:", input$tipo_format, "\n")
    cat("Período:", input$ano_ini, "a", input$ano_fim, " (máx. 5 anos)\n")
    cat("Unidade:", unidade_label(), "\n")
    cat("Cache (memoise):", ifelse(isTRUE(input$usar_memoise), "ON", "OFF"), "\n")
    
    if (inherits(x, "error")) {
      cat("\n(Erro ao carregar dados)\n")
      cat(conditionMessage(x), "\n")
    } else {
      cat("\nEmpresa:", x$denom, "\n")
      cat("Código CVM:", x$cd_cvm, "\n")
      if (input$tipo_doc == "DFC_MI" && input$modo_previa != "all") {
        cat("\nPrévia DFC_MI: exibindo somente contas essenciais.\n")
      }
    }
  })
  
  output$tbl_preview <- renderTable({
    df <- dados_demo_previa()
    req(df)
    df %>% slice_head(n = 50)
  }, striped = TRUE, hover = TRUE, spacing = "s")
  
  output$baixar_xlsx <- downloadHandler(
    filename = function() {
      paste0(input$tipo_doc, "_", input$tipo_format, "_", input$codigo, "_",
             input$ano_ini, "_", input$ano_fim, "_", Sys.Date(), ".xlsx")
    },
    content = function(file) {
      x <- dados_demo_completo()
      req(x)
      dfw <- x$wide
      
      wb <- createWorkbook()
      sheet_name <- str_sub(paste0(input$tipo_doc, "_", input$codigo), 1, 31)
      addWorksheet(wb, sheet_name)
      writeData(wb, sheet = sheet_name, x = dfw)
      saveWorkbook(wb, file, overwrite = TRUE)
    }
  )
  
  # =========================================================
  # ABA 2: Índices e Gráficos
  # =========================================================
  indicadores <- reactive({
    req(input$codigo)
    validar_periodo()
    
    codigo_num <- as.integer(input$codigo)
    escala <- as.numeric(input$unidade)
    
    withProgress(message = "Calculando indicadores (DRE + BP)...", value = 0, {
      incProgress(0.15)
      dre <- baixar_demo_long(codigo_num, input$ano_ini, input$ano_fim, "DRE", input$tipo_format, escala, input$usar_memoise)
      incProgress(0.45)
      bpa <- baixar_demo_long(codigo_num, input$ano_ini, input$ano_fim, "BPA", input$tipo_format, escala, input$usar_memoise)
      incProgress(0.75)
      bpp <- baixar_demo_long(codigo_num, input$ano_ini, input$ano_fim, "BPP", input$tipo_format, escala, input$usar_memoise)
      incProgress(0.9)
      
      anos <- tibble(ANO = seq(input$ano_ini, input$ano_fim))
      
      receita <- pegar_serie_por_cd(dre, MAPA_CONTAS$receita_liquida) %>% rename(receita = valor)
      ebit    <- pegar_serie_por_cd(dre, MAPA_CONTAS$resultado_operacional) %>% rename(ebit = valor)
      lb      <- pegar_serie_por_cd(dre, MAPA_CONTAS$lucro_bruto) %>% rename(lucro_bruto = valor)
      ll      <- pegar_serie_por_cd(dre, MAPA_CONTAS$lucro_liquido) %>% rename(lucro_liquido = valor)
      
      at      <- pegar_serie_por_cd(bpa, MAPA_CONTAS$ativo_total) %>% rename(ativo_total = valor)
      ac      <- pegar_serie_por_cd(bpa, MAPA_CONTAS$ativo_circulante) %>% rename(ativo_circulante = valor)
      est     <- pegar_serie_por_cd(bpa, MAPA_CONTAS$estoques) %>% rename(estoques = valor)
      caixa   <- pegar_serie_por_cd(bpa, MAPA_CONTAS$caixa_equivalentes) %>% rename(caixa_eq = valor)
      
      pc      <- pegar_serie_por_cd(bpp, MAPA_CONTAS$passivo_circulante) %>% rename(passivo_circulante = valor)
      pl      <- pegar_serie_por_cd(bpp, MAPA_CONTAS$patrimonio_liquido) %>% rename(patrimonio_liquido = valor)
      
      div_cp  <- pegar_serie_por_cd(bpp, MAPA_CONTAS$divida_curto_prazo) %>% rename(divida_cp = valor)
      div_lp  <- pegar_serie_por_cd(bpp, MAPA_CONTAS$divida_longo_prazo) %>% rename(divida_lp = valor)
      
      df <- anos %>%
        left_join(receita, by="ANO") %>%
        left_join(ebit, by="ANO") %>%
        left_join(lb, by="ANO") %>%
        left_join(ll, by="ANO") %>%
        left_join(at, by="ANO") %>%
        left_join(ac, by="ANO") %>%
        left_join(est, by="ANO") %>%
        left_join(caixa, by="ANO") %>%
        left_join(pc, by="ANO") %>%
        left_join(pl, by="ANO") %>%
        left_join(div_cp, by="ANO") %>%
        left_join(div_lp, by="ANO") %>%
        mutate(
          margem_bruta = ifelse(!is.na(receita) & receita != 0, lucro_bruto / receita, NA_real_),
          margem_operacional = ifelse(!is.na(receita) & receita != 0, ebit / receita, NA_real_),
          margem_liquida = ifelse(!is.na(receita) & receita != 0, lucro_liquido / receita, NA_real_),
          
          roa = ifelse(!is.na(ativo_total) & ativo_total != 0, lucro_liquido / ativo_total, NA_real_),
          roe = ifelse(!is.na(patrimonio_liquido) & patrimonio_liquido != 0, lucro_liquido / patrimonio_liquido, NA_real_),
          
          divida_bruta   = coalesce(divida_cp, 0) + coalesce(divida_lp, 0),
          divida_liquida = divida_bruta - coalesce(caixa_eq, 0),
          
          # Dívida Líq / PL
          divida_liq_pl = ifelse(!is.na(patrimonio_liquido) & patrimonio_liquido != 0,
                                 divida_liquida / patrimonio_liquido, NA_real_),
          
          endividamento_total = ifelse(!is.na(ativo_total) & ativo_total != 0,
                                       (coalesce(divida_cp, 0) + coalesce(divida_lp, 0)) / ativo_total, NA_real_),
          endividamento_curto_prazo = ifelse(!is.na(ativo_total) & ativo_total != 0,
                                             passivo_circulante / ativo_total, NA_real_),
          
          liquidez_corrente = ifelse(!is.na(passivo_circulante) & passivo_circulante != 0,
                                     ativo_circulante / passivo_circulante, NA_real_),
          ccl = ativo_circulante - passivo_circulante,
          liquidez_seca = ifelse(!is.na(passivo_circulante) & passivo_circulante != 0,
                                 (ativo_circulante - coalesce(estoques, 0)) / passivo_circulante, NA_real_),
          liquidez_imediata = ifelse(!is.na(passivo_circulante) & passivo_circulante != 0,
                                     coalesce(caixa_eq, 0) / passivo_circulante, NA_real_)
        )
      
      incProgress(1)
      df
    })
  })
  
  output$tbl_indicadores <- renderTable({
    df <- indicadores()
    df %>%
      transmute(
        ANO,
        Vendas = receita,
        EBIT = ebit,
        `Margem Bruta (%)` = round(100*margem_bruta, 2),
        `Margem Operacional (%)` = round(100*margem_operacional, 2),
        `Margem Líquida (%)` = round(100*margem_liquida, 2),
        `ROA (%)` = round(100*roa, 2),
        `ROE (%)` = round(100*roe, 2),
        `Dívida Líq/PL (%)` = round(100*divida_liq_pl, 2),
        `Endividamento Total (%)` = round(100*endividamento_total, 2),
        `Endividamento CP (%)` = round(100*endividamento_curto_prazo, 2),
        `Liquidez Corrente` = round(liquidez_corrente, 3),
        `CCL` = round(ccl, 2),
        `Liquidez Seca` = round(liquidez_seca, 3),
        `Liquidez Imediata` = round(liquidez_imediata, 3)
      )
  }, striped = TRUE, hover = TRUE, spacing = "s")
  
  # ---- Gráficos ----
  
  # 1) Vendas
  output$plot_vendas_ebit <- renderPlot({
    df <- indicadores() %>%
      select(ANO, Vendas = receita) %>%
      mutate(Vendas = ifelse(is.finite(Vendas), Vendas, NA_real_)) %>%
      filter(!is.na(Vendas))
    
    ggplot(df, aes(x = ANO, y = Vendas)) +
      geom_line() +
      geom_point() +
      labs(
        title = paste0("Evolução das Vendas (", unidade_label(), ")"),
        x = "Ano",
        y = ""
      )
  })
  
  # 2) Margens 
  output$plot_margens <- renderPlot({
    df <- indicadores() %>%
      select(
        ANO,
        `Margem Bruta` = margem_bruta,
        `Margem Operacional (EBIT/Receita)` = margem_operacional,
        `Margem Líquida` = margem_liquida
      ) %>%
      pivot_longer(-ANO, names_to = "indicador", values_to = "valor") %>%
      filter(!is.na(valor), is.finite(valor))
    
    ggplot(df, aes(x = ANO, y = valor, color = indicador)) +
      geom_line() +
      geom_point() +
      scale_y_continuous(labels = function(x) paste0(round(100 * x, 1), "%")) +
      labs(title = "Margens", x = "Ano", y = "", color = "Indicador")
  })
  
  # 3) Rentabilidade
  output$plot_rentab <- renderPlot({
    df <- indicadores() %>%
      select(ANO, ROA = roa, ROE = roe) %>%
      pivot_longer(-ANO, names_to = "indicador", values_to = "valor") %>%
      filter(!is.na(valor), is.finite(valor))
    
    ggplot(df, aes(x = ANO, y = valor, color = indicador)) +
      geom_line() +
      geom_point() +
      scale_y_continuous(labels = function(x) paste0(round(100 * x, 1), "%")) +
      labs(title = "Rentabilidade", x = "Ano", y = "", color = "Indicador")
  })
  
  # 4) Liquidez
  output$plot_liq <- renderPlot({
    df <- indicadores() %>%
      select(
        ANO,
        `Liquidez Corrente` = liquidez_corrente,
        `Liquidez Seca` = liquidez_seca,
        `Liquidez Imediata` = liquidez_imediata
      ) %>%
      pivot_longer(-ANO, names_to = "indicador", values_to = "valor") %>%
      filter(!is.na(valor), is.finite(valor))
    
    ggplot(df, aes(x = ANO, y = valor, color = indicador)) +
      geom_line() +
      geom_point() +
      labs(title = "Liquidez", x = "Ano", y = "", color = "Indicador")
  })
  
  # 5) Endividamento/Alavancagem
  output$plot_endiv <- renderPlot({
    df <- indicadores() %>%
      select(
        ANO,
        `Dívida Líq/PL` = divida_liq_pl,
        `Endividamento Total (Dívida/Ativo)` = endividamento_total,
        `Endividamento CP (PC/Ativo)` = endividamento_curto_prazo
      ) %>%
      pivot_longer(-ANO, names_to = "indicador", values_to = "valor") %>%
      filter(!is.na(valor), is.finite(valor))
    
    ggplot(df, aes(x = ANO, y = valor, color = indicador)) +
      geom_line() +
      geom_point() +
      scale_y_continuous(labels = function(x) paste0(round(100 * x, 1), "%")) +
      labs(title = "Alavancagem / Endividamento", x = "Ano", y = "", color = "Indicador")
  })
  
  # Aba Fórmulas
  output$tbl_formulas <- renderTable({
    FORMULAS
  }, striped = TRUE, hover = TRUE, spacing = "s")
}

shinyApp(ui, server)
