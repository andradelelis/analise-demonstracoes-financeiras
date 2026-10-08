# Análise de demonstrações financeiras

Aplicativo desenvolvido em R com Shiny para apoiar o ensino e a análise de demonstrações financeiras.

## Acessar o aplicativo

**[Abrir aplicativo de análise de demonstrações financeiras](https://lelis-pedro-andrade.shinyapps.io/app-dfcs/)**

O aplicativo está publicado no shinyapps.io e pode ser utilizado pelo navegador, sem instalar R ou RStudio.

## Objetivo

Oferecer um recurso de apoio ao estudo de Administração Financeira e Análise das Demonstrações Contábeis, para uso em aulas, atividades práticas e análises financeiras.

## Tecnologias

- **R:** linguagem de programação.
- **Shiny:** estrutura para construção do aplicativo interativo.
- **shinyapps.io:** serviço onde o aplicativo está publicado.

## Código-fonte

O arquivo [app.R](app.R) contém a interface, a busca de empresas, a obtenção de dados DFP-CVM pelo pacote GetDFPData2, o cálculo de indicadores e a exportação de demonstrações para Excel.

### Funcionalidades

- Busca de empresas por nome e seleção do código CVM.
- Demonstrações consolidadas ou individuais: DRE, DFC pelo método indireto, BPA, BPP, DMPL e DVA.
- Consulta de até cinco anos e escolha da escala de apresentação.
- Indicadores de margens, rentabilidade, liquidez e endividamento, com gráficos.
- Aba com as fórmulas utilizadas.

### Executar no computador

1. Baixe o repositório em **Code → Download ZIP** e extraia a pasta.
2. Instale R e RStudio.
3. No Console do RStudio, instale os pacotes:

```r
install.packages(c(
  "shiny", "GetDFPData2", "dplyr", "stringr",
  "openxlsx", "tidyr", "ggplot2", "tibble"
))
```

4. Abra o arquivo `app.R` no RStudio e clique em **Run App**. Também é possível executar, informando a pasta extraída:

```r
shiny::runApp("C:/caminho/analise-demonstracoes-financeiras-main")
```

A consulta de dados requer acesso à internet.

### Versão e definições dos indicadores

O código disponibilizado é a cópia local fornecida pelo responsável, com período inicial de 2021 a 2025. Na conferência do aplicativo publicado, o período inicial era 2020 a 2024; a equivalência integral entre as versões não foi confirmada.

ROA e ROE utilizam ativo e patrimônio líquido de fim de exercício, sem médias. O indicador denominado “Endividamento Total” no código corresponde à dívida financeira de curto e longo prazo dividida pelo ativo total. Alguns cálculos tratam contas ausentes como zero.

Alterações neste repositório não atualizam automaticamente o aplicativo no shinyapps.io.

## Responsável pelo projeto

**Lélis Andrade**  
Professor de Finanças — IFMG, Campus Formiga.

[Perfil no GitHub](https://github.com/andradelelis)

## Sugestões e problemas

Para sugerir melhorias ou relatar problemas, abra uma [issue neste repositório](https://github.com/andradelelis/analise-demonstracoes-financeiras/issues).
