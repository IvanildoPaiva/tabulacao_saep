# 📊 Sistema de Tabulação Automatizada - SAEP

> Uma ferramenta web robusta para processamento, correção e geração de relatórios de desempenho individual do SAEP, preservando 100% da inteligência e formatação das planilhas Excel originais.

![Status do Projeto](https://img.shields.io/badge/Status-Finalizado_v1.0-success)
![Tecnologia](https://img.shields.io/badge/Tech-HTML5_|_JS_|_CSS3-blue)
![Engine](https://img.shields.io/badge/Engine-XlsxPopulate-orange)

## 🎯 O Problema
Professores  precisavam tabular manualmente os dados brutos do sistema SAEP para uma planilha de diagnóstico visual. O processo manual gerava erros de formatação, quebrava fórmulas do Excel e resultava em gráficos vazios (`#DIV/0!`) devido a inconsistências nos dados de entrada (como espaços extras em códigos "C3 ").

## 🚀 A Solução
Este sistema roda inteiramente no navegador (Client-side), lê os dados brutos, aplica correções lógicas e preenche a planilha mestre "cirurgicamente", mantendo gráficos, macros e formatações condicionais intactas.

### ✨ Principais Funcionalidades

* **Preservação Total:** Utiliza a biblioteca `xlsx-populate` para editar o Excel sem reescrever seu XML, garantindo que gráficos e fórmulas complexas não sejam perdidos.
* **Correção de Dados (Sanitização):** Remove automaticamente espaços fantasmas e caracteres inválidos dos códigos de capacidade (ex: converte `"C3 "` para `"C3"`), permitindo que o `PROCV` e `SE` do Excel funcionem.
* **Auto-Preenchimento de Descrições:** Cruza o código da capacidade (ex: C1) com o texto descritivo no arquivo de dados e preenche automaticamente a aba de Diagnóstico.
* **Recálculo Forçado:** Configura o arquivo para forçar o Excel a recalcular todas as fórmulas (`fullCalcOnLoad`) ao abrir, eliminando erros de exibição inicial.
* **Interface Amigável:** Design limpo, responsivo e com feedback visual de progresso.

---

## 🛠️ Tecnologias Utilizadas

* **HTML5 & CSS3:** Estrutura semântica e estilização moderna (Flexbox/Grid).
* **JavaScript (ES6+):** Lógica de processamento assíncrono.
* **[SheetJS (xlsx)](https://sheetjs.com/):** Para leitura ultrarrápida dos dados brutos.
* **[xlsx-populate](https://github.com/dtjohnson/xlsx-populate):** Para escrita segura e preservação de objetos do Excel.
* **[FileSaver.js](https://github.com/eligrey/FileSaver.js):** Para gerenciar o download do arquivo gerado no navegador.

---

