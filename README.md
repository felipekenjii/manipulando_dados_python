# Extração e Automação de Dados de Holerites com Python

Este projeto foi desenvolvido durante o minicurso **"Automação de Processos com Python: Extração e Manipulação de Dados em Massa"**, realizado na **Semana da SETEC** na FATEC.  
O objetivo foi **automatizar a leitura de arquivos PDF contendo holerites** (fichas financeiras) e **extrair informações relevantes**, como códigos, datas e valores, organizando tudo em uma planilha Excel.

---

## 🚀 Objetivo do Projeto

A proposta foi demonstrar como o Python pode ser utilizado para **automatizar tarefas repetitivas de manipulação de dados**, extraindo informações estruturadas de documentos PDF e exportando-as de forma limpa e organizada para planilhas.

---

## 🧠 O que o script faz

O script `extracao_holerite.py` executa as seguintes etapas:

1. **Abre o arquivo PDF** com o módulo [`pdfplumber`](https://github.com/jsvine/pdfplumber);
2. **Lê o texto de cada página** e identifica padrões como:
   - Código do lançamento;
   - Nome do provento/desconto;
   - Data de pagamento;
   - Valor de referência e valor final;
3. **Extrai os dados usando expressões regulares (Regex)**;
4. **Converte os valores e datas para os formatos corretos**;
5. **Exporta os dados para uma planilha Excel (`.xlsx`)**;
6. **Formata automaticamente as colunas de data** na planilha.

---

## 🧩 Tecnologias Utilizadas

- **Python**
- **Bibliotecas**:
  - `pdfplumber` → leitura e extração de texto de PDFs  
  - `re` → uso de expressões regulares  
  - `pandas` → manipulação de dados tabulares  
  - `openpyxl` → criação e formatação de arquivos Excel  
  - `tqdm` → barra de progresso visual no terminal  

---

## 💡 Aprendizados

Durante o desenvolvimento deste projeto, foram exploradas técnicas essenciais de automação com Python, incluindo:
- Extração de texto de PDFs com pdfplumber;
- Uso de Regex para capturar padrões específicos;
- Limpeza e conversão de dados com pandas;
- Geração e formatação de planilhas Excel;
- Implementação de feedback visual com tqdm.

---

## 🎓 Créditos

- Projeto desenvolvido por Gustavo Rodrigues Moreira durante o minicurso da Semana da SETEC – FATEC.
- Curso ministrado sobre Automação de Processos com Python: Extração e Manipulação de Dados em Massa.

---
