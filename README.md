# Demonstração — Simulador da reforma 2027 para serviços no Simples

> Projeto de portfólio de **Victória Pedrosa**. **Demonstração** de simulador da reforma 2027 para serviços no Simples — versão com dados fictícios (nomes, CNPJs, e-mails e IDs internos substituídos).

## Problema de negócio
Empresas de serviços no Simples precisam escolher entre DAS integral e Simples híbrido em 2027; exige extratos, classificação e comunicação ao cliente.

## Antes x depois
| | Antes | Depois |
|---|---|---|
| Como é feito | Levantamento manual de extratos e simulação em planilhas separadas. | Painel registra empresas e apurações, classifica CNAE x item x NBS com apoio do Gemini, calcula os cenários e controla comunicados, com log de correções. |

## Ganho
- Carteira inteira simulada e decisões registradas por cliente.

## Tecnologias
APIs REST, Gemini API, Google Apps Script, Google Sheets, HTML/JavaScript, Web App (HtmlService)

## Arquivos
- `Classificacao.gs`
- `Codigo.gs`
- `Gemini.gs`
- `Motor.html`
- `Painel.html`

## Como usar
Crie um projeto no Google Apps Script, copie os arquivos `.gs`/`.html` e configure as Propriedades do script indicadas no código.

## Autora
Victória Pedrosa — Product Owner do Time de IA, automação de processos contábeis e fiscais.
