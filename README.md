# Relatório de Vendas Automatizado com Python

Este script realiza a análise de vendas a partir de uma base de dados em Excel, calcula indicadores por loja e envia automaticamente um relatório por e-mail utilizando o Outlook.

## Funcionalidades

- Leitura de planilha Excel com pandas
- Cálculo de:
  - Faturamento por loja
  - Quantidade de produtos vendidos por loja
  - Ticket médio por loja
- Geração automática de relatório em HTML
- Envio de e-mail com o relatório usando Outlook (via `win32com`)

## Pré-requisitos

- Python instalado
- Pacotes:
  - pandas
  - pywin32
- Microsoft Outlook instalado e configurado

## Como usar

1. Coloque o arquivo `Vendas.xlsx` na mesma pasta do script.
2. Atualize o e-mail do destinatário no campo `mail.To`.
3. Execute o script.

O relatório será enviado automaticamente por e-mail.

---

📌 **Obs.:** Este script é ideal para automatizar tarefas de análise de dados no dia a dia corporativo.

