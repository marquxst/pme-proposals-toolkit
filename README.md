# PME Proposals Toolkit

Script em Python que consolida dados de **propostas comerciais PME** (planos de saúde PME/PF e afins) a partir de uma planilha CSV e gera automaticamente um **relatório Excel formatado**, com resumo por operadora, resumo por status e um gráfico — agilizando o acompanhamento que normalmente seria feito manualmente.

Projeto inspirado no dia a dia de emissão e acompanhamento de propostas comerciais PME.

## ✨ Funcionalidades

- Lê um CSV com as propostas (cliente, operadora, vidas, valor estimado, status, data)
- Gera uma planilha Excel com 3 abas:
  - **Propostas**: dados brutos formatados
  - **Resumo por Operadora**: total de propostas, vidas e valor por operadora + gráfico de barras
  - **Resumo por Status**: quantidade e valor total por status (Aprovada, Em análise, Recusada)
- Formatação automática: cabeçalhos coloridos, moeda em R$, colunas ajustadas à largura do conteúdo

## 🛠️ Tecnologias

- Python 3
- pandas
- openpyxl

## ▶️ Como rodar

```bash
git clone https://github.com/marquxst/pme-proposals-toolkit.git
cd pme-proposals-toolkit
pip install -r requirements.txt

python report.py propostas_exemplo.csv relatorio.xlsx
```

O arquivo `relatorio.xlsx` é gerado na pasta atual, pronto para enviar ou anexar em um acompanhamento comercial.

> `propostas_exemplo.csv` contém dados fictícios apenas para demonstração — use sua própria planilha com as mesmas colunas (`cliente`, `operadora`, `vidas`, `valor_estimado`, `status`, `data_envio`).

## 📌 Próximos passos (ideias)

- Ler diretamente de uma planilha do Google Sheets/Excel Online
- Enviar o relatório automaticamente por e-mail ao final da geração
- Dashboard web (Power BI ou HTML) alimentado pelo mesmo CSV
