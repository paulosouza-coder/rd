# 📊 **Integrador RD Station - Dados em Excel e Google Sheets**

Bem-vindo ao **Integrador RD Station**! Este projeto foi desenvolvido para facilitar a extração e organização de dados das APIs do RD Station, permitindo exportar informações valiosas sobre tarefas e negócios diretamente para **Excel** e **Google Sheets**.

✨ **Automatize seu fluxo de dados com facilidade!** ✨

## 🚀 O que faz o projeto?

Este script coleta e processa informações sobre tarefas (`tasks`) e negócios (`deals`) da sua conta RD Station e os exporta para:

- **Excel**: Salve os dados em planilhas bem organizadas.
- **Google Sheets**: Atualize planilhas online automaticamente com as últimas informações.

## 🌟 Funcionalidades

- **Extração de Tarefas (Tasks)**: Coleta informações sobre tarefas, como assunto, status, data de conclusão, usuários e mais.
- **Extração de Negócios (Deals)**: Inclui detalhes de negócios, como valor total, estágio, usuário responsável, e outros campos personalizados.
- **Exportação para Excel e Google Sheets**: Envie os dados diretamente para o Excel ou para uma planilha do Google Sheets, mantendo suas informações organizadas e sempre atualizadas.

## 📦 Pré-requisitos

Antes de rodar o script, você precisará de algumas bibliotecas e configurar credenciais para acessar a API do RD Station e o Google Sheets.

1. **Bibliotecas necessárias**:
   - `pandas`
   - `gspread`
   - `google-auth`
   - `json`
   
   Instale essas dependências usando pip:
   ```bash
   pip install pandas gspread google-auth
   ```

## ☁️ Alternativa: Google Apps Script (sem servidor, sem custo)

Em vez de rodar os scripts Python localmente, é possível manter a planilha
sincronizada automaticamente usando o Google Apps Script, direto na planilha
do Google Sheets — sem depender de máquina local e sem custo.

Os arquivos estão em [`apps_script/`](apps_script/):

1. Na planilha do Google Sheets, abra **Extensões > Apps Script**.
2. Cole o conteúdo de `apps_script/Code.gs` no editor (arquivo `Code.gs`).
3. Rode a função `configurarToken` uma vez para salvar o token da API do RD
   Station (guardado em Script Properties, não fica no código).
4. Rode `sincronizarRDStation` uma vez manualmente para autorizar o script.
5. Rode `criarGatilhoHorario` para agendar a sincronização automática (a cada
   hora).

Recarregue a planilha para ver o menu **RD Station** com essas ações.




