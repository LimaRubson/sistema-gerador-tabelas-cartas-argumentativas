import streamlit as st
import pandas as pd
from google.oauth2 import service_account
from googleapiclient.discovery import build
import json
import os

st.set_page_config(page_title="Gerador de Tabela Dinâmica - Carta Argumentativa", layout="wide")
st.title("📊 Gerador de Tabela Dinâmica - Carta Argumentativa")

# --- Upload do arquivo Excel ---
st.sidebar.header("📁 Upload de Arquivos")
excel_file = st.sidebar.file_uploader("Arquivo Excel (.xlsx)", type=["xlsx"])

# Carrega as credenciais do secrets
service_account_info = st.secrets["GOOGLE_SERVICE_ACCOUNT"]
credentialsl = service_account.Credentials.from_service_account_info(service_account_info)

# Verifica se o arquivo de credenciais existe na raiz
CREDENTIALS_PATH = "credentials.json"
if not os.path.exists(CREDENTIALS_PATH):
    st.error("❌ O arquivo 'credentials.json' não foi encontrado na raiz do projeto.")
else:
    if excel_file:
        ARQUIVO_ORIGEM = excel_file
        ABA1 = 'Worksheet'
        ABA2 = 'Worksheet2'

        COLUNAS_ESPERADAS = [
            'Redação ID', 'Nome do Prompt', 'Prompt', 'Texto da Redação', 'Tema', 
            'Competência 1 - IA', 'Competência 1 - Humano', 'Divergencia Competência 1', 'Modulo Divergencia Competência 1',
            'Competência 2 - IA', 'Competência 2 - humano', 'Divergencia Competência 2', 'Modulo Divergencia Competência 2',
            'Competência 3 - IA', 'Competência 3 - Humano', 'Divergencia Competência 3', 'Modulo Divergencia Competência 3',
            'Competência 4 - IA', 'Competência 4 - Humano', 'Divergencia Competência 4', 'Modulo Divergencia Competência 4',
            'Competência 5 - IA', 'Competência 5 - Humano', 'Divergencia Competência 5', 'Modulo Divergencia Competência 5',
            'Nota - IA', 'Nota - Humano', 'Divergencia Nota', 'Modulo Divergencia Nota',
            'Feedback Competência 1', 'Feedback Competência 2', 'Feedback Competência 3', 'Feedback Competência 4',
            'Feedback Competência 5', 'Feedback Geral'
        ]

        SPREADSHEET_TITLE = 'Planilha com Tabela Dinâmica'
        DATA_SHEET_TITLE = 'Dados'
        PIVOT_SHEET_TITLE = 'Tabela Dinâmica'
        DIVERGENCIA_SHEET_TITLE = 'Divergência'
        DIVERGENCIA_TOTAL_SHEET_TITLE = 'Divergência Total'

        with st.spinner("🔄 Lendo dados do Excel..."):
            df1 = pd.read_excel(ARQUIVO_ORIGEM, sheet_name=ABA1, skiprows=0)
            df2 = pd.read_excel(ARQUIVO_ORIGEM, sheet_name=ABA2, skiprows=1, header=None)
            df2.columns = df1.columns
            df_combinado = pd.concat([df1, df2], ignore_index=True)
            df_combinado = df_combinado[COLUNAS_ESPERADAS].fillna("")

        # Autenticação
        SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
        credentialsl.with_scopes(SCOPES)
        sheets_service = build('sheets', 'v4', credentials=credentialsl)
        drive_service = build('drive', 'v3', credentials=credentialsl)

        with st.spinner("📁 Criando planilha no Google Sheets..."):
            spreadsheet_body = {
                'properties': {'title': SPREADSHEET_TITLE},
                'sheets': [
                    {'properties': {'title': DATA_SHEET_TITLE}},
                    {'properties': {'title': PIVOT_SHEET_TITLE}},
                ]
            }

            spreadsheet = sheets_service.spreadsheets().create(body=spreadsheet_body, fields='spreadsheetId').execute()
            spreadsheet_id = spreadsheet['spreadsheetId']

            drive_service.permissions().create(
                fileId=spreadsheet_id,
                body={'role': 'writer', 'type': 'anyone'},
                fields='id'
            ).execute()

            values = [list(df_combinado.columns)] + df_combinado.values.tolist()
            sheets_service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id,
                range=f'{DATA_SHEET_TITLE}!A1',
                valueInputOption='RAW',
                body={'values': values}
            ).execute()

            spreadsheet_info = sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
            sheets = spreadsheet_info['sheets']
            sheet_id_dados = next(s['properties']['sheetId'] for s in sheets if s['properties']['title'] == DATA_SHEET_TITLE)
            sheet_id_pivot = next(s['properties']['sheetId'] for s in sheets if s['properties']['title'] == PIVOT_SHEET_TITLE)

        with st.spinner("📊 Criando Tabela Dinâmica..."):
            requests = [{
                "updateCells": {
                    "rows": [{
                        "values": [{
                            "pivotTable": {
                                "source": {
                                    "sheetId": sheet_id_dados,
                                    "startRowIndex": 0,
                                    "startColumnIndex": 0,
                                    "endRowIndex": len(values),
                                    "endColumnIndex": len(values[0])
                                },
                                "rows": [{
                                    "sourceColumnOffset": 0,
                                    "showTotals": True,
                                    "sortOrder": "ASCENDING"
                                }],
                                "columns": [{
                                    "sourceColumnOffset": 1,
                                    "showTotals": False,
                                    "sortOrder": "ASCENDING"
                                }],
                                "values": [
                                    {"summarizeFunction": "SUM", "sourceColumnOffset": 2},
                                    {"summarizeFunction": "SUM", "sourceColumnOffset": 3},
                                    {"summarizeFunction": "SUM", "sourceColumnOffset": 4},
                                    {"summarizeFunction": "SUM", "sourceColumnOffset": 5}
                                ],
                                "criteria": {
                                    "1": {
                                        "visibleValues": list(df_combinado['Nome do Prompt'].unique())
                                    }
                                }
                            }
                        }]
                    }],
                    "start": {"sheetId": sheet_id_pivot, "rowIndex": 2, "columnIndex": 0},
                    "fields": "pivotTable"
                }
            }]

            sheets_service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body={"requests": requests}
            ).execute()

        with st.spinner("🧮 Calculando divergências..."):
            prompts = df_combinado['Nome do Prompt'].unique()
            df_dict = {p: df_combinado[df_combinado['Nome do Prompt'] == p].set_index('Redação ID') for p in prompts}

            PROMPT_GPT, PROMPT_NILMA = prompts[0], prompts[1]
            ids_comuns = df_dict[PROMPT_GPT].index.intersection(df_dict[PROMPT_NILMA].index)

            df_div = df_dict[PROMPT_GPT].loc[ids_comuns, ["Competência 1 - IA", "Competência 2 - IA", "Competência 3 - IA", "Competência 4 - IA"]] - \
                     df_dict[PROMPT_NILMA].loc[ids_comuns, ["Competência 1 - IA", "Competência 2 - IA", "Competência 3 - IA", "Competência 4 - IA"]]
            df_div = df_div.reset_index().round(2)

            sheets_service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body={"requests": [{"addSheet": {"properties": {"title": DIVERGENCIA_SHEET_TITLE}}}]}
            ).execute()

            sheets_service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id,
                range=f'{DIVERGENCIA_SHEET_TITLE}!A1',
                valueInputOption='RAW',
                body={'values': [list(df_div.columns)] + df_div.values.tolist()}
            ).execute()

            df_total_gpt = df_dict[PROMPT_GPT].loc[ids_comuns, ["Competência 1 - IA", "Competência 2 - IA", "Competência 3 - IA", "Competência 4 - IA"]].sum(axis=1)
            df_total_nilma = df_dict[PROMPT_NILMA].loc[ids_comuns, ["Competência 1 - IA", "Competência 2 - IA", "Competência 3 - IA", "Competência 4 - IA"]].sum(axis=1)

            df_div_total = pd.DataFrame({
                'Redação ID': ids_comuns,
                f'Soma Total - {PROMPT_GPT}': df_total_gpt,
                f'Soma Total - {PROMPT_NILMA}': df_total_nilma,
                'Divergência (GPT - NILMA)': (df_total_gpt - df_total_nilma).round(2)
            }).reset_index(drop=True)

            sheets_service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body={"requests": [{"addSheet": {"properties": {"title": DIVERGENCIA_TOTAL_SHEET_TITLE}}}]}
            ).execute()

            sheets_service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id,
                range=f'{DIVERGENCIA_TOTAL_SHEET_TITLE}!A1',
                valueInputOption='RAW',
                body={'values': [list(df_div_total.columns)] + df_div_total.values.tolist()}
            ).execute()

        st.success("✅ Processo concluído com sucesso!")
        st.markdown(f"🔗 **[Acesse a planilha gerada aqui](https://docs.google.com/spreadsheets/d/{spreadsheet_id})**")

    else:
        st.warning("⚠️ Envie o Excel para continuar.")
