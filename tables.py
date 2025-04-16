import streamlit as st

def table_div_comp(sheet_id_dados, df_combinado, sheet_id_pivot_div_comp, values, sheets_service, spreadsheet_id):
    with st.spinner("📊 Criando Tabela Dinâmica - Divergência entre Competência IA e HU"):
        requests = [
            # Primeira Tabela - Módulo Divergência Competência 1
            {
                "updateCells": {
                    "rows": [{
                        "values": [{
                            "pivotTable": {
                                "source": {
                                    "sheetId": sheet_id_dados,
                                    "startRowIndex": 0,
                                    "startColumnIndex": 0,
                                    "endRowIndex": len(df_combinado),
                                    "endColumnIndex": len(df_combinado.columns)
                                },
                                "rows": [{
                                    "sourceColumnOffset": 8,  # Coluna usada como linha
                                    "showTotals": True,
                                    "sortOrder": "ASCENDING"
                                }],
                                "columns": [],
                                "values": [
                                    {"summarizeFunction": "COUNT", "sourceColumnOffset": 8},
                                    {"summarizeFunction": "COUNT", "sourceColumnOffset": 8}
                                ],
                                "criteria": {
                                    "1": {
                                        "visibleValues": list(df_combinado['Nome do Prompt'].unique())
                                    }
                                }
                            }
                        }]
                    }],
                    "start": {"sheetId": sheet_id_pivot_div_comp, "rowIndex": 2, "columnIndex": 0},
                    "fields": "pivotTable"
                }
            },

            # Segunda Tabela - Módulo Divergência Competência 2
            {
                "updateCells": {
                    "rows": [{
                        "values": [{
                            "pivotTable": {
                                "source": {
                                    "sheetId": sheet_id_dados,
                                    "startRowIndex": 0,
                                    "startColumnIndex": 0,
                                    "endRowIndex": len(df_combinado),
                                    "endColumnIndex": len(df_combinado.columns)
                                },
                                "rows": [{
                                    "sourceColumnOffset": 12,  # Coluna usada como linha
                                    "showTotals": True,
                                    "sortOrder": "ASCENDING"
                                }],
                                "columns": [],
                                "values": [
                                    {"summarizeFunction": "COUNT", "sourceColumnOffset": 12},
                                    {"summarizeFunction": "COUNT", "sourceColumnOffset": 12}
                                ],
                                "criteria": {
                                    "1": {
                                        "visibleValues": list(df_combinado['Nome do Prompt'].unique())
                                    }
                                }
                            }
                        }]
                    }],
                    "start": {"sheetId": sheet_id_pivot_div_comp, "rowIndex": 12, "columnIndex": 0},
                    "fields": "pivotTable"
                }
            },

            # Terceira Tabela - Módulo Divergência Competência 3
            {
                "updateCells": {
                    "rows": [{
                        "values": [{
                            "pivotTable": {
                                "source": {
                                    "sheetId": sheet_id_dados,
                                    "startRowIndex": 0,
                                    "startColumnIndex": 0,
                                    "endRowIndex": len(df_combinado),
                                    "endColumnIndex": len(df_combinado.columns)
                                },
                                "rows": [{
                                    "sourceColumnOffset": 16,  # Coluna usada como linha
                                    "showTotals": True,
                                    "sortOrder": "ASCENDING"
                                }],
                                "columns": [],
                                "values": [
                                    {"summarizeFunction": "COUNT", "sourceColumnOffset": 16},
                                    {"summarizeFunction": "COUNT", "sourceColumnOffset": 16}
                                ],
                                "criteria": {
                                    "1": {
                                        "visibleValues": list(df_combinado['Nome do Prompt'].unique())
                                    }
                                }
                            }
                        }]
                    }],
                    "start": {"sheetId": sheet_id_pivot_div_comp, "rowIndex": 22, "columnIndex": 0},
                    "fields": "pivotTable"
                }
            },

            # Quarta Tabela - Módulo Divergência Competência 4
            {
                "updateCells": {
                    "rows": [{
                        "values": [{
                            "pivotTable": {
                                "source": {
                                    "sheetId": sheet_id_dados,
                                    "startRowIndex": 0,
                                    "startColumnIndex": 0,
                                    "endRowIndex": len(df_combinado),
                                    "endColumnIndex": len(df_combinado.columns)
                                },
                                "rows": [{
                                    "sourceColumnOffset": 20,  # Coluna usada como linha
                                    "showTotals": True,
                                    "sortOrder": "ASCENDING"
                                }],
                                "columns": [],
                                "values": [
                                    {"summarizeFunction": "COUNT", "sourceColumnOffset": 20},
                                    {"summarizeFunction": "COUNT", "sourceColumnOffset": 20}
                                ],
                                "criteria": {
                                    "1": {
                                        "visibleValues": list(df_combinado['Nome do Prompt'].unique())
                                    }
                                }
                            }
                        }]
                    }],
                    "start": {"sheetId": sheet_id_pivot_div_comp, "rowIndex": 32, "columnIndex": 0},
                    "fields": "pivotTable"
                }
            }
        ]

        sheets_service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={"requests": requests}
        ).execute()
