import pandas as pd

# Carregar os arquivos
liquidacao = pd.read_csv('/home/felipe-ribeiro/Projetos/contadoria/data_custas_liquidacao/Central de Contadoria - TJPE - Núcleos de liquidação - Consolidação.csv')
custas = pd.read_csv('/home/felipe-ribeiro/Projetos/contadoria/data_custas_liquidacao/Central de Contadoria - TJPE - Núcleos de custas - Consolidação.csv')
tempo_real = pd.read_excel('/home/felipe-ribeiro/Projetos/contadoria/data_transform/Consolidado.xlsx')

# Selecionar colunas específicas de liquidacao
liquidacao = liquidacao[['Tipo', 'Núcleo', 'Posição Geral', 'Posição Prioridade',
       'Número do processo', 'Vara', 'Data Remessa Contadoria',
        'Prioridade', 'Calculista', 'Data da atribuição', 'Cumprimento', 'Data da conclusão', 
       'Valor Total Devido - Custas', 'Observações']]

# Mesclar liquidacao e custas
consolidacao = pd.concat([liquidacao, custas], axis=0, ignore_index=True)

# Renomear colunas de tempo_real para corresponder às de consolidacao
tempo_real.rename(columns={
    'nucleo': 'Núcleo',
    'processo': 'Número do processo',
    'vara': 'Vara',
    'data': 'Data Remessa Contadoria',
    'prioridades': 'Prioridade',
    'dias': 'Tempo na contadoria'
}, inplace=True)

# Identificar processos que estão com 'Cumprimento' pendente na consolidação
processos_pendentes = consolidacao[consolidacao['Cumprimento'] == 'Pendente']['Número do processo']

# Filtrar tempo_real para incluir apenas processos que não estão na lista de pendentes
novos_processos = tempo_real[~tempo_real['Número do processo'].isin(processos_pendentes)]

# Adicionar a coluna 'Cumprimento' com valor 'Pendente' para novos processos
novos_processos['Cumprimento'] = 'Pendente'

# Concatenar novos processos ao DataFrame consolidacao
consolidacao = pd.concat([consolidacao, novos_processos], axis=0, ignore_index=True)

# Exibir o número de linhas do DataFrame consolidado atualizado
print(f"Número de linhas em consolidacao após adição: {len(consolidacao)}")

# Contar a quantidade de processos onde 'Cumprimento' é 'Pendente'
pendentes_count = len(consolidacao[consolidacao['Cumprimento'] == 'Pendente'])
print(f"Quantidade de processos pendentes: {pendentes_count}")

# Salvar o DataFrame consolidado atualizado
consolidacao.to_csv('consolidacao.csv', index=False, encoding='utf-8')
consolidacao.to_excel('consolidacao.xlsx', sheet_name='consolidacao', index=False)

def load_consolidacao():
    import os.path

    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from googleapiclient.discovery import build
    from googleapiclient.errors import HttpError

    # Autenticação
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    creds = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                "/home/felipe-ribeiro/Projetos/contadoria/credentials.json", SCOPES
            )
            creds = flow.run_local_server(port=0)
        with open("token.json", "w") as token:
            token.write(creds.to_json())

    service = build('sheets', 'v4', credentials=creds)

    # ID da planilha do Google Sheets
    SPREADSHEET_ID = '1awhOgdpa_Kkwsj3NxFnlCVPGJPdM3bR6kasS7alyOVc'

    # Leitura do arquivo XLS local, incluindo todas as abas
    file_path = 'consolidacao.xlsx'
    sheets = pd.read_excel(file_path, sheet_name=None)

    for sheet_name, df in sheets.items():
        # Tratar NaNs
        df = df.fillna("")
        values = [df.columns.values.tolist()] + df.values.tolist()
        
        # Determinar o intervalo exato
        max_row = df.shape[0] + 1
        max_col = chr(65 + df.shape[1] - 1)
        range_name = f"{sheet_name}!A1:{max_col}{max_row}"
    
        body = {'values': values}
        
        # Limpar e atualizar o conteúdo
        service.spreadsheets().values().clear(spreadsheetId=SPREADSHEET_ID, range=range_name).execute()
        result = service.spreadsheets().values().update(
            spreadsheetId=SPREADSHEET_ID, range=range_name,
            valueInputOption='RAW', body=body).execute()

        print(f'{result.get("updatedCells")} células atualizadas na aba {sheet_name}.')

load_consolidacao()
