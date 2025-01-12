import pandas as pd

consolidacao = pd.read_csv('consolidacao.csv')
meta2024 = pd.read_csv('Metas - 2024.csv')
meta2025 = pd.read_csv('Metas - 2025.csv')

#mesclar as duas metas
meta = pd.concat([meta2024, meta2025], axis=0, ignore_index=True)

#salvar as meta em um csv
meta.to_csv('meta.csv', index=False, encoding='utf-8')


# Convertendo as datas para datetime
meta['Data'] = pd.to_datetime(meta['Data'], format='%d/%m/%Y')
consolidacao['Data da conclusão'] = pd.to_datetime(consolidacao['Data da conclusão'], format='%d/%m/%Y')

# Agrupando a tabela2 para somar as metas por Núcleo, Calculista e data
metas_realizadas = consolidacao.groupby([ 'Calculista', 'Data da conclusão'])['Meta'].sum().reset_index()
metas_realizadas = metas_realizadas.rename(columns={'Data da conclusão': 'Data'})

# Atualizando a tabela1 com as metas realizadas
for index, row in meta.iterrows():
    meta_realizada = metas_realizadas[
        (metas_realizadas['Calculista'] == row['Calculista']) & 
        (metas_realizadas['Data'] == row['Data'])
    ]['Meta'].values
    
    if len(meta_realizada) > 0:
        meta.at[index, 'META REALIZADA'] = meta_realizada[0]
    else:
        meta.at[index, 'META REALIZADA'] = 0  # Coloca 0 quando não houver meta realizada naquele dia

# Convertendo as datas de volta para string no formato brasileiro
meta['Data'] = meta['Data'].dt.strftime('%d/%m/%Y')

# Salvando a tabela1 atualizada
meta.to_csv('meta_atualizada.csv', index=False)
meta.to_excel('meta_atualizada.xlsx', index=False)

# Mostrando algumas verificações
print("\nPrimeiras linhas da tabela atualizada:")
print(meta.head())

print("\nTotal de metas realizadas por Núcleo:")
print(meta.groupby('Núcleo')['META REALIZADA'].sum())

print("\nMédia diária de metas realizadas por Calculista:")
print(meta.groupby('Calculista')['META REALIZADA'].mean())

def load_consolidacao():
    import os.path
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from googleapiclient.discovery import build

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
    SPREADSHEET_ID = '1ylBBejID_uhguxZVr-XemB5hgN4ABrv5bEgosNObNTE'

    # Leitura do arquivo Excel local
    df = pd.read_excel('meta_atualizada.xlsx')
    
    # Tratar NaNs
    df = df.fillna("")
    values = [df.columns.values.tolist()] + df.values.tolist()
    
    # Determinar o intervalo
    max_row = len(values)
    num_cols = len(values[0])
    max_col = chr(64 + num_cols).upper()
    range_name = f"Metas!A1:{max_col}{max_row}"  # Nome da aba específico
    
    try:
        # Limpar o conteúdo existente
        service.spreadsheets().values().clear(
            spreadsheetId=SPREADSHEET_ID, 
            range=range_name
        ).execute()
        
        # Atualizar com novos dados
        result = service.spreadsheets().values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=range_name,
            valueInputOption='RAW',
            body={'values': values}
        ).execute()

        print(f'Atualizadas {result.get("updatedCells")} células na planilha.')
        
    except Exception as e:
        print(f"Erro ao atualizar planilha: {str(e)}")
        print("Tentando criar nova aba...")
        
        try:
            body = {
                'requests': [{
                    'addSheet': {
                        'properties': {
                            'title': 'Metas'
                        }
                    }
                }]
            }
            service.spreadsheets().batchUpdate(
                spreadsheetId=SPREADSHEET_ID,
                body=body
            ).execute()
            
            # Tentar atualizar novamente
            result = service.spreadsheets().values().update(
                spreadsheetId=SPREADSHEET_ID,
                range=range_name,
                valueInputOption='RAW',
                body={'values': values}
            ).execute()
            
            print(f'Nova aba criada e {result.get("updatedCells")} células atualizadas.')
            
        except Exception as e:
            print(f"Erro ao criar nova aba: {str(e)}")

# Chamar a função
load_consolidacao()