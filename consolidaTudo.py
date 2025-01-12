import pandas as pd

# Carregar os arquivos
liquidacao04 = pd.read_csv('/home/felipe-ribeiro/Projetos/contadoria/data_custas_liquidacao/Central de Contadoria - TJPE - Núcleos de liquidação - Consolidação.csv')
custas04 = pd.read_csv('/home/felipe-ribeiro/Projetos/contadoria/data_custas_liquidacao/Central de Contadoria - TJPE - Núcleos de custas - Consolidação.csv')
liquidacao = pd.read_csv('/home/felipe-ribeiro/Projetos/contadoria/data_custas_liquidacao/Central de Contadoria - TJPE - Núcleos de liquidação - CONSOLIDACAO.csv') 
custas = pd.read_csv('/home/felipe-ribeiro/Projetos/contadoria/data_custas_liquidacao/Central de Contadoria - TJPE - Núcleos de custas - CONSOLIDACAO.csv')
#tempo_real = pd.read_excel('/home/felipe-ribeiro/Projetos/contadoria/data_transform/Consolidado.xlsx')

# Selecionar colunas específicas de liquidacao
liquidacao04 = liquidacao04[['Núcleo', 'Posição Geral', 'Posição Prioridade',
       'Número do processo', 'Vara', 'Data Remessa Contadoria',
        'Prioridade', 'Calculista', 'Data da atribuição', 'Cumprimento', 'Data da conclusão', 
       'Valor Total Devido - Custas', 'Observações']]

# Selecionar colunas específicas de custas
custas04 = custas04[['Núcleo', 'Posição Geral', 'Posição Prioridade',
       'Número do processo', 'Vara', 'Data Remessa Contadoria',
        'Prioridade', 'Calculista', 'Data da atribuição', 'Cumprimento', 'Data da conclusão', 
       'Valor Total Devido - Custas', 'Observações']]

#Remove os registros onde 'Cumprimento' é 'Pendente' do DataFrame custas04
custas04 = custas04[custas04['Cumprimento'] != 'Pendente']

# Remove os registros onde 'Cumprimento' é 'Pendente' do DataFrame liquidacao04
liquidacao04 = liquidacao04[liquidacao04['Cumprimento'] != 'Pendente']


# Mesclar liquidacao e custas
consolidacao = pd.concat([liquidacao04, custas04, liquidacao, custas], axis=0, ignore_index=True)

# Primeiro, limpar possíveis caracteres especiais como R$ e vírgulas
consolidacao['Valor Total Devido - Custas'] = consolidacao['Valor Total Devido - Custas'].str.replace('R$', '').str.replace('.', '').str.replace(',', '.').astype(float)

# Primeiro, converter as colunas de data para datetime
consolidacao['Data Remessa Contadoria'] = pd.to_datetime(consolidacao['Data Remessa Contadoria'], format='%d/%m/%Y', errors='coerce')
consolidacao['Data da atribuição'] = pd.to_datetime(consolidacao['Data da atribuição'], format='%d/%m/%Y', errors='coerce')
consolidacao['Data da conclusão'] = pd.to_datetime(consolidacao['Data da conclusão'], format='%d/%m/%Y', errors='coerce')

# Data atual
data_atual = pd.Timestamp.now()

# Criar coluna 'Tempo na Contadoria'
def calcular_tempo_contadoria(row):
    if row['Cumprimento'] == 'Pendente':
        if pd.notnull(row['Data Remessa Contadoria']):
            return (data_atual - row['Data Remessa Contadoria']).days
    else:
        if pd.notnull(row['Data Remessa Contadoria']) and pd.notnull(row['Data da conclusão']):
            return (row['Data da conclusão'] - row['Data Remessa Contadoria']).days
    return None

# Criar coluna 'Tempo com o Contador'
def calcular_tempo_contador(row):
    if row['Cumprimento'] == 'Pendente':
        if pd.notnull(row['Data da atribuição']):
            return (data_atual - row['Data da atribuição']).days
    else:
        if pd.notnull(row['Data da atribuição']) and pd.notnull(row['Data da conclusão']):
            return (row['Data da conclusão'] - row['Data da atribuição']).days
    return None

# Criar coluna 'Meta'
def definir_meta(cumprimento):
    status_meta = ['Cálculo realizado', 'Cálculo atualizado', 
                  'Devolvido: ausência de documentos para os cálculos',
                  'Devolvido: ausência de parâmetros',
                  'Devolvido: esclarecimento realizado']
    return 1 if cumprimento in status_meta else 0

# Aplicar as funções para criar as novas colunas
consolidacao['Tempo na Contadoria'] = consolidacao.apply(calcular_tempo_contadoria, axis=1)
consolidacao['Tempo com o Contador'] = consolidacao.apply(calcular_tempo_contador, axis=1)
consolidacao['Meta'] = consolidacao['Cumprimento'].apply(definir_meta)


# Verificar os resultados
print("\nPrimeiras linhas com as novas colunas:")
print(consolidacao[['Data Remessa Contadoria', 'Data da atribuição', 'Data da conclusão', 
                   'Cumprimento', 'Tempo na Contadoria', 'Tempo com o Contador', 'Meta']].head())

# Mostrar algumas estatísticas
print("\nEstatísticas das novas colunas:")
print("\nTempo na Contadoria:")
print(consolidacao['Tempo na Contadoria'].describe())
print("\nTempo com o Contador:")
print(consolidacao['Tempo com o Contador'].describe())
print("\nTotal de processos que atingiram a meta:", consolidacao['Meta'].sum())
# Exibir o número de linhas do DataFrame consolidado atualizado
print(f"Número de linhas em consolidacao após adição: {len(consolidacao)}")

# Contar a quantidade de processos onde 'Cumprimento' é 'Pendente'
pendentes_count = len(consolidacao[consolidacao['Cumprimento'] == 'Pendente'])
print(f"Quantidade de processos pendentes: {pendentes_count}")

# Converter colunas de data para string no formato brasileiro
date_columns = ['Data Remessa Contadoria', 'Data da atribuição', 'Data da conclusão']
for col in date_columns:
    consolidacao[col] = consolidacao[col].dt.strftime('%d/%m/%Y').str.replace("'", "")

# Verificar o formato
print(consolidacao[date_columns].head())

# Converter colunas numéricas para o formato correto
consolidacao['Tempo na Contadoria'] = consolidacao['Tempo na Contadoria'].astype('Int64')
consolidacao['Tempo com o Contador'] = consolidacao['Tempo com o Contador'].astype('Int64')

# Criar um dicionário com as substituições para Contadoria de Cálculos Judiciais
substituicoes_ccj = {
    '1ª CONTADORIA DE CÁLCULOS JUDICIAIS': '1ª CCJ',
    '2ª CONTADORIA DE CÁLCULOS JUDICIAIS': '2ª CCJ',
    '3ª CONTADORIA DE CÁLCULOS JUDICIAIS': '3ª CCJ',
    '4ª CONTADORIA DE CÁLCULOS JUDICIAIS': '4ª CCJ',
    '5ª CONTADORIA DE CÁLCULOS JUDICIAIS': '5ª CCJ',
    '6ª CONTADORIA DE CÁLCULOS JUDICIAIS': '6ª CCJ'
}

# Criar um dicionário com as substituições para Contadoria de Custas
substituicoes_cc = {
    '1ª CONTADORIA DE CUSTAS': '1ª CC',
    '2ª CONTADORIA DE CUSTAS': '2ª CC',
    '3ª CONTADORIA DE CUSTAS': '3ª CC',
    '4ª CONTADORIA DE CUSTAS': '4ª CC',
    '5ª CONTADORIA DE CUSTAS': '5ª CC',
    '6ª CONTADORIA DE CUSTAS': '6ª CC',
    '7ª CONTADORIA DE CUSTAS': '7ª CC'
}

# Combinar os dois dicionários
todas_substituicoes = {**substituicoes_ccj, **substituicoes_cc}

# Fazer as substituições
consolidacao['Núcleo'] = consolidacao['Núcleo'].replace(todas_substituicoes)

# Verificar o resultado
print("\nValores únicos na coluna Núcleo após as substituições:")
print(consolidacao['Núcleo'].unique())
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
        
        def format_value(value):
            if pd.isna(value):
                return ""
            if isinstance(value, pd.Timestamp):
                return value.strftime('%d/%m/%Y')
            return value

        # Antes de enviar para o Google Sheets
        values = [[format_value(val) for val in row] for row in df.values.tolist()]
        values.insert(0, df.columns.tolist())
            
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