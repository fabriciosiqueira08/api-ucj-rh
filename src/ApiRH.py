import requests
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

# Definições
PIPEFY_API_TOKEN = "eyJhbGciOiJIUzUxMiJ9.eyJpc3MiOiJQaXBlZnkiLCJpYXQiOjE3MTE3MTY3MzgsImp0aSI6IjMwYzBiNjMzLTg5ZWEtNDM0ZC1hODMxLTZkYjc3OTM0MTBkNiIsInN1YiI6MjEyODA2LCJ1c2VyIjp7ImlkIjoyMTI4MDYsImVtYWlsIjoicmhAdWNqLmNvbS5iciIsImFwcGxpY2F0aW9uIjozMDAzMzgyMzIsInNjb3BlcyI6W119LCJpbnRlcmZhY2VfdXVpZCI6bnVsbH0.79e7athW43b4WrBvWOsxa4wsIEUbQlVRzdU6rlZ4pmjDB2ABiv8sOyPu0jv18Gj5HCkue4QIMavqAqE2CnMHiQ"
PIPEFY_GRAPHQL_ENDPOINT = 'https://api.pipefy.com/graphql'
EXCEL_FILENAME = 'Dados_Pipe_RH_ENPS.xlsx'

# Cabeçalhos para as colunas do Excel
HEADERS = ['ID do Pipe', 'Nome do Pipe', 'Nome da Fase', 'ID do Card', 'Título do Card', 'Nome do Campo', 'Valor do Campo']

# IDs dos Pipes (substitua pelos IDs reais dos seus pipes)
PIPE_IDS = {
    'RH - E-NPS': '301823995',
    'RH - Matriz de Cursos': '301682389',
    'RH - Painel Controle Membros': '301654957'
}

# Função para consultar os dados do Pipefy
def fetch_pipefy_data(pipe_id):
    query = f"""
    query {{
      pipe(id: "{pipe_id}") {{
        id
        name
        phases {{
          id
          name
          cards {{
            edges {{
              node {{
                id
                title
                fields {{
                  name
                  value
                }}
              }}
            }}
          }}
        }}
      }}
    }}
    """

    headers = {'Authorization': f'Bearer {PIPEFY_API_TOKEN}'}
    print(f"Buscando dados para o pipe ID: {pipe_id}")
    response = requests.post(PIPEFY_GRAPHQL_ENDPOINT, json={'query': query}, headers=headers)
    data = response.json()
    print(f"Resposta da API para o pipe {pipe_id}: {data}")
    return data 

# Função para criar/atualizar o arquivo do Excel com os dados do Pipefy
def update_excel(wb, data, sheet_name):
    if sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        ws.delete_rows(2, ws.max_row)
    else:
        ws = wb.create_sheet(title=sheet_name)
    print(f"Aba '{sheet_name}' selecionada ou criada.")

    # Adicionando os cabeçalhos
    for col_num, header in enumerate(HEADERS, 1):
        ws.cell(row=1, column=col_num, value=header)

    # Adicionando os dados
    row_num = 2
    for phase in data['data']['pipe']['phases']:
        for card_edge in phase['cards']['edges']:
            card = card_edge['node']
            for field in card['fields']:
                ws.cell(row=row_num, column=1, value=data['data']['pipe']['id'])
                ws.cell(row=row_num, column=2, value=data['data']['pipe']['name'])
                ws.cell(row=row_num, column=3, value=phase['name'])
                ws.cell(row=row_num, column=4, value=card['id'])
                ws.cell(row=row_num, column=5, value=card['title'])
                ws.cell(row=row_num, column=6, value=field['name'])
                ws.cell(row=row_num, column=7, value=field['value'])
                row_num += 1

    # Ajustando a largura das colunas
    for col in ws.columns:
        max_length = 0
        column = col[0].column  # Get the column name
        for cell in col:
            try:  # Necessary to avoid error on empty cells
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2)
        ws.column_dimensions[get_column_letter(column)].width = adjusted_width

    wb.save(EXCEL_FILENAME)
    print(f'Arquivo "{EXCEL_FILENAME}" salvo com sucesso.')

# Função principal para executar o script
def main():
    wb = Workbook()
    wb.remove(wb.active)  # Remove a aba padrão vazia

    for pipe_name, pipe_id in PIPE_IDS.items():
        print(f"Iniciando a consulta dos dados do Pipefy para: {pipe_name}")
        data = fetch_pipefy_data(pipe_id)

        if 'errors' in data:
            print(f"Erro ao fazer a consulta GraphQL para {pipe_name}:")
            for error in data['errors']:
                print(error['message'])
        else:
            print(f"Dados do Pipefy obtidos com sucesso para {pipe_name}. Atualizando o arquivo do Excel...")
            update_excel(wb, data, pipe_name)

    print(f"Salvando o arquivo '{EXCEL_FILENAME}'.")            

    wb.save(EXCEL_FILENAME)
    print(f'Arquivo "{EXCEL_FILENAME}" salvo com sucesso com todas as abas.')

if __name__ == "__main__":
    main()
