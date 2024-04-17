import requests
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
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
        phases {{
          name
          cards {{
            edges {{
              node {{
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

    # Lista de títulos para os cabeçalhos conforme especificado
    headers = [
        "Período:", "NPS", "Grupo minoritário?", "Cargo:",
        "Sinto-me envolvido com o trabalho que faço.",
        "Estou entusiasmado com meu trabalho.",
        "Em meu trabalho, sinto-me cheio de energia.",
        "Eu entendo como meu trabalho contribui para o alcance das metas e objetivos da empresa.",
        "Eu sinto que faço a diferença no meu time.",
        "Eu sinto que, se eu cometer um erro, isso não se voltará contra mim.",
        "Sinto que a cultura da UCJ está alinhada com as minhas crenças e valores.",
        "A UCJ possui lideranças com as quais me identifico.",
        "Sinto que a minha liderança direta se preocupa comigo como pessoa.",
        "Sinto que a minha liderança direta constrói um ambiente positivo, ou seja, temos uma comunicação aberta e transparente, falamos de dificuldades e temos uma cultura de feedbacks constantes.",
        "Eu estou satisfeito em relação ao tempo que dedico para o meu trabalho, meus estudos, minha família, meus amigos e minha saúde.",
        "Eu sinto que a minha liderança direta encoraja e apoia meu desenvolvimento.",
        "Sinto que tenho voz ativa opinar e fazer acontecer as transformações em que acredito.",
        "Sinto que sou comunicado (a) das informações relevantes para o meu trabalho e sobre assuntos gerais relevantes na empresa.",
        "Sinto que o ambiente em que trabalho colabora para a minha produtividade.",
        "O que motivou sua resposta.",
        "NPS produtos", "Comente sobre o que motivou essa resposta.",
        "O que podemos fazer para melhorar enquanto empresa?", "Qual(is)?"
    ]

    # Aplicando os cabeçalhos e seus estilos
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_num, value=header)
        cell.font = Font(name='Arial', size=11, bold=True)
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = Border(left=Side(style='thin'), right=Side(style='thin'),
                             top=Side(style='thin'), bottom=Side(style='thin'))

    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}1"

    row_num = 2

    # Estilo de fonte para o restante das células
    normal_font = Font(name='Arial', size=11, bold=False)
    alignment_bottom = Alignment(vertical='bottom')

    index_col_t = headers.index("O que motivou sua resposta.") + 1
    index_col_v = headers.index("Comente sobre o que motivou essa resposta.") + 1
    index_col_w = headers.index("O que podemos fazer para melhorar enquanto empresa?") + 1


    for phase in data['data']['pipe']['phases']:
        if phase['name'] not in ["Caixa de entrada", "Concluído"]:
            for card_edge in phase['cards']['edges']:
                card = card_edge['node']
                field_values = {header: "" for header in headers}  # Inicializa todos os campos com string vazia
                field_values["Período:"] = phase['name']
                field_values["NPS"] = card['title']

                for field in card['fields']:
                    header_name = next((h for h in headers if h.endswith(field['name'] + ":")), None)
                    if field['name'] in headers:
                        field_values[field['name']] = field['value']  # Preenche os valores onde o cabeçalho coincide com o nome do campo
                    elif header_name:
                        field_values[header_name] = field['value']
                    elif field['name'] == "Você pertence a algum grupo minoritário?":
                        field_values["Grupo minoritário?"] = field['value']
                    elif field['name'] == "Você é:":
                        field_values["Cargo:"] = field['value']
                    elif field['name'] == "Em uma escala de 0 a 10, o quanto você recomendaria os produtos da UCJ para amigos ou familiares?":
                        field_values["NPS produtos"] = field['value']
                    elif field['name'] == "Comente sobre o que motivou sua resposta.":
                        field_values["O que motivou sua resposta."] = field['value']
                # Preenche a linha com os valores coletados
                for col_num, header in enumerate(headers, 1):
                    cell = ws.cell(row=row_num, column=col_num, value=field_values[header])
                    cell.font = normal_font
                    cell.alignment = alignment_bottom
                    if col_num in [index_col_t, index_col_v, index_col_w]:
                        cell.alignment = Alignment(wrap_text=True)  # Aplica quebra de texto nas colunas específicas


                row_num += 1  # Incrementa a linha após preencher todos os campos

    # Ajuste das colunas conforme anteriormente
        for col in ws.columns:
            max_length = max((len(str(cell.value)) if cell.value is not None else 0 for cell in col), default=0)
            adjusted_width = max_length + 2
            ws.column_dimensions[get_column_letter(col[0].column)].width = adjusted_width
    
    standard_width = 8.43
    for col in range(5, 20):  # Colunas E(5) até S(19)
        ws.column_dimensions[get_column_letter(col)].width = standard_width

    column_t_index = headers.index("O que motivou sua resposta.") + 1  # Adjust if the header index changes
    t_header_value = ws.cell(row=1, column=column_t_index).value
    ws.column_dimensions[get_column_letter(column_t_index)].width = len(t_header_value) + 2

    column_t_index = headers.index("Comente sobre o que motivou essa resposta.") + 1  # Adjust if the header index changes
    t_header_value = ws.cell(row=1, column=column_t_index).value
    ws.column_dimensions[get_column_letter(column_t_index)].width = len(t_header_value) + 2
    
    column_t_index = headers.index("O que podemos fazer para melhorar enquanto empresa?") + 1  # Adjust if the header index changes
    t_header_value = ws.cell(row=1, column=column_t_index).value
    ws.column_dimensions[get_column_letter(column_t_index)].width = len(t_header_value) + 2


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
