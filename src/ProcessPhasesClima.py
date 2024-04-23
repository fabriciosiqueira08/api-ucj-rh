from openpyxl.styles import Font, Alignment
from ProcessCardClima import process_card_clima


def process_phases_clima(ws, headers, all_phases, row_num):

    # Estilo de fonte para o restante das células
    normal_font = Font(name='Arial', size=10, bold=False)
    alignment_bottom = Alignment(vertical='bottom')

    for phase in all_phases:
        if isinstance(phase, dict):  # Remove a verificação do nome da fase
            phase_name = phase.get('name', '')  # Obter o nome da fase, com fallback para string vazia se não encontrado
            for card_edge in phase['cards']['edges']:
                card = card_edge['node']
                field_values = process_card_clima(card, headers)
                field_values['Mês'] = phase_name

                # Mova o preenchimento da planilha para dentro deste loop
                for col_num, header in enumerate(headers, 1):
                    cell = ws.cell(row=row_num, column=col_num, value=field_values.get(header, ""))  # Uso seguro de .get com valor padrão
                    cell.font = normal_font
                    cell.alignment = alignment_bottom

                row_num += 1  # Incremento de linha dentro do loop de cards

    return ws, row_num