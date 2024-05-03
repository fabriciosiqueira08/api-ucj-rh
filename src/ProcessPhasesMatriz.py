from openpyxl.styles import Font, Alignment
from ProcessCard import process_card
from datetime import datetime

def process_phases_matriz(ws, headers, all_phases, row_num):
    normal_font = Font(name='Arial', size=10, bold=False)
    alignment_bottom = Alignment(vertical='bottom')

    for phase in all_phases:
        if isinstance(phase, dict) and phase.get('name') in ["In-Company", "Individual"]:
            for card_edge in phase['cards']['edges']:
                card = card_edge['node']
                field_values = process_card(card, headers)

                # Obtém e formata a data de criação
                created_at_str = card.get('createdAt', '')
                if created_at_str:
                    try:
                        created_at = datetime.fromisoformat(created_at_str.replace("Z", "+00:00"))
                        created_at_formatted = created_at.strftime('%d/%m/%Y')
                    except ValueError:
                        created_at_formatted = created_at_str
                else:
                    created_at_formatted = ''

                # Preenche a coluna "Data:"
                ws.cell(row=row_num, column=1, value=created_at_formatted)

                # Preenche as outras colunas
                for col_num, header in enumerate(headers[1:], 2):
                    cell_value = field_values.get(header, "")
                    cell = ws.cell(row=row_num, column=col_num, value=cell_value)
                    cell.font = normal_font
                    cell.alignment = alignment_bottom

                row_num += 1  # Mover para a próxima linha

    return ws, row_num
