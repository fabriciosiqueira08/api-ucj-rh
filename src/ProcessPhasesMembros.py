from ProcessCardMembros import process_card_membros
from openpyxl.styles import Font, Alignment

def process_phases_membros(ws, headers, all_phases, row_num):
    # Estilo de fonte para o restante das células
    normal_font = Font(name='Arial', size=10, bold=False)
    alignment_bottom = Alignment(vertical='bottom')

    for phase in all_phases:
        if isinstance(phase, dict) and phase.get('name') == "Ativo":  # Ajuste para igualdade se apenas uma fase é relevante
            for card_edge in phase['cards']['edges']:
                card = card_edge['node']
                field_values = process_card_membros(card, headers)

                for col_num, header in enumerate(headers, 1):
                    cell = ws.cell(row=row_num, column=col_num, value=field_values[header])
                    cell.font = normal_font
                    cell.alignment = alignment_bottom

                row_num += 1

    return ws, row_num