from openpyxl.styles import Font, Alignment
from ProcessCard import process_card

# Função para processar as fases da Matriz de Cursos
def process_phases_matriz(ws, headers, all_phases, row_num):
        # Estilo de fonte para o restante das células
        normal_font = Font(name='Arial', size=10, bold=False)
        alignment_bottom = Alignment(vertical='bottom')
        
        
        for phase in all_phases:    
            if isinstance(phase, dict) and phase.get('name') in ["In-Company", "Individual"]:  # Assegura que 'phase' é um dicionário
                for card_edge in phase['cards']['edges']:
                    card = card_edge['node']
                    field_values = process_card(card, headers)

                    for col_num, header in enumerate(headers, 1):
                        cell = ws.cell(row=row_num, column=col_num, value=field_values[header])
                        cell.font = normal_font  # Certifique-se de que 'normal_font' está definido
                        cell.alignment = alignment_bottom  # Certifique-se de que 'alignment_bottom' está definido

                    row_num += 1


        return ws, row_num