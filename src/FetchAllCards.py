from FetchPipefyData import fetch_pipefy_data

def fetch_all_cards(pipe_id):
    all_phases = []
    initial_data = fetch_pipefy_data(pipe_id)  # Esta função precisa tratar corretamente o cursor inicial como None
    
    for phase_data in initial_data['data']['pipe']['phases']:

        if phase_data['name'] == "Histórico":
            continue  # Pula a fase "Histórico"

        phase = phase_data
        cards = phase_data['cards']['edges']
        pageInfo = phase_data['cards']['pageInfo']
        cursor = pageInfo['endCursor']

        # Assegurar que a fase atual seja corretamente paginada
        while pageInfo['hasNextPage']:
            more_data = fetch_pipefy_data(pipe_id, cursor)  # Adicionar parâmetro para fase
            more_phase_data = next((p for p in more_data['data']['pipe']['phases'] if p['name'] == phase['name']), None)
            if more_phase_data:
                more_cards = more_phase_data['cards']['edges']
                cards.extend(more_cards)
                pageInfo = more_phase_data['cards']['pageInfo']
                cursor = pageInfo['endCursor']
            else:
                break  # Interrompe se a fase específica não for encontrada

        phase['cards']['edges'] = cards
        all_phases.append(phase)
        print(f"Fase '{phase['name']}' processada com {len(cards)} cartões.")

    return all_phases
