from FetchAllCards import fetch_all_cards
from UpdateExcelEnps import update_excel_e_nps
from UpdateExcelMatrizCursos import update_excel_matriz_cursos
from UpdateExcelPainelControleMembros import update_excel_painel_controle_membros
from openpyxl import Workbook, load_workbook
from Definitions import PIPE_IDS, PIPE_TO_FILE

# Função principal para executar o script
def main():
    # Mapeamento de pipes para suas funções de atualização específicas
    update_functions = {
        'RH - E-NPS': update_excel_e_nps,
        'RH - Matriz de Cursos': update_excel_matriz_cursos,
        'RH - Painel Controle Membros': update_excel_painel_controle_membros,
    }

    for pipe_name, (filename, sheet_name) in PIPE_TO_FILE.items():
        print(f"Iniciando a consulta dos dados do Pipefy para: {pipe_name}")
        all_phases = fetch_all_cards(PIPE_IDS[pipe_name])
        
        try:
            wb = load_workbook(filename)
            print(f"Arquivo '{filename}' carregado com sucesso.")
        except FileNotFoundError:
            wb = Workbook()
            print(f"Arquivo '{filename}' não encontrado, criando novo arquivo.")
            wb.remove(wb.active)  # Remover a aba padrão vazia

        # Chamada da função de atualização específica, agora passando todas as fases paginadas
        update_function = update_functions[pipe_name]
        update_function(wb, all_phases, sheet_name)

        wb.save(filename)
        print(f'Arquivo "{filename}" salvo com sucesso com a aba atualizada.')

if __name__ == "__main__":
    main()
