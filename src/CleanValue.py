# Função para limpar os valores de campo de lista
def clean_value(value):
    if isinstance(value, str) and value.startswith("[\"") and value.endswith("\"]"):
        # Assume que há apenas um item na lista e remove os caracteres indesejados
        return value[2:-2]  # Remove os dois primeiros e os dois últimos caracteres
    return value