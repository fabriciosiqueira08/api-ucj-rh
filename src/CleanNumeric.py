import re

def clean_numeric(value):
    # Remove caracteres não numéricos exceto o ponto decimal
    if isinstance(value, str):
        cleaned_value = re.sub(r'[^\d.]+', '', value)
        try:
            # Tentativa de converter para float
            return float(cleaned_value)
        except ValueError:
            # Se falhar, retorna None para indicar falha na conversão
            return None
    elif isinstance(value, (int, float)):
        return value
    else:
        return None