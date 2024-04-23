from CleanValue import clean_value


def process_card_membros(card, headers):
    # Inicializa um dicionário com todos os headers definidos para strings vazias.
    field_values = {header: "" for header in headers}

    # Atribui o título do node ao campo "Membro" especificamente.
    field_values["Membro"] = card["title"]

    # Processa os outros campos que podem estar presentes em 'fields'.
    for field in card['fields']:
        field_name = field['name']
        if field_name in headers and field_name != "Membro":
            value = field['value']
            if value is None:
                field_values[field_name] = ""
            else:
                try:
                    field_values[field_name] = int(value)
                except ValueError:
                    try:
                        field_values[field_name] = float(value.replace(',', '.'))
                    except ValueError:
                        field_values[field_name] = clean_value(value)

    return field_values
