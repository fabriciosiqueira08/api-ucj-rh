from CleanValue import clean_value


def process_card_clima (card, headers):
    field_values = {header: "" for header in headers}
    field_values["Membro"] = card["title"]

    for field in card['fields']:
        if field['name'] in headers:
            value = field['value']
            try:
                field_values[field['name']] = int(value)
            except ValueError:
                try:
                    field_values[field['name']] = float(value.replace(',', '.'))
                except ValueError:
                    field_values[field['name']] = clean_value(field['value'])

    return field_values