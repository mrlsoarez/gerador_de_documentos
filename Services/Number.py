from num2words import num2words

def converter_currency(valor):
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def pegar_numero_extenso(numero, cambio=False):

    if not cambio:
        return num2words(numero, lang='pt_BR')

    valor_formatado = (
        f"{numero:,.2f}"
        .replace(",", "X")
        .replace(".", ",")
        .replace("X", ".")
    )

    reais = int(numero)
    centavos = round((numero - reais) * 100)

    texto_reais = num2words(reais, lang='pt_BR')
    texto_centavos = num2words(centavos, lang='pt_BR')

    if reais == 1:
        texto_reais += " real"
    else:
        texto_reais += " reais"

    if centavos == 1:
        texto_centavos += " centavo"
    else:
        texto_centavos += " centavos"

    return f"{texto_reais} e {texto_centavos}"
