from num2words import num2words

def converter_currency(valor):
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def pegar_numero_extenso(numero):
    texto = num2words(numero, lang='pt_BR')
    return texto
