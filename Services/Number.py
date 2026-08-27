from num2words import num2words



def pegar_numero_extenso(numero):
    texto = num2words(numero, lang='pt_BR')
    return texto
