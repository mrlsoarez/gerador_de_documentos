import requests 

def limpar_cnpj(cnpj):
    return "".join(filter(str.isdigit, cnpj))

def buscar_cnpj(cnpj):
    cnpj = limpar_cnpj(cnpj)
    url = f"https://publica.cnpj.ws/cnpj/{cnpj}"
    session = requests.Session()
    response = session.get(url)
    return  response.json()