# documento tecnico de reequilibrio economico
# apostilamento
# termo aditivo
# contrato

# main thing: buscar informaçoes a partir de planilha

from Services.Planilha import abrir_planilha
from Entities.Processo import Processo
from Entities.Secretaria import Secretaria 

from Classes.Documento import Documento
from Classes.Apostilamento import Apostilamento
from Classes.Portaria import Portaria
from Classes.Contrato import Contrato

def env():
    return {
        "planilha_principal": r"C:\Users\Usuario\Documents\MRL\MODELOS BASE\GERADOR DE DOCUMENTO.xlsx",
        "modelo_timbre": r"C:\Users\Usuario\Documents\MRL\MODELOS BASE\3. MODELOS DE DOCUMENTO\MODELO TIMBRE.docx",
        "modelo_portaria": r"C:\Users\Usuario\Documents\MRL\MODELOS BASE\3. MODELOS DE DOCUMENTO\MODELO DE PORTARIA.docx",
        "planilhas": r"C:\Users\Usuario\Documents\MRL\MODELOS BASE\4. PLANS",
        "destino": r"C:\Users\Usuario\Documents\docs"
    }

def iniciar(sheet):
    
    def mapa(sheet):
        return {
            "info_processo": {
                "processo": sheet["B5"].value,
                "cnpj": sheet["B6"].value,
                "fornecedor": sheet["B7"].value, 
                "modalidade": sheet["B8"].value,
                "n_modalidade": sheet["B9"].value,
                "n_contrato": sheet["B10"].value,
                "objeto": sheet["B11"].value
            },
            "info_secretaria": {
                "nome_secretaria": sheet["B13"].value,
                "fundo": sheet["B14"].value,
                "nome": sheet["B15"].value,
                "cnpj": sheet["B16"].value,
                "matricula": sheet["B17"].value,
                "decreto": sheet["B18"].value,
                "data": sheet["B19"].value,
            },
            "modulos": {
                "portaria": {"inicializar_modulo": sheet["B22"].value, "classe": Portaria, "planilha": f"{enderecos['planilhas']}\PORTARIA.xlsx"},
                "termo_aditivo": {"inicializar_modulo": sheet["B23"].value, "classe": ""},
                "contrato": {"inicializar_modulo": sheet["B24"].value, "classe": Contrato, "planilha": f"{enderecos['planilhas']}\CONTRATO.xlsx"},
                "apostilamento": {"inicializar_modulo": sheet["B25"].value, "classe": Apostilamento, "planilha": f"{enderecos['planilhas']}\APOSTILAMENTO.xlsx"},
                "reequilibrio": {"inicializar_modulo": sheet["B26"].value, "classe": ""}
            }
        }
    
    def definir_info_multipla():
        print()
    
    info_inicial = mapa(sheet)  
    
    info_processo = Processo(info_inicial["info_processo"]["processo"],
                            info_inicial["info_processo"]["fornecedor"],
                            info_inicial["info_processo"]["cnpj"],
                            info_inicial["info_processo"]["modalidade"],
                            info_inicial["info_processo"]["n_modalidade"],
                            info_inicial["info_processo"]["n_contrato"],
                            info_inicial["info_processo"]["objeto"])  
    info_secretaria = Secretaria(info_inicial["info_secretaria"]["nome_secretaria"],
                                 info_inicial["info_secretaria"]["fundo"], 
                                 info_inicial["info_secretaria"]["nome"],
                                 info_inicial["info_secretaria"]["cnpj"],
                                 info_inicial["info_secretaria"]["matricula"],
                                 info_inicial["info_secretaria"]["decreto"],
                                 info_inicial["info_secretaria"]["data"]) 
    
    if (info_inicial["modulos"]["portaria"]["inicializar_modulo"]):
        array = []
        contador = 28
        while True: 
            if (sheet["A" + str(contador)].value == None):
                break
            array.append(
                {
                    "fornecedor": sheet["A" + str(contador)].value,
                    "cnpj": sheet["B" + str(contador)].value,
                    "n_contrato": sheet["C" + str(contador)].value
                }
            )
            contador += 1
        info_processo.set_fornecedor(array)
    doc = Documento(info_processo, info_secretaria, info_inicial["modulos"])
    return doc 

enderecos = env()
sheet = abrir_planilha(enderecos["planilha_principal"], "Gerador")
doc_principal = iniciar(sheet)

for keys in doc_principal.info:
    if (doc_principal.info[keys]["inicializar_modulo"]):
        doc_segundario = doc_principal.info[keys]["classe"](doc_principal.info_processo, doc_principal.info_secretarias, enderecos["modelo_timbre"], enderecos["destino"])
        sheet = abrir_planilha(doc_principal.info[keys]["planilha"], "Base")
        doc_segundario.set_mapeamento(sheet)
        doc_segundario.criar_documento()
    
"""
import requests, json

cnpj_exemplo = "48071593000103"
url = f"https://publica.cnpj.ws/cnpj/{cnpj_exemplo}"
session = requests.Session()
response = session.get(url)
dados = response.json()
print(dados)


{'cnpj_raiz': '03576220', 'razao_social': 'MUNICIPIO DE BATAGUASSU', 'capital_social': '0.00', 'responsavel_federativo': 'BATAGUASSU - MS', 'atualizado_em': '2026-08-08T03:00:00.000Z', 'porte': {'id': '05', 'descricao': 'Demais'}, 'natureza_juridica': {'id': '1244', 'descricao': 'Município'}, 'qualificacao_do_responsavel': {'id': 5, 'descricao': 'Administrador '}, 'socios': [], 'simples': None, 'estabelecimento': {'cnpj': '03576220000156', 'atividades_secundarias': [{'id': '3811400', 'secao': 'E', 'divisao': '38', 'grupo': '38.1', 'classe': '38.11-4', 'subclasse': '3811-4/00', 'descricao': 'Coleta de resíduos não perigosos'}, {'id': '4213800', 'secao': 'F', 'divisao': '42', 'grupo': '42.1', 'classe': '42.13-8', 'subclasse': '4213-8/00', 'descricao': 'Obras de urbanização - ruas, praças e calçadas'}, {'id': '7500100', 'secao': 'M', 'divisao': '75', 'grupo': '75.0', 'classe': '75.00-1', 'subclasse': '7500-1/00', 'descricao': 'Atividades veterinárias'}, {'id': '8512100', 'secao': 'P', 'divisao': '85', 'grupo': '85.1', 'classe': '85.12-1', 'subclasse': '8512-1/00', 'descricao': 'Educação infantil - pré-escola'}, {'id': '8513900', 'secao': 'P', 'divisao': '85', 'grupo': '85.1', 'classe': '85.13-9', 'subclasse': '8513-9/00', 'descricao': 'Ensino fundamental'}, {'id': '8591100', 'secao': 'P', 'divisao': '85', 'grupo': '85.9', 'classe': '85.91-1', 'subclasse': '8591-1/00', 'descricao': 'Ensino de esportes'}, {'id': '8630503', 'secao': 'Q', 'divisao': '86', 'grupo': '86.3', 'classe': '86.30-5', 'subclasse': '8630-5/03', 'descricao': 'Atividade médica ambulatorial restrita a consultas'}, {'id': '8800600', 'secao': 'Q', 'divisao': '88', 'grupo': '88.0', 'classe': '88.00-6', 'subclasse': '8800-6/00', 'descricao': 'Serviços de assistência social sem alojamento'}], 'cnpj_raiz': '03576220', 'cnpj_ordem': '0001', 'cnpj_digito_verificador': '56', 'tipo': 'Matriz', 'nome_fantasia': 'BATAGUASSU PREFEITURA GABINETE DO PREFEITO', 'situacao_cadastral': 'Ativa', 'data_situacao_cadastral': '2004-11-20', 'data_inicio_atividade': '1976-02-24', 'nome_cidade_exterior': None, 'tipo_logradouro': 'AVENIDA', 'logradouro': 'AQUIDAUANA', 'numero': '1001', 'complemento': None, 'bairro': 'CENTRO', 'cep': '79780000', 'ddd1': '67', 'telefone1': '40429000', 'ddd2': None, 'telefone2': None, 'ddd_fax': '67', 'fax': '35411277', 'email': 'gabinete@bataguassu.ms.gov.br', 'situacao_especial': None, 'data_situacao_especial': None, 'atualizado_em': '2026-08-08T03:00:00.000Z', 'atividade_principal': {'id': '8411600', 'secao': 'O', 'divisao': '84', 'grupo': '84.1', 'classe': '84.11-6', 'subclasse': '8411-6/00', 'descricao': 'Administração pública em geral'}, 'pais': {'id': '1058', 'iso2': 'BR', 'iso3': 'BRA', 'nome': 'Brasil', 'comex_id': '105'}, 'estado': {'id': 12, 'nome': 'Mato Grosso do Sul', 'sigla': 'MS', 'ibge_id': 50}, 'cidade': {'id': 5115, 'nome': 'Bataguassu', 'ibge_id': 5001904, 'siafi_id': '9037'}, 'motivo_situacao_cadastral': None, 'inscricoes_estaduais': []}}
    
"""