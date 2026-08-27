from Entities.Processo import Processo
from Entities.Item import Item 

from docx import Document

## parei no primeiro texto, defiir cargo secretária
class Contrato:
    def __init__(self, processo, secretaria, modelo, destino):
        self.processo = None 
        self.secretaria = secretaria
        self.modelo = modelo
        self.destino = destino
        
    def set_mapeamento(self, sheet):
        
        def definir_items():
            contador = 10
            array = []
            while True: 
                if (sheet["A" + str(contador)].value == None):
                    break 
                array.append(
                    Item(
                        sheet["A" + str(contador)].value,
                        sheet["B" + str(contador)].value,
                        sheet["C" + str(contador)].value,
                        sheet["D" + str(contador)].value,
                        sheet["E" + str(contador)].value,
                        sheet["F" + str(contador)].value,
                        sheet["G" + str(contador)].value
                    )
                )
                contador += 1
            return array
        
        self.mapeamento = {
            "n_contrato": sheet["B2"].value,
            "processo": sheet["B3"].value,
            "cnpj": sheet["B4"].value,
            "modalidade": sheet["B5"].value,
            "n_modalidade": sheet["B6"].value, 
            "objeto": sheet["B8"].value,
            "items": sheet["B10"].value
        }
        
        self.set_info_processo()      
        self.items = definir_items()
        
    def set_info_processo(self):
        self.processo = Processo(self.mapeamento["processo"], self.mapeamento["cnpj"], "", self.mapeamento["modalidade"], self.mapeamento["n_modalidade"], self.mapeamento["n_contrato"], self.mapeamento["objeto"])

    def buscar_info_empresa(self): 
        pass 
    
    def criar_documento(self):
        
        doc = Document(self.modelo)

        processo = self.processo 
        secretaria = self.secretaria
        
        def definir_cabecalho(): 
            print(f"CONTRATO ADMINISTRATIVO Nº {processo.n_contrato}\nPROCESSO ADMINISTRATIVO Nº {processo.n_processo}\n{processo.modalidade} Nº {processo.n_modalidade}")
            pass 
        
        def definir_texto_inicial():
            print(f"O {secretaria.fundo.upper()}/MS, com sede à Avenida Aquidauana, nº 1001, Centro, Bataguassu/MS inscrito(a) no CNPJ sob o nº {secretaria.cnpj}, neste ato representado, por {secretaria.nome.upper()}, Secretária Municipal de Agricultura e Meio Ambiente, nomeada pelo Decreto nº 04, de 02 de janeiro de 2025, publicado no Diário Oficial da Assomasul de 07 de janeiro de 2025, portadora da Matrícula Funcional nº 11.7676-1, doravante denominada CONTRATANTE")
        
        definir_cabecalho()
        definir_texto_inicial()