class Processo:
    def __init__(self, processo, cnpj, fornecedor, modalidade, n_modalidade, n_contrato, objeto):
        self.n_processo = processo 
        self.cnpj = cnpj
        self.fornecedor = fornecedor 
        self.modalidade = modalidade 
        self.n_modalidade = n_modalidade
        self.n_contrato = n_contrato
        self.objeto = objeto    
        
    def set_fornecedor(self, dado):
        self.fornecedor = dado 
