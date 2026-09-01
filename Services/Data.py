from datetime import date

def EncontrarData(parametro, data, *args):
    
    calendario = {
        "01": "JANEIRO",
        "02": "FEVEREIRO",
        "03": "MARÇO",
        "04": "ABRIL",
        "05": "MAIO",
        "06": "JUNHO",
        "07": "JULHO",
        "08": "AGOSTO",
        "09": "SETEMBRO",
        "10": "OUTUBRO",
        "11": "NOVEMBRO",
        "12": "DEZEMBRO"
    }
    
    """
    def PegarDataDeHoje():
        return str(date.today().strftime("%d/%m/20%y"))
    
    data = PegarDataDeHoje()
    """
    
    def PegarDiaDeHoje(data):
        return data[:2]
    
    def PegarMes(data, extenso = False):
        if (extenso): return calendario[data[3:5]]
        return data[3:5]
    
    def PegarAno(data):
        return data[6:]
    
    if (parametro == "dia"):
        return PegarDiaDeHoje(data)
    elif (parametro == "mes"):
        return PegarMes(data, args[0])
    elif (parametro == "ano"):
        return PegarAno(data)

