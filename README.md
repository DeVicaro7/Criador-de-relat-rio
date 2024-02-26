# Criador-de-relat-rio
Leitor de xml (notas fiscais) para extração de dados e preenchimento de planilha de estoque em excel.




from xml.dom import minidom
import pandas as pd
#1- FAZER LEITURA DE XML PARA TRATAMENTO DE DADOS E EXTRAIR:
# #NOTA DE INDUSTRIALIZAÇÃO #NOTA DE REMESSA #CODIGO DO PRODUTO
# #REFERENCIA CLIENTE #QUANTIDADE EM KG #QUANTIDADE DE PALETES
# #ESPÉCIE #VALOR UNITÁRIO #LEMBRETE:
# CRIAR OBJETOS PARA REALIZAR OS PROCESSOS ABAIXO
with open("29240202036483000614550010003667791276107693-procNFe.xml", "r", encoding="utf-8") as fileopening:
    xml = minidom.parse(fileopening) #FRAGMENTANDO O XML #BUSCANDO OS ELEMENTOS DAS TAGS
    nf_industrializacao = xml.getElementsByTagName("infCpl")
    nf_remessa = xml.getElementsByTagName("nNF")
    cod_prod = xml.getElementsByTagName("cProd")
    ref_cliente = xml.getElementsByTagName("infCpl")
    qtd_kg = xml.getElementsByTagName("qCom")
    qtd_plts = xml.getElementsByTagName("qVol")
    especie = xml.getElementsByTagName("esp")
    valor_unitario = xml.getElementsByTagName("vUnCom")
    for n in nf_industrializacao:
        tratamento_nf = str(n.firstChild.data).split("NF-e") #TRATAMENTO PARA RETIRADA DOS DADOS
        tratamento2 = str(tratamento_nf[1]).split(" ") #ENDEREÇANDO CADA PALAVRA DENTRO DO TEXTO NO FORMATO DE LISTA
        dt_nf_ind = tratamento2[1] #EXTRAINDO (NOTA DE INDUSTRIALIZAÇÃO)
        for n in nf_remessa:
            dt_nf_reme = str(n.firstChild.data) #EXTRAINDO (NOTA DE REMESSA)
        for n in cod_prod:
            dt_cod_prod = str(n.firstChild.data) #EXTRAINDO (CÓDIGO DO PRODUTO)
        for n in ref_cliente:
            tratamento_rf = str(n.firstChild.data).split("MP") #TRATAMENTO PARA RETIRADA DOS DADOS
            tratamento2 = tratamento_rf[2]
            tratamento3 = str(tratamento2).split(" ") #ENDEREÇANDO CADA PALAVRA DENTRO DO TEXTO NO FORMATO DE LISTA
            dt_ref_clie = tratamento3[1] #EXTRAINDO (REFERÊNCIA DO CLIENTE)
        for n in qtd_kg:
            dt_qtd_kg = str(n.firstChild.data) #EXTRAINDO (QUANTIDADE EM KG)
        for n in qtd_plts:
            dt_qtd_plts = str(n.firstChild.data) #EXTRAINDO (QUANTIDADE DE PALETES)
        for n in especie:
            dt_especie = str(n.firstChild.data) #EXTRAINDO (ESPECIE)
        for n in valor_unitario:
            dt_valor_uni = str(n.firstChild.data) #EXTRAINDO (VALOR UNITÁRIO) #CRIAR UMA COMUNICAÇÃO COM A PLANILHA
#CRIANDO VINCULO COM A PLANILHA EM EXCEL
lista = { "NOTA INDUSTRIALICAÇÃO": [f"{dt_nf_ind}"],
          "NOTA DE REMESSA": [f"{dt_nf_reme}"],
          "CÓD PRODUTO": [f"{dt_cod_prod}"],
          "MP": [f"{dt_ref_clie}"],
          "QTD KG": [f"{dt_qtd_kg}"],
          "QTD PALETES": [f"{dt_qtd_plts}"],
          "ESPECIE": [f"{dt_especie}"],
          "VALOR UNIT": [f"{dt_valor_uni}"]
          }
pd.DataFrame(lista).to_excel("./testando_arquivo.xlsx")
