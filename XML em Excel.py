import os
import tkinter.filedialog
from tkinter import *

import pandas as pd
import xmltodict
from tqdm import tqdm


def ler_xml_nfe_receita(nota):
    with open(nota, 'rb') as arquivo:
        documento = xmltodict.parse(arquivo, encoding='ANSI')
    dic_nota = documento['NFeLog']['procNFe']['NFe']['infNFe']
    ide = dic_nota['ide']
    emit = dic_nota['emit']
    dest = dic_nota['dest']
    tot = dic_nota['total']['ICMSTot']
    if 'infAdic' in dic_nota:
        try:
            if 'infCpl' in dic_nota['infAdic']:
                inf = dic_nota['infAdic']
        except:
            inf = 'Não possui Informações Complementares'
    else:
        inf = 'Não possui Informações Complementares'

    chave = documento['NFeLog']['procNFe']['NFe']['infNFe']['@Id'][3:]
    nf = ide['nNF']
    serie = ide['serie']
    data_emi = ide['dhEmi'][:10]
    dia_emi = data_emi[8:]
    mes_emi = data_emi[5:7]
    ano_emi = data_emi[0:4]
    data_emissao = f'{dia_emi}/{mes_emi}/{ano_emi}'
    tipo_op = ide['tpNF']
    local_op = ide['idDest']
    try:
        cnpj_emitente = emit['CNPJ']
    except:
        cnpj_emitente = emit['CPF']
    uf_emitente = emit['enderEmit']['UF']
    if 'IE' in emit:
        ie_emitente = emit['IE']
    else:
        ie_emitente = 'ISENTO'
    if 'idEstrangeiro' in dest:
        cnpj_destinatario = 'Estrangeiro - Não Possui CNPJ'
    else:
        try:
            cnpj_destinatario = dest['CNPJ']
        except:
            cnpj_destinatario = dest['CPF']
    uf_destinatario = dest['enderDest']['UF']
    if 'IE' in dest:
        ie_destinatario = dest['IE']
    else:
        ie_destinatario = 'ISENTO'
    bc_icms = tot['vBC']
    valor_icms = tot['vICMS']
    icms_desonerado = tot['vICMSDeson']
    fcp = tot['vFCP']
    fcp_st = tot['vFCPST']
    bc_icms_st = tot['vBCST']
    valor_icms_st = tot['vST']
    valorProd = tot['vProd']
    valorFrete = tot['vFrete']
    valorSeg = tot['vSeg']
    valorDesc = tot['vDesc']
    valorIPI = tot['vIPI']
    valorPis = tot['vPIS']
    valorCofins = tot['vCOFINS']
    valorNF = tot['vNF']
    try:
        if 'infCpl' in inf:
            infAdc = inf['infCpl']
        else:
            infAdc = 'Não possui Informações Complementares'
    except:
        infAdc = 'Não possui Informações Complementares'

    if tipo_op == '0':
        tipo_operacao = 'Entrada'
    else:
        tipo_operacao = 'Saída'

    if local_op == '1':
        operacao = 'Operação Interna'
    elif local_op == '2':
        operacao = 'Operação Interestadual'
    else:
        operacao = 'Operação com Exterior'
    if 'eveNFe' in documento['NFeLog']:
        try:
            dic_evento = documento['NFeLog']['eveNFe']
            carta_correcao_r = 'Não houve'
            cancelamento_r = 'Não houve'
            for evento in dic_evento:
                eventos = evento['evento']['infEvento']['tpEvento']
                if '110110' in eventos:
                    carta_correcao_r = evento['evento']['infEvento']['detEvento']['xCorrecao']
                else:
                    carta_correcao = 'Não houve'
                if carta_correcao_r != 'Não houve':
                    carta_correcao = carta_correcao_r
                if '110111' in eventos:
                    cancelamento_r = 'Nota Cancelada'
                else:
                    cancelamento = 'Não houve'
                if cancelamento_r != 'Não houve':
                    cancelamento = cancelamento_r
        except:
            dic_evento = documento['NFeLog']['eveNFe']
            evento_unico = dic_evento['evento']['infEvento']['tpEvento']
            if '110110' in str(evento_unico):
                carta_correcao = documento['NFeLog']['eveNFe']['evento']['infEvento']['detEvento']['xCorrecao']
                cancelamento = 'Não houve'
            elif '110111' in str(evento_unico):
                cancelamento = 'Nota Cancelada'
                carta_correcao = 'Não houve'
            else:
                carta_correcao = 'Não houve'
                cancelamento = 'Não houve'
    else:
        carta_correcao = 'Não houve'
        cancelamento = 'Não houve'

    nfe_receita = {
        'Número da NF': nf,
        'Série': serie,
        'Data de Emissao': data_emissao,
        'Chave de Acesso': chave,
        'Tipo de Operacao': tipo_operacao,
        'Operação': operacao,
        'CNPJ/CPF do Emitente': cnpj_emitente,
        'IE Emitente': ie_emitente,
        'UF Emitente': uf_emitente,
        'CNPJ/CPF do destinatario': cnpj_destinatario,
        'IE Destinatario': ie_destinatario,
        'UF Destinatario': uf_destinatario,
        'Valor dos Produtos': float(valorProd),
        'Valor do Desconto': float(valorDesc),
        'Valor do Frete': float(valorFrete),
        'Valor do Seguro': float(valorSeg),
        'Valor da NF': float(valorNF),
        'Base do ICMS': float(bc_icms),
        'Valor do ICMS': float(valor_icms),
        'Base ICMS ST': float(bc_icms_st),
        'Valor ICMS ST': float(valor_icms_st),
        'Valor do IPI': float(valorIPI),
        'Valor ICMS desonerado': float(icms_desonerado),
        'Valor fcp': float(fcp),
        'Valor fcp ST': float(fcp_st),
        'Valor Pis': float(valorPis),
        'Valor Cofins': float(valorCofins),
        'Informação Adicional': infAdc,
        'Carta de Correção': carta_correcao,
        'Cancelamento': cancelamento,
    }
    return nfe_receita


def ler_xml_nfe(nota):
    with open(nota, 'rb') as arquivo:
        documento = xmltodict.parse(arquivo)
    if 'nfeProc' in documento:
        dic_nota = documento['nfeProc']['NFe']['infNFe']
        ide = dic_nota['ide']
        emit = dic_nota['emit']
        dest = dic_nota['dest']
        tot = dic_nota['total']['ICMSTot']
        if 'infAdic' in dic_nota:
            try:
                if 'infCpl' in dic_nota['infAdic']:
                    inf = dic_nota['infAdic']
            except:
                inf = 'Não possui Informações Complementares'
        else:
            inf = 'Não possui Informações Complementares'

        chave = documento['nfeProc']['NFe']['infNFe']['@Id'][3:]
        nf = ide['nNF']
        serie = ide['serie']
        data_emi = ide['dhEmi'][:10]
        dia_emi = data_emi[8:]
        mes_emi = data_emi[5:7]
        ano_emi = data_emi[0:4]
        data_emissao = f'{dia_emi}/{mes_emi}/{ano_emi}'
        tipo_op = ide['tpNF']
        local_op = ide['idDest']
        try:
            cnpj_emitente = emit['CNPJ']
        except:
            cnpj_emitente = emit['CPF']
        uf_emitente = emit['enderEmit']['UF']
        if 'IE' in emit:
            ie_emitente = emit['IE']
        else:
            ie_emitente = 'ISENTO'
        if 'idEstrangeiro' in dest:
            cnpj_destinatario = 'Estrangeiro - Não Possui CNPJ'
        else:
            try:
                cnpj_destinatario = dest['CNPJ']
            except:
                cnpj_destinatario = dest['CPF']
        uf_destinatario = dest['enderDest']['UF']
        if 'IE' in dest:
            ie_destinatario = dest['IE']
        else:
            ie_destinatario = 'ISENTO'
        bc_icms = tot['vBC']
        valor_icms = tot['vICMS']
        icms_desonerado = tot['vICMSDeson']
        fcp = tot['vFCP']
        fcp_st = tot['vFCPST']
        bc_icms_st = tot['vBCST']
        valor_icms_st = tot['vST']
        valorProd = tot['vProd']
        valorFrete = tot['vFrete']
        valorSeg = tot['vSeg']
        valorDesc = tot['vDesc']
        valorIPI = tot['vIPI']
        valorPis = tot['vPIS']
        valorCofins = tot['vCOFINS']
        valorNF = tot['vNF']
        try:
            if 'infCpl' in inf:
                infAdc = inf['infCpl']
            else:
                infAdc = 'Não possui Informações Complementares'
        except:
            infAdc = 'Não possui Informações Complementares'

        if tipo_op == '0':
            tipo_operacao = 'Entrada'
        else:
            tipo_operacao = 'Saída'

        if local_op == '1':
            operacao = 'Operação Interna'
        elif local_op == '2':
            operacao = 'Operação Interestadual'
        else:
            operacao = 'Operação com Exterior'
        carta_correcao = 'Não possui info no XML'
        cancelamento = 'Não possui info no XML'

        nfe = {
            'Número da NF': nf,
            'Série': serie,
            'Data de Emissao': data_emissao,
            'Chave de Acesso': chave,
            'Tipo de Operacao': tipo_operacao,
            'Operação': operacao,
            'CNPJ/CPF do Emitente': cnpj_emitente,
            'IE Emitente': ie_emitente,
            'UF Emitente': uf_emitente,
            'CNPJ/CPF do destinatario': cnpj_destinatario,
            'IE Destinatario': ie_destinatario,
            'UF Destinatario': uf_destinatario,
            'Valor dos Produtos': float(valorProd),
            'Valor do Desconto': float(valorDesc),
            'Valor do Frete': float(valorFrete),
            'Valor do Seguro': float(valorSeg),
            'Valor da NF': float(valorNF),
            'Base do ICMS': float(bc_icms),
            'Valor do ICMS': float(valor_icms),
            'Base ICMS ST': float(bc_icms_st),
            'Valor ICMS ST': float(valor_icms_st),
            'Valor do IPI': float(valorIPI),
            'Valor ICMS desonerado': float(icms_desonerado),
            'Valor fcp': float(fcp),
            'Valor fcp ST': float(fcp_st),
            'Valor Pis': float(valorPis),
            'Valor Cofins': float(valorCofins),
            'Informação Adicional': infAdc,
            'Carta de Correção': carta_correcao,
            'Cancelamento': cancelamento,
        }
        return nfe


    else:
        chave = documento['NFe']['infNFe']['@Id'][3:]
        nf = documento['NFe']['infNFe']['ide']['nNF']
        nfe = {
            'Número da NF': nf,
            'Série': 'Verificar se não está denegada',
            'Data de Emissao': 'Verificar se não está denegada',
            'Chave de Acesso': chave,
            'Tipo de Operacao': 'Verificar se não está denegada',
            'Operação': 'Verificar se não está denegada',
            'CNPJ/CPF do Emitente': 'Verificar se não está denegada',
            'IE Emitente': 'Verificar se não está denegada',
            'UF Emitente': 'Verificar se não está denegada',
            'CNPJ/CPF do destinatario': 'Verificar se não está denegada',
            'IE Destinatario': 'Verificar se não está denegada',
            'UF Destinatario': 'Verificar se não está denegada',
            'Valor dos Produtos': 'Verificar se não está denegada',
            'Valor do Desconto': 'Verificar se não está denegada',
            'Valor do Frete': 'Verificar se não está denegada',
            'Valor do Seguro': 'Verificar se não está denegada',
            'Valor da NF': 'Verificar se não está denegada',
            'Base do ICMS': 'Verificar se não está denegada',
            'Valor do ICMS': 'Verificar se não está denegada',
            'Base ICMS ST': 'Verificar se não está denegada',
            'Valor ICMS ST': 'Verificar se não está denegada',
            'Valor do IPI': 'Verificar se não está denegada',
            'Valor ICMS desonerado': 'Verificar se não está denegada',
            'Valor fcp': 'Verificar se não está denegada',
            'Valor fcp ST': 'Verificar se não está denegada',
            'Valor Pis': 'Verificar se não está denegada',
            'Valor Cofins': 'Verificar se não está denegada',
            'Informação Adicional': 'Verificar se não está denegada',
            'Carta de Correção': 'Verificar se não está denegada',
            'Cancelamento': 'Verificar se não está denegada',
        }
        return nfe


def ler_xml_nfce(nota):
    with open(nota, 'rb') as arquivo:
        documento = xmltodict.parse(arquivo)
    dic_nota = documento['nfeProc']['NFe']['infNFe']
    ide = dic_nota['ide']
    emit = dic_nota['emit']
    tot = dic_nota['total']['ICMSTot']

    chave = documento['nfeProc']['NFe']['infNFe']['@Id'][3:]
    nf = ide['nNF']
    serie = ide['serie']
    data_emi = ide['dhEmi'][:10]
    dia_emi = data_emi[8:]
    mes_emi = data_emi[5:7]
    ano_emi = data_emi[0:4]
    data_emissao = f'{dia_emi}/{mes_emi}/{ano_emi}'
    tipo_op = ide['tpNF']
    local_op = ide['idDest']
    try:
        cnpj_emitente = emit['CNPJ']
    except:
        cnpj_emitente = emit['CPF']
    uf_emitente = emit['enderEmit']['UF']
    ie_emitente = emit['IE']
    if 'dest' in dic_nota:
        dest = dic_nota['dest']
        try:
            cnpj_destinatario = dest['CNPJ']
        except:
            cnpj_destinatario = dest['CPF']
        if 'enderDest' in dest:
            uf_destinatario = dest['enderDest']['UF']
        else:
            uf_destinatario = 'Sem info'
        if 'IE' in dest:
            ie_destinatario = dest['IE']
        else:
            ie_destinatario = 'ISENTO'
    else:
        cnpj_destinatario = 'Sem info'
        uf_destinatario = 'Sem info'
        ie_destinatario = 'Sem info'
    bc_icms = tot['vBC']
    valor_icms = tot['vICMS']
    icms_desonerado = tot['vICMSDeson']
    fcp = tot['vFCP']
    fcp_st = tot['vFCPST']
    bc_icms_st = tot['vBCST']
    valor_icms_st = tot['vST']
    valorProd = tot['vProd']
    valorFrete = tot['vFrete']
    valorSeg = tot['vSeg']
    valorDesc = tot['vDesc']
    valorIPI = tot['vIPI']
    valorPis = tot['vPIS']
    valorCofins = tot['vCOFINS']
    valorNF = tot['vNF']
    if tipo_op == '0':
        tipo_operacao = 'Entrada'
    else:
        tipo_operacao = 'Saída'

    if local_op == '1':
        operacao = 'Operação Interna'
    elif local_op == '2':
        operacao = 'Operação Interestadual'
    else:
        operacao = 'Operação com Exterior'

    nfce = {
        'Número da NF': nf,
        'Série': serie,
        'Data de Emissao': data_emissao,
        'Chave de Acesso': chave,
        'Tipo de Operacao': tipo_operacao,
        'Operação': operacao,
        'CNPJ/CPF do Emitente': cnpj_emitente,
        'IE Emitente': ie_emitente,
        'UF Emitente': uf_emitente,
        'CNPJ/CPF do destinatario': cnpj_destinatario,
        'IE Destinatario': ie_destinatario,
        'UF Destinatario': uf_destinatario,
        'Valor dos Produtos': float(valorProd),
        'Valor do Desconto': float(valorDesc),
        'Valor do Frete': float(valorFrete),
        'Valor do Seguro': float(valorSeg),
        'Valor da NF': float(valorNF),
        'Base do ICMS': float(bc_icms),
        'Valor do ICMS': float(valor_icms),
        'Base ICMS ST': float(bc_icms_st),
        'Valor ICMS ST': float(valor_icms_st),
        'Valor do IPI': float(valorIPI),
        'Valor ICMS desonerado': float(icms_desonerado),
        'Valor fcp': float(fcp),
        'Valor fcp ST': float(fcp_st),
        'Valor Pis': float(valorPis),
        'Valor Cofins': float(valorCofins),
    }
    return nfce


def ler_xml_cte(nota):
    with open(nota, 'rb') as arquivo:
        documento = xmltodict.parse(arquivo)
    dic_cte = documento['cteProc']['CTe']['infCte']
    ide = documento['cteProc']['CTe']['infCte']['ide']
    emit = documento['cteProc']['CTe']['infCte']['emit']
    rem = documento['cteProc']['CTe']['infCte']['rem']
    dest = documento['cteProc']['CTe']['infCte']['dest']
    valor = documento['cteProc']['CTe']['infCte']['vPrest']
    mod_icms = dic_cte['imp']['ICMS']

    chave = dic_cte['@Id'][3:]
    numero_cte = ide['nCT']
    serie = ide['serie']
    data_emi = ide['dhEmi']
    dia_emi = data_emi[8:10]
    mes_emi = data_emi[5:7]
    ano_emi = data_emi[0:4]
    data_emissao = f'{dia_emi}/{mes_emi}/{ano_emi}'
    cfop = ide['CFOP']
    municipio_inicio = ide['xMunIni']
    uf_inicio = ide['UFIni']
    municipio_final = ide['xMunFim']
    uf_final = ide['UFFim']
    try:
        cnpj_transportadora = emit['CNPJ']
    except:
        cnpj_transportadora = emit['CPF']
    if 'IE' in emit:
        ie_transportadora = emit['IE']
    else:
        ie_transportadora = 'ISENTO'
    try:
        cnpj_remetente = rem['CNPJ']
    except:
        cnpj_remetente = rem['CPF']
    if 'IE' in rem:
        ie_remetente = rem['IE']
    else:
        ie_remetente = 'ISENTO'
    try:
        cnpj_destinatario = dest['CNPJ']
    except:
        cnpj_destinatario = dest['CPF']
    if 'IE' in dest:
        ie_destinatario = dest['IE']
    else:
        ie_destinatario = 'ISENTO'
    valor_frete = valor['vTPrest']

    if 'infCTeNorm' in dic_cte:
        valor_carga = dic_cte['infCTeNorm']['infCarga']['vCarga']
    else:
        valor_carga = 0.00
    if 'ICMS00' in mod_icms:
        cst_icms = mod_icms['ICMS00']['CST']
        bc_icms = mod_icms['ICMS00']['vBC']
        icms = mod_icms['ICMS00']['vICMS']

    elif 'ICMS20' in mod_icms:
        cst_icms = mod_icms['ICMS20']['CST']
        bc_icms = mod_icms['ICMS20']['vBC']
        icms = mod_icms['ICMS20']['vICMS']

    elif 'ICMS45' in mod_icms:
        cst_icms = mod_icms['ICMS45']['CST']
        bc_icms = 0.00
        icms = 0.00

    elif 'ICMS60' in mod_icms:
        cst_icms = mod_icms['ICMS60']['CST']
        bc_icms = 0.00
        icms = 0.00

    elif 'ICMS90' in mod_icms:
        cst_icms = mod_icms['ICMS90']['CST']
        if 'vBC' in mod_icms['ICMS90']:
            bc_icms = mod_icms['ICMS90']['vBC']
        else:
            bc_icms = 0.00
        if 'vICMS' in mod_icms['ICMS90']:
            icms = mod_icms['ICMS90']['vICMS']
        else:
            icms = 0.00

    elif 'ICMSOutraUF' in mod_icms:
        cst_icms = mod_icms['ICMSOutraUF']['CST']
        bc_icms = 0.00
        icms = 0.00

    elif 'ICMSSN' in mod_icms:
        cst_icms = mod_icms['ICMSSN']['CST']
        bc_icms = 0.00
        icms = 0.00

    elif 'ICMSUFFim' in mod_icms:
        cst_icms = 'Sem CST - UF Fim'
        bc_icms = 0.00
        icms = 0.00

    cte = {
        'Número CTE': numero_cte,
        'Série': serie,
        'Data de Emissao': data_emissao,
        'Chave de Acesso': chave,
        'CFOP': str(cfop),
        'Município de Origem': str(municipio_inicio),
        'UF de Origem': str(uf_inicio),
        'Município de Destino': str(municipio_final),
        'UF de Destino': str(uf_final),
        'CNPJ transportadora': str(cnpj_transportadora),
        'IE transportadora': str(ie_transportadora),
        'CNPJ/CPF do Remetente': str(cnpj_remetente),
        'IE Remetente': str(ie_remetente),
        'CNPJ/CPF do destinatario': str(cnpj_destinatario),
        'IE Destinatario': str(ie_destinatario),
        'Total Frete': float(valor_frete),
        'CST ICMS': str(cst_icms),
        'BC ICMS': float(bc_icms),
        'ICMS': float(icms),
        'Valor da Carga': float(valor_carga),

    }
    return cte


def ler_evento_nfe(nota):
    with open(nota, 'rb') as arquivo:
        documento = xmltodict.parse(arquivo)
    if 'procEventoNFe' in documento:
        chave = documento['procEventoNFe']['evento']['infEvento']['chNFe']
        eventos = documento['procEventoNFe']['evento']['infEvento']['tpEvento']
        dhEvento = documento['procEventoNFe']['evento']['infEvento']['dhEvento']
        ano = dhEvento[0:4]
        mes = dhEvento[5:7]
        dia = dhEvento[8:10]
        data_even = f'{dia}/{mes}/{ano}'

        if '110111' in eventos:
            evento = 'Nota Cancelada'
        elif '110110' in eventos:
            carta = documento['procEventoNFe']['evento']['infEvento']['detEvento']['xCorrecao']
            evento = f'Descrição Carta de Correção: {carta}'
        else:
            evento = f'Outro Evento - Código: {eventos}'

        xml_evento = {
            'Chave de Acesso': chave,
            'Data do Evento': data_even,
            'Evento': evento,
        }

    elif 'procEventoCTe' in documento:
        chave = documento['procEventoCTe']['eventoCTe']['infEvento']['chCTe']
        eventos = documento['procEventoCTe']['eventoCTe']['infEvento']['tpEvento']
        dhEvento = documento['procEventoCTe']['eventoCTe']['infEvento']['dhEvento']
        ano = dhEvento[0:4]
        mes = dhEvento[5:7]
        dia = dhEvento[8:10]
        data_even = f'{dia}/{mes}/{ano}'

        if '110111' in eventos:
            evento = 'Nota Cancelada'
        elif '110110' in eventos:
            carta = documento['procEventoCTe']['eventoCTe']['infEvento']['detEvento']['evCCeCTe']['infCorrecao'][
                'valorAlterado']
            evento = f'Descrição Carta de Correção CTE: {carta}'
        else:
            evento = f'Outro Evento - Código: {eventos}'

        xml_evento = {
            'Chave de Acesso': chave,
            'Data do Evento': data_even,
            'Evento': evento,
        }

    return xml_evento


janela = Tk()
caminho = tkinter.filedialog.askdirectory(title='Selecione a pasta com os arquivos XML')
janela.destroy()
lista_arquivos = os.listdir(caminho)
lista_arquivos_xml = []
for arquivo in lista_arquivos:
    if '.xml' in arquivo.lower():
        lista_arquivos_xml.append(arquivo)

lista_nfe_dic = []
lista_nfce_dic = []
lista_cte_dic = []
lista_evento_dic = []

for nota in tqdm(lista_arquivos_xml):
    try:
        with open(f'{caminho}/{nota}', 'rb') as arquivo:
            nome = arquivo.name
            if '.xml' in nome.lower():
                documento = xmltodict.parse(arquivo)
                if 'procInutNFe' in documento:
                    pass
                elif 'ProcInutNFe' in documento:
                    pass
                elif 'retInutNFe' in documento:
                    pass
                elif 'inutNFe' in documento:
                    pass
                elif 'retEnvEvento' in documento:
                    pass
                elif 'resEvento' in documento:
                    pass
                elif 'procEventoNFe' in documento:
                    lista_evento_dic.append(ler_evento_nfe(f'{caminho}/{nota}'))
                elif 'procEventoCTe' in documento:
                    lista_evento_dic.append(ler_evento_nfe(f'{caminho}/{nota}'))
                elif 'cteProc' in documento:
                    lista_cte_dic.append(ler_xml_cte(f'{caminho}/{nota}'))
                elif 'NFe' in documento:
                    lista_nfe_dic.append(ler_xml_nfe(f'{caminho}/{nota}'))
                elif 'NFeLog' in documento:
                    lista_nfe_dic.append(ler_xml_nfe_receita(f'{caminho}/{nota}'))
                elif documento['nfeProc']['NFe']['infNFe']['ide']['mod'] == '55':
                    lista_nfe_dic.append(ler_xml_nfe(f'{caminho}/{nota}'))
                else:
                    lista_nfce_dic.append(ler_xml_nfce(f'{caminho}/{nota}'))
            else:
                pass
    except:
        with open(f'{caminho}/{nota}', 'rb') as arquivo:
            nome = arquivo.name
            if '.xml' in nome.lower():
                documento = xmltodict.parse(arquivo, encoding='ANSI')
                if 'procInutNFe' in documento:
                    pass
                elif 'ProcInutNFe' in documento:
                    pass
                elif 'retInutNFe' in documento:
                    pass
                elif 'inutNFe' in documento:
                    pass
                elif 'retEnvEvento' in documento:
                    pass
                elif 'resEvento' in documento:
                    pass
                elif 'procEventoNFe' in documento:
                    lista_evento_dic.append(ler_evento_nfe(f'{caminho}/{nota}'))
                elif 'procEventoCTe' in documento:
                    lista_evento_dic.append(ler_evento_nfe(f'{caminho}/{nota}'))
                elif 'cteProc' in documento:
                    lista_cte_dic.append(ler_xml_cte(f'{caminho}/{nota}'))
                elif 'NFe' in documento:
                    lista_nfe_dic.append(ler_xml_nfe(f'{caminho}/{nota}'))
                elif 'NFeLog' in documento:
                    lista_nfe_dic.append(ler_xml_nfe_receita(f'{caminho}/{nota}'))
                elif documento['nfeProc']['NFe']['infNFe']['ide']['mod'] == '55':
                    lista_nfe_dic.append(ler_xml_nfe(f'{caminho}/{nota}'))
                else:
                    lista_nfce_dic.append(ler_xml_nfce(f'{caminho}/{nota}'))
            else:
                pass

if lista_nfe_dic != []:
    tabela_nfe = pd.DataFrame.from_dict(lista_nfe_dic)
    tabela_nfe.to_excel(f'{caminho}/+ NFE.xlsx', index=False)
if lista_nfce_dic != []:
    tabela_nfce = pd.DataFrame.from_dict(lista_nfce_dic)
    tabela_nfce.to_excel(f'{caminho}/+ NFCE.xlsx', index=False)
if lista_cte_dic != []:
    tabela_cte = pd.DataFrame.from_dict(lista_cte_dic)
    tabela_cte.to_excel(f'{caminho}/+ CTE.xlsx', index=False)
if lista_evento_dic != []:
    tabela_evento = pd.DataFrame.from_dict(lista_evento_dic)
    tabela_evento.to_excel(f'{caminho}/+ Eventos.xlsx', index=False)
