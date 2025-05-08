import xmltodict
import os
import pandas as pd
from datetime import datetime
from dateutil import parser

# Nome do arquivo Excel que você quer usar
NOME_ARQUIVO_EXCEL = "NotasFiscais.xlsx"

# Função para transformar data em formato dd/mm/aaaa removendo o timezone
def parse_data_emissao(data_str):
    try:
        # Usa parser inteligente do dateutil (aceita ISO 8601, datas com TZ, etc.)
        dt = parser.parse(data_str)
        if dt.tzinfo is not None:
            dt = dt.replace(tzinfo=None)
        return dt
    except Exception:
        try:
            # Tenta formato brasileiro como fallback
            return datetime.strptime(data_str, "%d/%m/%Y")
        except Exception:
            print(f"[AVISO] Data inválida ignorada: {data_str}")
            return None  # ou lançar erro, ou deixar em branco

# Função para pegar informações do arquivo XML
def pegar_infos(nome_arquivo):
    print(f"{nome_arquivo} processado...")
    with open(f"nfs/{nome_arquivo}", "rb") as arquivo_xml:
        dic_arquivo = xmltodict.parse(arquivo_xml)
        infos_nf = {}
        if "xmlNfpse" in dic_arquivo:
            infos_nf = dic_arquivo["xmlNfpse"]
            numero_NF = infos_nf["numeroSerie"]
            data_emissao_str = infos_nf["dataEmissao"]
            CNPJ_emissor = infos_nf["cnpjPrestador"]
            razao_emissor = infos_nf["razaoSocialPrestador"]
            DOC_tomador = infos_nf["identificacaoTomador"]
            razao_tomador = infos_nf["razaoSocialTomador"]
            texto_corpo = ""
            if isinstance(infos_nf["itensServico"]["itemServico"], list):
                for item in infos_nf["itensServico"]["itemServico"]:
                    texto_corpo += item["descricaoServico"] + "\n"
            else:
                texto_corpo = infos_nf["itensServico"]["itemServico"]["descricaoServico"]
            texto_adic = infos_nf.get("dadosAdicionais", "")
            valor_bruto = infos_nf["valorTotalServicos"]
        elif "nfeProc" in dic_arquivo:
            infos_nf = dic_arquivo["nfeProc"]["NFe"]["infNFe"]
            numero_NF = infos_nf["ide"]["nNF"]
            data_emissao_str = infos_nf["ide"]["dhEmi"]
            CNPJ_emissor = infos_nf["emit"]["CNPJ"]
            razao_emissor = infos_nf["emit"]["xNome"]
            DOC_tomador = infos_nf["dest"].get("CNPJ", infos_nf["dest"].get("CPF", ""))
            razao_tomador = infos_nf["dest"]["xNome"]
            texto_corpo = infos_nf["infAdic"].get("infCpl", "")
            texto_adic = ""
            if "infAdFisco" in infos_nf["infAdic"]:
                texto_adic = infos_nf["infAdic"]["infAdFisco"]
            elif "infAdProd" in infos_nf["det"]:
                texto_adic = infos_nf["det"]["infAdProd"]
            valor_bruto = infos_nf["total"]["ICMSTot"]["vNF"]
        elif "RetornoConsulta" in dic_arquivo:
            infos_nf = dic_arquivo["RetornoConsulta"]["NFe"]
            numero_NF = infos_nf["ChaveNFe"]["NumeroNFe"]
            data_emissao_str = infos_nf["DataEmissaoNFe"]
            CNPJ_emissor = infos_nf["CPFCNPJPrestador"]["CNPJ"]
            razao_emissor = infos_nf["RazaoSocialPrestador"]
            DOC_tomador = infos_nf["CPFCNPJTomador"]["CNPJ"]
            razao_tomador = infos_nf["RazaoSocialTomador"]
            texto_corpo = infos_nf["Discriminacao"]
            texto_adic = ""
            valor_bruto = infos_nf["ValorServicos"]
        elif "NFe" in dic_arquivo:
            infos_nf = dic_arquivo["NFe"]
            numero_NF = infos_nf["ChaveNFe"]["NumeroNFe"]
            data_emissao_str = infos_nf["DataEmissaoNFe"]
            CNPJ_emissor = infos_nf["CPFCNPJPrestador"]["CNPJ"]
            razao_emissor = infos_nf["RazaoSocialPrestador"]
            DOC_tomador = infos_nf["CPFCNPJTomador"]["CNPJ"]
            razao_tomador = infos_nf["RazaoSocialTomador"]
            texto_corpo = infos_nf["Discriminacao"]
            texto_adic = ""
            valor_bruto = infos_nf["ValorServicos"]
        elif "nfe" in dic_arquivo:
            infos_nf = dic_arquivo["nfe"]["NotaFiscal"]
            numero_NF = infos_nf["NfeCabecario"]["numeroNota"]
            data_emissao_str = infos_nf["NfeCabecario"]["dataEmissao"]
            CNPJ_emissor = infos_nf["DadosPrestador"]["documento"]
            razao_emissor = infos_nf["DadosPrestador"]["razaoSocial"]
            DOC_tomador = infos_nf["DadosTomador"]["documento"]
            razao_tomador = infos_nf["DadosTomador"]["razaoSocial"]
            texto_corpo = infos_nf["DetalhesServico"]["descricao"]
            texto_adic = ""
            valor_bruto = infos_nf["DetalhesServico"]["valorServico"]
        elif "ConsultarNfseResposta" in dic_arquivo:
            infos_nf = dic_arquivo["ConsultarNfseResposta"]["ListaNfse"]["CompNfse"]["Nfse"]["InfNfse"]
            numero_NF = infos_nf["Numero"]
            data_emissao_str = infos_nf["DataEmissao"]
            CNPJ_emissor = infos_nf["PrestadorServico"]["IdentificacaoPrestador"]["Cnpj"]
            razao_emissor = infos_nf["PrestadorServico"]["RazaoSocial"]
            DOC_tomador = infos_nf["TomadorServico"]["IdentificacaoTomador"]["CpfCnpj"]["Cnpj"]
            razao_tomador = infos_nf["TomadorServico"]["RazaoSocial"]
            texto_corpo = infos_nf["Servico"]["Discriminacao"]
            texto_adic = infos_nf["Servico"].get("OutrasInformacoes", "")
            valor_bruto = infos_nf["Servico"]["Valores"]["ValorServicos"]
        elif "ns4:Nfse" in dic_arquivo:
            infos_nf = dic_arquivo["ns4:Nfse"]["ns4:InfNfse"]
            numero_NF = infos_nf["ns4:Numero"]
            data_emissao_str = infos_nf["ns4:DataEmissao"]
            CNPJ_emissor = infos_nf["ns4:PrestadorServico"]["ns4:IdentificacaoPrestador"]["ns4:Cnpj"]
            razao_emissor = infos_nf["ns4:PrestadorServico"]["ns4:RazaoSocial"]
            DOC_tomador = infos_nf["ns4:TomadorServico"]["ns4:IdentificacaoTomador"]["ns4:CpfCnpj"]["ns4:Cnpj"]
            razao_tomador = infos_nf["ns4:TomadorServico"]["ns4:RazaoSocial"]
            texto_corpo = infos_nf["ns4:Servico"]["ns4:Discriminacao"]
            texto_adic = ""
            valor_bruto = infos_nf["ns4:Servico"]["ns4:Valores"]["ns4:ValorServicos"]
        elif "CompNfse" in dic_arquivo:
            infos_nf = dic_arquivo["CompNfse"]["Nfse"]["InfNfse"]
            numero_NF = infos_nf["Numero"]
            data_emissao_str = infos_nf["DataEmissao"]
            CNPJ_emissor = infos_nf["DeclaracaoPrestacaoServico"]["InfDeclaracaoPrestacaoServico"]["Prestador"]["CpfCnpj"]["Cnpj"]
            razao_emissor = infos_nf["PrestadorServico"]["RazaoSocial"]
            DOC_tomador = infos_nf["DeclaracaoPrestacaoServico"]["InfDeclaracaoPrestacaoServico"]["TomadorServico"]["IdentificacaoTomador"]["CpfCnpj"]["Cnpj"]
            razao_tomador = infos_nf["DeclaracaoPrestacaoServico"]["InfDeclaracaoPrestacaoServico"]["TomadorServico"]["RazaoSocial"]
            texto_corpo = infos_nf["DeclaracaoPrestacaoServico"]["InfDeclaracaoPrestacaoServico"]["Servico"]["Discriminacao"]
            texto_adic = ""
            valor_bruto = infos_nf["DeclaracaoPrestacaoServico"]["InfDeclaracaoPrestacaoServico"]["Servico"]["Valores"]["ValorServicos"]
        else:
            print("Modelo XML Diferente!")
            return None

        descricao = [texto_corpo, texto_adic]
        pedidos = ["42000", "43000", "45000"]
        contratos = ["47000", "48000"]
        num_pedido = None
        num_contrato = None

        for item_na_descricao in descricao:
            if isinstance(item_na_descricao, str):
                for pedido in pedidos:
                    for palavra in item_na_descricao.split():
                        if pedido in palavra:
                            num_pedido = palavra[palavra.find(pedido):palavra.find(pedido)+10]
                for contrato in contratos:
                    for palavra in item_na_descricao.split():
                        if contrato in palavra:
                            num_contrato = palavra[palavra.find(contrato):palavra.find(contrato)+10]
            elif isinstance(item_na_descricao, list):
                for item_da_lista in item_na_descricao:
                    if isinstance(item_da_lista, str):
                        for pedido in pedidos:
                            if pedido in item_da_lista:
                                num_pedido = item_da_lista
                                break # Se encontrou um pedido, pode sair do loop interno
                        if num_pedido:
                            break # Se encontrou um pedido, pode sair do loop externo
                    elif isinstance(item_da_lista, dict):
                        for vlr in item_da_lista.values():
                            if isinstance(vlr, str):
                                for pedido in pedidos:
                                    if pedido in vlr:
                                        num_pedido = vlr
                                        break
                                if num_pedido:
                                    break
                        if num_pedido:
                            break

        num_pedido_formatado = int(num_pedido) if num_pedido and num_pedido.isdigit() else ""
        num_contrato_formatado = int(num_contrato) if num_contrato and num_contrato.isdigit() else ""
        valor_formatado = float(valor_bruto.replace(",", ".")) if isinstance(valor_bruto, str) else float(valor_bruto)
        data_emissao = parse_data_emissao(data_emissao_str)
        numero_NF_formatado = int(numero_NF) if numero_NF and numero_NF.isdigit() else None
        CNPJ_emissor_formatado = int(CNPJ_emissor) if CNPJ_emissor and CNPJ_emissor.isdigit() else None
        DOC_tomador_formatado = int(DOC_tomador) if DOC_tomador and DOC_tomador.isdigit() else None

        return {
            "Número NF": numero_NF_formatado,
            "Data Emissão": data_emissao,
            "CNPJ Prestador": CNPJ_emissor_formatado,
            "Razão Social Prestador": razao_emissor,
            "CNPJ/CPF Tomador": DOC_tomador_formatado,
            "Razão Social Tomador": razao_tomador,
            "Contrato": num_contrato_formatado,
            "Pedido": num_pedido_formatado,
            "Valor": valor_formatado,
            "Nome do Arquivo xml": nome_arquivo
        }

# Caminho da pasta com os arquivos XML
pasta_nfs = './nfs'
lista_arquivos = os.listdir(pasta_nfs)
qt_arquivos = 0
colunas = ["Número NF", "Data Emissão", "CNPJ Prestador", "Razão Social Prestador", "CNPJ/CPF Tomador", "Razão Social Tomador", "Contrato", "Pedido", "Valor", "Nome do Arquivo xml"]

try:
    # Tenta carregar a planilha existente
    df_existente = pd.read_excel(NOME_ARQUIVO_EXCEL)
except FileNotFoundError:
    # Se o arquivo não existir, cria um DataFrame vazio
    df_existente = pd.DataFrame(columns=colunas)

novos_valores = []

for arquivo in lista_arquivos:
    if arquivo.endswith(".xml"):
        infos = pegar_infos(arquivo)
        if infos:
            # Verifica se a nota fiscal já existe na planilha pelo número e CNPJ do prestador
            existe = ((df_existente['Número NF'] == infos['Número NF']) &
                      (df_existente['CNPJ Prestador'] == infos['CNPJ Prestador'])).any()
            if not existe:
                novos_valores.append(infos)
                qt_arquivos += 1
            else:
                print(f"[AVISO] NF {infos['Número NF']} do CNPJ {infos['CNPJ Prestador']} já existe na planilha.")

if novos_valores:
    df_novos = pd.DataFrame(novos_valores)
    df_final = pd.concat([df_existente, df_novos], ignore_index=True)
else:
    df_final = df_existente

with pd.ExcelWriter(NOME_ARQUIVO_EXCEL, engine="xlsxwriter", date_format="dd/mm/yyyy") as writer:
    df_final.to_excel(writer, index=False)

print(f"Total de {qt_arquivos} novos arquivos processados e adicionados (se não existentes) à planilha '{NOME_ARQUIVO_EXCEL}'.")