import xmltodict
import os
import pandas as pd
from datetime import datetime
from dateutil import parser
import xlsxwriter

# Nome do arquivo Excel que você quer usar
NOME_ARQUIVO_EXCEL = "NotasFiscais.xlsx"

# Função para transformar data em formato dd/mm/aaaa removendo o timezone
def parse_data_emissao(data_str):
    try:
        dt = parser.parse(data_str)
        if dt.tzinfo is not None:
            dt = dt.replace(tzinfo=None)
        return dt
    except Exception:
        try:
            return datetime.strptime(data_str, "%d/%m/%Y")
        except Exception:
            print(f"[AVISO] Data inválida ignorada: {data_str}")
            return None

def pegar_infos(nome_arquivo):
    print(f"{nome_arquivo} processado...")
    with open(f"nfs/{nome_arquivo}", "rb") as arquivo_xml:
        dic_arquivo = xmltodict.parse(arquivo_xml)
        infos_nf = {}
        numero_NF = None
        data_emissao_str = None
        CNPJ_emissor = None
        razao_emissor = None
        DOC_tomador = None
        razao_tomador = None
        texto_corpo = ""
        texto_adic = ""
        valor_bruto = None
        num_pedido = None
        num_contrato = None
        data_emissao = None

        if "xmlNfpse" in dic_arquivo:
            infos_nf = dic_arquivo["xmlNfpse"]
            numero_NF = infos_nf["numeroSerie"]
            data_emissao_str = infos_nf["dataEmissao"]
            CNPJ_emissor = infos_nf["cnpjPrestador"]
            razao_emissor = infos_nf["razaoSocialPrestador"]
            DOC_tomador = infos_nf["identificacaoTomador"]
            razao_tomador = infos_nf["razaoSocialTomador"]
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
                                break
                        if num_pedido:
                            break
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

        num_pedido_formatado = int(num_pedido) if num_pedido and num_pedido.isdigit() else None
        num_contrato_formatado = int(num_contrato) if num_contrato and num_contrato.isdigit() else None
        valor_formatado = float(valor_bruto.replace(",", ".")) if isinstance(valor_bruto, str) else float(valor_bruto)
        data_emissao = parse_data_emissao(data_emissao_str) if data_emissao_str else None
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

pasta_nfs = './nfs'
lista_arquivos = os.listdir(pasta_nfs)
qt_arquivos = 0
colunas = ["Número NF", "Data Emissão", "CNPJ Prestador", "Razão Social Prestador", "CNPJ/CPF Tomador", "Razão Social Tomador", "Contrato", "Pedido", "Valor", "Nome do Arquivo xml"]
novos_valores = []

try:
    df_existente = pd.read_excel(NOME_ARQUIVO_EXCEL)
except FileNotFoundError:
    df_existente = pd.DataFrame(columns=colunas)

for arquivo in lista_arquivos:
    if arquivo.lower().endswith(".xml"):
        infos = pegar_infos(arquivo)
        if infos:
            existe = ((df_existente['Número NF'] == infos['Número NF']) &
                      (df_existente['CNPJ Prestador'] == infos['CNPJ Prestador'])).any()
            if not existe:
                novos_valores.append(infos)
                qt_arquivos += 1
            else:
                print(f"[AVISO] NF {infos['Número NF']} do CNPJ {infos['CNPJ Prestador']} já existe na planilha.")
        else:
            print(f"[AVISO] Arquivo '{arquivo}' não pôde ser processado.")

if novos_valores:
    df_novos = pd.DataFrame(novos_valores)
    df_final = pd.concat([df_existente, df_novos], ignore_index=True)
else:
    df_final = df_existente

# Usando xlsxwriter para formatação
with pd.ExcelWriter(NOME_ARQUIVO_EXCEL, engine="xlsxwriter", date_format="dd/mm/yyyy") as writer:
    df_final.to_excel(writer, sheet_name='Notas Fiscais', startrow=1, index=False, header=False)

    workbook = writer.book
    worksheet = writer.sheets['Notas Fiscais']

    # Formato para o cabeçalho
    header_format = workbook.add_format({
        'bold': True,
        'text_wrap': True,
        'align': 'center',
        'valign': 'vcenter',
        'fg_color': '#D3D3D3',
        'border': 1
    })

    # Formato para as células de dados (centralizado e com borda fina)
    cell_format = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'border': 1
    })

    # Formatos específicos para as colunas
    number_format = workbook.add_format({'num_format': '0', 'align': 'center', 'valign': 'vcenter', 'border': 1})
    date_format = workbook.add_format({'num_format': 'dd/mm/yy', 'align': 'center', 'valign': 'vcenter', 'border': 1})
    cnpj_format = workbook.add_format({'num_format': '00"."000"."000"/"0000-00', 'align': 'center', 'valign': 'vcenter', 'border': 1})
    cpf_format = workbook.add_format({'num_format': '000"."000"."000"-"00', 'align': 'center', 'valign': 'vcenter', 'border': 1})
    valor_format = workbook.add_format({'num_format': '#,##0.00', 'align': 'right', 'valign': 'vcenter', 'border': 1})

    # Escreve o cabeçalho com o formato
    for col_num, value in enumerate(df_final.columns.values):
        worksheet.write(0, col_num, value, header_format)

    # Escreve os dados com o formato das células e formatos específicos das colunas
    for row_num, row_data in df_final.iterrows():
        for col_num, value in enumerate(row_data):
            if pd.isna(value):
                worksheet.write(row_num + 1, col_num, "", cell_format)
            else:
                column_name = df_final.columns[col_num]
                if column_name == "Número NF":
                    worksheet.write_number(row_num + 1, col_num, value, number_format)
                elif column_name == "Data Emissão":
                    if isinstance(value, datetime):
                        worksheet.write_datetime(row_num + 1, col_num, value, date_format)
                    else:
                        worksheet.write(row_num + 1, col_num, "", cell_format)
                elif column_name == "CNPJ Prestador":
                    if pd.notna(value):
                        worksheet.write_number(row_num + 1, col_num, value, cnpj_format)
                    else:
                        worksheet.write(row_num + 1, col_num, "", cell_format)
                elif column_name == "CNPJ/CPF Tomador":
                    if pd.notna(value):
                        str_value = str(int(value)) if isinstance(value, (int, float)) else str(value)
                        if len(str_value) == 11:
                            worksheet.write_number(row_num + 1, col_num, int(value), cpf_format)
                        #elif len(str_value) == 11:
                        #    worksheet.write_string(row_num + 1, col_num, str_value, cpf_format)
                        else:
                            worksheet.write_number(row_num + 1, col_num, value, cnpj_format)
                    else:
                        worksheet.write(row_num + 1, col_num, "", cell_format)
                elif column_name == "Valor":
                    worksheet.write_number(row_num + 1, col_num, value, valor_format)
                else:
                    worksheet.write(row_num + 1, col_num, value, cell_format)

    # Ajusta a largura das colunas automaticamente
    for i, col in enumerate(df_final.columns):
        column_len = len(col) + 2 # Largura mínima baseada no cabeçalho
        content_len = df_final[col].astype(str).str.len().max() + 2
        if col == "Número NF":
            worksheet.set_column(i, i, max(column_len, content_len))
        elif col == "Data Emissão":
            worksheet.set_column(i, i, max(column_len, 10 + 2)) # "dd/mm/yyyy" tem 10 caracteres
        elif col == "CNPJ Prestador":
            worksheet.set_column(i, i, max(column_len, 18 + 2)) # "00.000.000/0000-00" tem 18 caracteres
        elif col == "CNPJ/CPF Tomador":
            worksheet.set_column(i, i, max(column_len, 18 + 2)) # Considera o máximo entre CNPJ e CPF formatado
        else:
            worksheet.set_column(i, i, max(column_len, content_len))

print(f"Total de {qt_arquivos} novos arquivos processados e adicionados à planilha '{NOME_ARQUIVO_EXCEL}' ...")