import xmltodict
import os
import pandas as pd
from datetime import datetime
from dateutil import parser

#Função para transformar data em formato dd/mm/aaaa removendo o timezone
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
#Função para pegar informações do arquivo
def pegar_infos(nome_arquivo, valores):
	print(f"{nome_arquivo} processado...")
	with open(f"nfs/{nome_arquivo}", "rb") as arquivo_xml:
		dic_arquivo = xmltodict.parse(arquivo_xml)
		#print(json.dumps(dic_arquivo, indent=4))
		if "xmlNfpse" in dic_arquivo:
			infos_nf = dic_arquivo["xmlNfpse"]
			numero_NF = infos_nf["numeroSerie"]
			data_emissao_str = infos_nf["dataEmissao"]
			CNPJ_emissor = infos_nf["cnpjPrestador"]
			razao_emissor = infos_nf["razaoSocialPrestador"]
			DOC_tomador = infos_nf["identificacaoTomador"]
			razao_tomador = infos_nf["razaoSocialTomador"]
			if isinstance(infos_nf["itensServico"]["itemServico"], list):
				texto_corpo = ""
				for item in infos_nf["itensServico"]["itemServico"]:
					texto_corpo += item["descricaoServico"] + "\n"
			else:
				texto_corpo = infos_nf["itensServico"]["itemServico"]["descricaoServico"]
			texto_adic = infos_nf["dadosAdicionais"]
			valor_bruto = infos_nf["valorTotalServicos"]
		elif "nfeProc" in dic_arquivo:
			infos_nf = dic_arquivo["nfeProc"]["NFe"]["infNFe"]
			numero_NF = infos_nf["ide"]["nNF"]
			data_emissao_str = infos_nf["ide"]["dhEmi"]
			CNPJ_emissor = infos_nf["emit"]["CNPJ"]
			razao_emissor = infos_nf["emit"]["xNome"]
			if "CNPJ" in infos_nf["dest"]:
				DOC_tomador = infos_nf["dest"]["CNPJ"]
			else:
				DOC_tomador = infos_nf["dest"]["CPF"]
			razao_tomador = infos_nf["dest"]["xNome"]
			texto_corpo = infos_nf["infAdic"]["infCpl"]
			if "infAdFisco" in infos_nf["infAdic"]:
				texto_adic = infos_nf["infAdic"]["infAdFisco"]
			elif "infAdProd" in infos_nf["det"]:
				texto_adic = infos_nf["det"]["infAdProd"]
			else:
				texto_adic = ""
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
			texto_adic = ["OutrasInformacoes"]
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
		#Lista de campos trazidos do XML
		descricao = [texto_corpo, texto_adic]
		pedidos = ["42000", "43000", "45000"]
		contratos = ["47000", "48000"]
		num_pedido = None
		num_contrato = None
		#Percorre os itens contidos na descrição (campos extraídos do XML) e a cada iteração, verifica se o tipo de dado é String.
		for item_na_descricao in descricao:
			if type(item_na_descricao) == str:
			#Caso o tipo seja String, verifica se cada palavra do item na descricao contém o texto do item da lista de pedidos.
				for pedido in pedidos:
					for palavra in item_na_descricao.split():
						if pedido in palavra:
							#Caso contenha, localiza o indice do item da lista de pedidos dentro da palavra e pega 10 caracteres a partir deste indice e
							#Atribui o texto extraído à variável num_pedido
							num_pedido = palavra[palavra.find(pedido):palavra.find(pedido)+10]
				#Também verifica se cada palavra do item na descricao contém o texto do item da lista de contratos.
				for contrato in contratos:
					for palavra in item_na_descricao.split():
						if contrato in palavra:
							#Caso contenha, localiza o indice do item da lista de contratos dentro da palavra e pega 10 caracteres a partir deste indice e
							#Atribui o texto extraído à variável num_contrato
							num_contrato = palavra[palavra.find(contrato):palavra.find(contrato)+10]


							
			#Caso o tipo seja Lista, para cada item na lista, percorre cada par chave/valor do dicionário naquele item e procura o texto dentro da chave/valor
			elif type(item_na_descricao) == list: 
				for item_da_lista in item_na_descricao: 
					for vlr in item_na_descricao:
						for pedido in pedidos:
							if pedido in vlr:
								num_pedido = vlr
				#Se o tipo do item não é "str" nem "lista", imprime o texto
			else:
				mensagem = f"Tipo {type(item_na_descricao)} é diferente"
		if num_pedido == None:
			num_pedido = ""
		if num_contrato == None:
			num_contrato = ""
		#Mudando o valor obtido como String para Float
		valor_formatado = float(valor_bruto.replace(",", "."))
		#Mudando a data para formato "dd/mm/aaaa" usando a função parse_data_emissao criada no início do arquivo.
		data_emissao = parse_data_emissao(data_emissao_str)

		numero_NF_formatado = int(numero_NF)

		CNPJ_emissor_formatado = int(CNPJ_emissor)

		DOC_tomador_formatado = int(DOC_tomador)
		if num_pedido == "":
			num_pedido_formatado = ""
		else:
			num_pedido_formatado = int(num_pedido)
		if num_contrato == "":
			num_contrato_formatado = ""
		else:
			num_contrato_formatado = int(num_contrato)
		
		valores.append([numero_NF_formatado, data_emissao, CNPJ_emissor_formatado, razao_emissor, DOC_tomador_formatado, razao_tomador, num_contrato_formatado, num_pedido_formatado, valor_formatado, nome_arquivo])

lista_arquivos = os.listdir(path='./nfs')
qt_arquivos = 0

colunas = ["Número NF", "Data Emissão", "CNPJ Prestador", "Razão Social Prestador", "CNPJ/CPF Tomador", "Razão Social Tomador", "Contrato", "Pedido", "Valor", "Nome do Arquivo xml"]
valores = []

for arquivo in lista_arquivos:
	pegar_infos(arquivo, valores)
	qt_arquivos += 1
print("Total de", qt_arquivos, "arquivos processados")
tabela = pd.DataFrame(columns=colunas, data=valores)
with pd.ExcelWriter("NotasFiscais.xlsx", engine="xlsxwriter", date_format="dd/mm/yyyy") as writer:
	tabela.to_excel(writer, index=False)