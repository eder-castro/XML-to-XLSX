import xmltodict
import os
import pandas as pd

import json

def pegar_infos(nome_arquivo):
	with open(f"xml/{nome_arquivo}", "rb") as arquivo_xml:
		dic_arquivo = xmltodict.parse(arquivo_xml)

		if "RetornoConsulta" in dic_arquivo:
			infos_nf = dic_arquivo["RetornoConsulta"]["NFe"]
		elif "nfeProc" in dic_arquivo:
			infos_nf = dic_arquivo["nfeProc"]["NFe"]["infNFe"]
		elif "xmlNfpse" in dic_arquivo:
			infos_nf = dic_arquivo["xmlNfpse"]
		elif "nfe" in dic_arquivo:
			infos_nf = dic_arquivo["nfe"]["NotaFiscal"]
		elif "CompNfse" in dic_arquivo:
			infos_nf = dic_arquivo["CompNfse"]["Nfse"]["InfNfse"]
		else:
			print("Modelo de XML incorreto!")

		if "xmlNfpse" in dic_arquivo:
			texto_corpo = infos_nf["itensServico"]["itemServico"]
			texto_corpo2 = infos_nf["dadosAdicionais"]
		elif "nfeProc" in dic_arquivo:
			texto_corpo = infos_nf["infAdic"]["infCpl"]
			texto_corpo2 = "Não tem Não tem Não tem Não tem Não tem Não tem"
		elif "RetornoConsulta" in dic_arquivo:
			texto_corpo = infos_nf["Discriminacao"]
			texto_corpo2 = "Não tem Não tem Não tem Não tem Não tem Não tem"
		elif "nfe" in dic_arquivo:
			texto_corpo = infos_nf["DetalhesServico"]["descricao"]
			texto_corpo2 = "Não tem Não tem Não tem Não tem Não tem Não tem"
		elif "CompNfse" in dic_arquivo:
			texto_corpo = infos_nf["DeclaracaoPrestacaoServico"]["InfDeclaracaoPrestacaoServico"]["Servico"]["Discriminacao"]
			texto_corpo2 = "Não tem Não tem Não tem Não tem Não tem Não tem"
		else:
			print("Modelo XML Diferente!")

	print(f"Informações do arquivo {nome_arquivo}")
	print("--------------------------------------------------------------------------------------")
	print(texto_corpo)
	print(texto_corpo2)
	print("--------------------------------------------------------------------------------------")

	#Lista de campos trazidos do XML
	descricao = [texto_corpo, texto_corpo2]
	pedidos = ["42000", "43000", "45000"]
	contratos = ["47000", "48000"]
	num_pedido = None
	num_contrato = None
	lista_teste = {"Ref.", "Fev/25-Contrato:", "4700000892-Pedido", "2025:", "4500012024"}
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
#	print("--------------------------------------------------------------------------------------")
	if num_pedido == None:
		num_pedido = ""
	if num_contrato == None:
		num_contrato = ""
#	print(f"Pedido = {num_pedido} e ", f"Contrato = {num_contrato}")

'''

	if type(texto_corpo) == list:
		if "45000" in texto_corpo:
			print(type(texto_corpo), "achou na lista")
	elif type(texto_corpo) == str:
			print(type(texto_corpo2), texto_corpo.find("45000"), "Achou na Str")
	else:
		print("Tipo diferente!")

	if type(texto_corpo2) == list:
		if "45000" in texto_corpo2:
			print(type(texto_corpo2), "Achou na segunda lista")
	elif type(texto_corpo2) == str:
			print(type(texto_corpo2), texto_corpo2.find("45000"), "Achou na 2 Str")
	else:
		print("Tipo diferente!")
'''


lista_arquivos = os.listdir(path='./xml')

for arquivo in lista_arquivos:
	pegar_infos(arquivo)
