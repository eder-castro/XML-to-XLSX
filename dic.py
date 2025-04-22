import re
dados = [{'Ref.Jan/25-Contrato': '4700000892-Pedido' ,'2025': '4500012024', 'CONTRATO': 'PS-0403-4700001608 -', 'PEDIDO': '4500009832;', 'aliquota': '0.02', 'baseCalculo': '3997.6', 'codigoCNAE': '6202300', 'cst': '0', 'descricaoCNAE': 'DESENVOLVIMENTO E LICENCIAMENTO DE PROGRAMAS DE COMPUTADOR CUSTOMIZAVEIS', 'descricaoServico': '(SERVICOS DE LICENCIAMENTO OU CESSAO DE DIREITO DE USO DE PROGRAMAS DE COMPUTACAO)  SERVICO CONTRATADO: LICENCIAMENTO HUBLY - COD. SERVICO 1.05 LC 116/03 - SERVICOS PRESTADOS EM FLORIANOPOLIS - INFORMACAO AO CONSUMIDOR CONFORME LEI 12741/12: 5,65%.\n :', 'idCNAE': '9102', 'quantidade': '1', 'valorTotal': '3997.6', 'valorUnitario': '3997.6'}, {'aliquota': '0.02', 'baseCalculo': '49.9', 'codigoCNAE': '6202300', 'cst': '0', 'descricaoCNAE': 'DESENVOLVIMENTO E LICENCIAMENTO DE PROGRAMAS DE COMPUTADOR CUSTOMIZAVEIS', 'descricaoServico': '(SERVICOS DE LICENCIAMENTO OU CESSAO DE DIREITO DE USO DE PROGRAMAS DE COMPUTACAO)  SERVICO CONTRATADO: LICENCIAMENTO HUBLY - COD. SERVICO 1.05 LC 116/03 - SERVICOS PRESTADOS EM FLORIANOPOLIS - INFORMACAO AO CONSUMIDOR CONFORME LEI 12741/12: 5,65%.\n :', 'idCNAE': '9102', 'quantidade': '1', 'valorTotal': '49.9', 'valorUnitario': '49.9'}]

pedidos = ["42000", "43000", "45000"]
contratos = ["47000", "48000"]
'''
for palavra in dados:
	if len(palavra) >= 10:
		# Verificar prefixos de pedidos
		for item in pedidos:
			if palavra.startswith(item):
				n_pedido = palavra[:10])  # Pegar apenas os primeiros 10 dígitos
'''

palavras = dados.replace('{', ' ').replace('}', ' ').replace(':', ' ').replace(';', ' ').replace('-', ' ').split()
print(palavras)

'''
def buscar_numeros_especificos(texto):
    if not isinstance(texto, str):
        return []
    
    # Padrão para buscar números que começam com 47 ou 45 seguidos de 8 dígitos
    padrao = r'(?:45)00\d{6}'
    
    # Encontrar todas as ocorrências
    resultados = re.findall(padrao, texto)
    
    return resultados

numeros_encontrados = []

for i, dicionario in enumerate(dados):
    print(f"\nPesquisando no item {i+1}:")
    
    # Iterar sobre cada par chave-valor
    for chave, valor in dicionario.items():
        resultados = buscar_numeros_especificos(str(valor))
        
        if resultados:
            print(f"  Na chave '{chave}' encontrado: {resultados}")
            numeros_encontrados.extend(resultados)

print("\nTodos os números encontrados:")
for numero in set(numeros_encontrados):  # Usando set para eliminar duplicatas
    print(f"  {numero}")

'''
'''
string[começo:fim:passo]

if "45000" in i:
				print ("ACHOU str = ", i)
			else:
				print("Não encontrado str")
			

substring = "Ren"
string = i
if string.find(substring) < 0:
    print("Não encontrou")
else:
    print("Encontrado na posição ", string.find(substring))

'''
'''	  
for pedido in pedidos:
	for ped in texto.split():
		if pedido in ped:
			num_pedido = ped[:10]
			print("Pedido: ", num_pedido)
			break

for contrato in contratos:
	for palavra in texto.split():
		if contrato in palavra:
			num_contrato = palavra[palavra.find(contrato):palavra.find(contrato)+10]
			print("Contrato: ", num_contrato)
			break
'''

