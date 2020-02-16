#!/usr/bin python3
#!/usr/bin/env python3
# coding: utf-8
# Author: Helio Giroto

# ESTE SCRIPT É PARA FORMAR UMA PLANILHA DO MÊS REFERENTE AO VALE REFEIÇÃO DOS FUNCIONÁRIOS DE UMA EMPRESA:

# pede nome do mes
print('DIGITE O NOME DO MÊS: ')
nome_mes = input()
nome_mes = nome_mes.upper()
print()

# pede preço do almoço (12.00)
print('QUAL O PREÇO DO ALMOÇO: ')
preco = input()

# cria uma lista para mes
mes = []

# cria o cabeçalho (referente a mes[0]) com nros de 0 a 31
cabecalho = []
for linhas in range(32):
	cabecalho.append(str(linhas))


# primeiro elemento [0] desta lista é o nome do mês
cabecalho[0] = nome_mes

# appenda cabeçalho à lista mês
mes.append(cabecalho)
#print(mes)


# enquanto resposta for diferente de FIM:
nome_funcionario = ''
while(nome_funcionario != "FIM"):
	print()
	# pede nome do funcionario
	print("********************************")
	print('NOME DO FUNCIONÁRIO:')
	print("(Se quer encerrar, digite: FIM)")
	nome_funcionario = input()
	# passa valor do input para tudo maiúsculo
	nome_funcionario = nome_funcionario.upper()
	print()
	if nome_funcionario == "FIM":
		break

	# cria nova lista com 32 posições e com nome do funcionario no index [0]
	funcionario = [''] *32
	funcionario[0] = nome_funcionario

	# pede dias que almoçou. Não precisa os dias estarem na ordem numérica.
	print("DIGITE OS DIAS QUE ELE ALMOÇOU (separados por espaço):")
	dados = input()
	listaDados = dados.split(" ")
	# passa para a lista funcionario o preço dentro do dia: funcionario[dia] = preço
	for item in listaDados:
		funcionario[ int(item) ] = int(preco)

	# appenda lista funcionario dentro da lista mes: mes.append(funcionario)
	mes.append(funcionario)

print(mes)


########################################################################

# Prepara planilha Excel:

# Importa módulo
import openpyxl
# Abre nova aba: 
wb = openpyxl.Workbook() 
# Deixa essa aba aberta para uso:
sheet = wb.active 


# qtde de colunas se refere ao nro de registros que o mes terá:
qtdeColunas = len(mes)

# Looping para pegar valores das listas e colocá-los na planilha.
# A diferença entre valores em array e em planilha é que se inverte linha para coluna e vice-versa.
for linha in range(32):
	for coluna in range(qtdeColunas):
		# O endereço da celula tem que ser +1 porque na planilha não há a célula/linha 0 !!
		# atençao no valor que recebe a celula! (mes[][] - não pode ser linha, coluna pq valores das listas e planilhas são invertidos !!)
		sheet.cell(row=linha+1, column=coluna+1).value = mes[coluna][linha]


# Células para calcular totais:
sheet['B33'] = "=SUM(B2:B32)"
sheet['C33'] = "=SUM(C2:C32)"
sheet['D33'] = "=SUM(D2:D32)"
sheet['E33'] = "=SUM(E2:E32)"
sheet['F33'] = "=SUM(F2:F32)"
sheet['G33'] = "=SUM(G2:G32)"
sheet['H33'] = "=SUM(H2:H32)"

sheet['I32'] = "TOTAL GERAL:"
sheet['I33'] = "=SUM(B33:H33)"

# salva o arquivo:
wb.save(nome_mes + "-valeRefeicao.xlsx") 

