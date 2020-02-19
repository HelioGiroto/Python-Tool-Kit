#!/usr/bin python3
#!/usr/bin/env python3
# coding: utf-8
# Author: Helio Giroto
# Date: 15/02/2020

# ESTE SCRIPT É PARA FORMAR UMA PLANILHA DO MÊS REFERENTE AO ABASTECIMENTO DE VEÍCULOS DE UMA EMPRESA:


print()	
print('*****************************')
print('* ABASTECIMENTO DE VEÍCULOS *')
print('*****************************')
print()
# pede nome do mes
print('DIGITE O MÊS: ')
nome_mes = input()
nome_mes = nome_mes.upper()

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
# print(mes)


# enquanto resposta for diferente de FIM:
placa_carro = ''
while(placa_carro != "FIM"):
	# pede nome do carro
	print()
	print()
	print("*****************************")
	print('PLACA DO CARRO (Ao terminar digite: FIM): ', end='')
	placa_carro = input()
	# passa valor do input para tudo maiúsculo
	placa_carro = placa_carro.upper().replace(' ','')
	print()

	if placa_carro == "FIM":
		break

	# cria nova lista com 32 posições e com nome do carro no index [0]
	carro = [''] *32
	carro[0] = placa_carro

	dia = ''
	# enquanto resposta for diferente de FIM:
	while(dia != "FIM"):
		# pede dias que almoçou.
		print("DIA (ou: 'FIM'): ", end='')
		dia = input()
		dia = dia.upper().replace(' ', '')
		# passa para a lista carro o preço dentro do dia: carro[dia] = preço
		if dia == "FIM":
			break
		# pede preço do almoço (12.00)
		print('VALOR R$       : ', end='')
		valor = input()
		valor = valor.replace(',', '.').replace(' ', '')
		print()
		carro[int(dia)] = float(valor)

	# appenda lista carro dentro da lista mes: mes.append(carro)
	mes.append(carro)


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
wb.save(nome_mes + "-abastecimento.xlsx") 

