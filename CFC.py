#!/usr/bin python3
#!/usr/bin/env python3
# coding: utf-8
# Author: Helio Giroto
# Date: 17/02/2020

# ESTE SCRIPT É PARA FORMAR UMA PLANILHA DO MÊS REFERENTE AO PAGAMENTO DO C.F.C.

from openpyxl import Workbook
import datetime, os

def mostraAjuda():
	print()
	print('\t-----------------')
	print('\t Tipo de Serviço')
	print('\t-----------------')
	print('\t  0- Ver opções')
	print('\t  1- EXAME MÉDICO')
	print('\t  2- EXAME PSICOTÉCNICO')
	print('\t  3- EXAME MÉDICO/PSICOTÉCNICO')
	print('\t  4- CURSO CFC PH')
	print('\t  5- CURSO RECICLAGEM')
	print('\t  6- CURSO RENOVAÇÃO')
	print('\t  7- AULA SIMULADOR')
	print('\t  8- ADMISSIONAL')
	print('\t  9- DEMISSIONAL')
	print('\t-----------------')
	print()


def mudaNome(nro):
	nro = nro.replace('1', 'EXAME MÉDICO')
	nro = nro.replace('2', 'EXAME PSICOTÉCNICO')
	nro = nro.replace('3', 'EXAME MÉDICO/PSICOTÉCNICO')
	nro = nro.replace('4', 'CURSO CFC PH')
	nro = nro.replace('5', 'CURSO RECICLAGEM')
	nro = nro.replace('6', 'CURSO RENOVAÇÃO')
	nro = nro.replace('7', 'AULA SIMULADOR')
	nro = nro.replace('8', 'ADMISSIONAL')
	nro = nro.replace('9', 'DEMISSIONAL')
	return nro


listaDados = [['DATA', 'NOME', 'SERVIÇO', 'VALOR']]

def lancaDados():
	while True:
		data = 'FIM'
		print('Data do Serviço* : ', end='')
		try:		
			data = input().upper().strip(' ')
			data = data.replace('-','/')
			data = datetime.datetime.strptime(data, '%d/%m/%Y')
		except:
			if data == 'FIM':
				salvaPlanilha(listaDados)
				#data = "FIM"
				break
			else:
				print('\n***Erro na data. Repita digitação...***')
				# abaixo: por razões desconhecidas, a data sempre herda o valor que estava quando cumpria esta condição (else)...
				# Por isso, tive que passar FIM como valor para quebrar os laços qdo o usuário quiser finalizar.
				# qdo o usuário erra na data durante o lançamento este erro permanece como valor de data e ao finalizar, a variável data permanece com este valor do erro.
				data = 'FIM'
				lancaDados()

		#print(data)
		if data == 'FIM':
			break
		print('Nome do Aluno....: ', end='')
		nome = input().upper()
		nome = nome.strip(' ')

		print('Tipo do Serviço..: ', end='')
		servico = input()
		servico = servico.strip(' ')
		if servico == '0':
			mostraAjuda()
			print('Tipo do Serviço..: ', end='')
			servico = input()

		print('Valor cobrado R$ : ', end='')
		valor = input()
		valor = valor.replace(',', '.').replace(' ', '')
		valor = float(valor)

		nomeTipo = mudaNome(servico)

		# junta dados inseridos numa lista:
		dadosServico = [data, nome, nomeTipo, valor]
		print(dadosServico)

		print()
		print("Confirma ?")
		print("----------") 
		print("\t(S)IM \n\t(N)ÃO")
		confirma = input()
		#print()

		if confirma == "n" or confirma == "N":
			#data = "FIM"
			print('Registro cancelado...')
			print()		
			# o problema é qdo chama a mesma função.. (?)
			#lancaDados()
		else:
			print(data)
			listaDados.append(dadosServico)
			os.system('clear')
			# windows:
			# os.system('cls')
	print()



def salvaPlanilha(listaDados):
	wb = Workbook()
	ws = wb.active

	for item in listaDados:
		ws.append(item)

	qtdeServicos = len(listaDados)
	ws.auto_filter.ref = 'A1:D' + str(qtdeServicos)
	#ws.auto_filter.add_sort_condition('B2:B' + str(qtdeServicos))


	somaExMed  = '=SUMIF(C2:C'+str(qtdeServicos)+', "EXAME MÉDICO", D2:D'+str(qtdeServicos)+')'
	somaPsi    = '=SUMIF(C2:C'+str(qtdeServicos)+', "EXAME PSICOTÉCNICO", D2:D'+str(qtdeServicos)+')'
	somaMedPsi = '=SUMIF(C2:C'+str(qtdeServicos)+', "EXAME MÉDICO/PSICOTÉCNICO", D2:D'+str(qtdeServicos)+')'
	somaCfc    = '=SUMIF(C2:C'+str(qtdeServicos)+', "CURSO CFC PH", D2:D'+str(qtdeServicos)+')'
	somaRec    = '=SUMIF(C2:C'+str(qtdeServicos)+', "CURSO RECICLAGEM", D2:D'+str(qtdeServicos)+')'
	somaRen    = '=SUMIF(C2:C'+str(qtdeServicos)+', "CURSO RENOVAÇÃO", D2:D'+str(qtdeServicos)+')'
	somaSim    = '=SUMIF(C2:C'+str(qtdeServicos)+', "AULA SIMULADOR", D2:D'+str(qtdeServicos)+')'
	somaAdm    = '=SUMIF(C2:C'+str(qtdeServicos)+', "ADMISSIONAL", D2:D'+str(qtdeServicos)+')'
	somaDem    = '=SUMIF(C2:C'+str(qtdeServicos)+', "DEMISSIONAL", D2:D'+str(qtdeServicos)+')'
	celTotal   = '=SUM(D2:D'+str(qtdeServicos)+')'

	qtdeExMed  = '=COUNTIF(C2:C'+str(qtdeServicos)+', "EXAME MÉDICO")'
	qtdePsi    = '=COUNTIF(C2:C'+str(qtdeServicos)+', "EXAME PSICOTÉCNICO")'
	qtdeMedPsi = '=COUNTIF(C2:C'+str(qtdeServicos)+', "EXAME MÉDICO/PSICOTÉCNICO")'
	qtdeCfc    = '=COUNTIF(C2:C'+str(qtdeServicos)+', "CURSO CFC PH")'
	qtdeRec    = '=COUNTIF(C2:C'+str(qtdeServicos)+', "CURSO RECICLAGEM")'
	qtdeRen    = '=COUNTIF(C2:C'+str(qtdeServicos)+', "CURSO RENOVAÇÃO")'
	qtdeSim    = '=COUNTIF(C2:C'+str(qtdeServicos)+', "AULA SIMULADOR")'
	qtdeAdm    = '=COUNTIF(C2:C'+str(qtdeServicos)+', "ADMISSIONAL")'
	qtdeDem    = '=COUNTIF(C2:C'+str(qtdeServicos)+', "DEMISSIONAL")'
	qtdeTotal   = '=SUM(H2:H10)'

	# Poderia ser se não tivesse o recurso de sortear as colunas:
	# ws.cell(row=qtdeServicos+1, column=3).value = 'VALOR TOTAL: '
	# ws.cell(row=qtdeServicos+1, column=4).value = celTotal

	ws.cell(row=1, column=7) = 'R$'
	ws.cell(row=1, column=8) = 'Qtde'

	ws.cell(row=2, column=6).value = 'Exame Médico: '
	ws.cell(row=2, column=7).value = somaExMed
	ws.cell(row=2, column=8).value = qtdeExMed

	ws.cell(row=3, column=6).value = 'Exame Psicotécnico: '
	ws.cell(row=3, column=7).value = somaPsi
	ws.cell(row=3, column=8).value = qtdePsi

	ws.cell(row=4, column=6).value = 'Exame Médico/Psicotécnico: '
	ws.cell(row=4, column=7).value = somaMedPsi
	ws.cell(row=4, column=8).value = qtdeMedPsi

	ws.cell(row=5, column=6).value = 'Curso CFC PH: '
	ws.cell(row=5, column=7).value = somaCfc
	ws.cell(row=5, column=8).value = qtdeCfc

	ws.cell(row=6, column=6).value = 'Curso Reciclagem: '
	ws.cell(row=6, column=7).value = somaRec
	ws.cell(row=6, column=8).value = qtdeRec

	ws.cell(row=7, column=6).value = 'Curso Renovação: '
	ws.cell(row=7, column=7).value = somaRen
	ws.cell(row=7, column=8).value = qtdeRen

	ws.cell(row=8, column=6).value = 'Aula Simulador: '
	ws.cell(row=8, column=7).value = somaSim
	ws.cell(row=8, column=8).value = qtdeSim

	ws.cell(row=9, column=6).value = 'Admissional: '
	ws.cell(row=9, column=7).value = somaAdm
	ws.cell(row=9, column=8).value = qtdeAdm

	ws.cell(row=10, column=6).value = 'Demissional: '
	ws.cell(row=10, column=7).value = somaDem
	ws.cell(row=10, column=8).value = qtdeDem

	ws.cell(row=12, column=6).value = 'VALOR TOTAL: '
	ws.cell(row=12, column=7).value = celTotal
	ws.cell(row=12, column=8).value = qtdeTotal

	wb.save(planilha + '.xlsx')

	# PARA WINDOWS:
	# caminho = os.path.join("C:\\", "Users", "user", "Desktop", "python")
	# wb.save(caminho + '\\' + planilha + ".xlsx")


	print()
	print('PLANILHA SALVA - ' + planilha + '.xlsx\n')


print()	
print('*********************')
print('** CONTROLE DE CFC **')
print('*********************')
print()

print('Nome da Planilha: ')
planilha = input().lower()

print()
print('\t(- Digite FIM em "DATA" para terminar -)')
print('\t(- Digite 0 em "TIPO" para ver opções -)')
print()

# mostraAjuda()

lancaDados()

#print(listaDados)

# https://openpyxl.readthedocs.io/en/stable/filters.html
# https://openpyxl.readthedocs.io/en/stable/worksheet_tables.html

# https://banco.bradesco/assets/pessoajuridica/aplicativos/navegador-exclusivo/windows/Instalador.exe


