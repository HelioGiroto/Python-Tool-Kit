#!/usr/bin python3
#!/usr/bin/env python3
# coding: utf-8
# Author: Helio Giroto
# Date: 24/02/2020

# ESTE SCRIPT ABRE PLANILHA E APPENDA DADOS. 
# Adiciona avaliações dos clientes.

# from openpyxl import Workbook 
import openpyxl as op
import datetime, os


def processa(novas_notas):

	# abre arquivo (w-book) de planilha:
	wb = op.load_workbook('avaliacao.xlsx')
	# Abre planilha (w-sheet) - aba: (podendo ser escolha do usuário)
	ws = wb.get_sheet_by_name('AUTOESCOLA')


	# first_row = list(ws.rows)[0][0] # pega dados de coluna infinita....
	# print(first_row)


	# imprime o nro da última linha (do rodapé):
	print(ws.max_row)
	nro_ult_linha = ws.max_row
	

	# Se tivesse que apenas cortar a ultima linha, appendar outras e ao final colar a ult.linha, bastaria isso:

	# conteúdo da ult. linha (Tem que iterar): 
	# rodape = []
	#for linha in ws.iter_rows(min_row=nro_ult_linha, max_row=nro_ult_linha, min_col=1, max_col=38):
	#	rodape = linha


	# apaga a última linha: (linha antiga com totais de porcentagem)
	ws.delete_rows(nro_ult_linha)


	'''
	ult_linha = ws[ws.max_row]
	for linha2 in list(ws.rows)[1]:
		print(linha2.value)
	'''


	# define novas notas (apenas simulação):
	# aqui deve vir de inputs....:
	# novas_notas = [2,10,1,8,2,10,2,7,1,8,0,10,1,7,2,8,2,10,0,8,1,6,2,8,0,5,2,10,1,5,2,9,1007,'01/01/2019',0,10,50,77]

	# Appenda (na última linha) as novas notas:
	ws.append(novas_notas)

	# perc_carinha = '=...(AK3:AK'+str(nro_ult_linha)+')'


	# sort - ordem:
	ws.auto_filter.ref = 'A2:AL' + str(nro_ult_linha)


	# O nro da ultima linha aumentou conforme foi appendando, por isso, se requer atualizar variável:
	nro_ult_linha = ws.max_row


	# =(SOMA(A3:A34))/(LINHAS(A3:A34)*2)*100
	tot_A = '=(SUM(A3:A'+str(nro_ult_linha)+'))/(ROWS(A3:A'+str(nro_ult_linha)+')*2)*100'
	tot_B = '=(SUM(B3:B'+str(nro_ult_linha)+'))/(ROWS(B3:B'+str(nro_ult_linha)+')*10)*100'
	tot_C = '=(SUM(C3:C'+str(nro_ult_linha)+'))/(ROWS(C3:C'+str(nro_ult_linha)+')*2)*100'
	tot_D = '=(SUM(D3:D'+str(nro_ult_linha)+'))/(ROWS(D3:D'+str(nro_ult_linha)+')*10)*100'
	tot_E = '=(SUM(E3:E'+str(nro_ult_linha)+'))/(ROWS(E3:E'+str(nro_ult_linha)+')*2)*100'
	tot_F = '=(SUM(F3:F'+str(nro_ult_linha)+'))/(ROWS(F3:F'+str(nro_ult_linha)+')*10)*100'
	tot_G = '=(SUM(G3:G'+str(nro_ult_linha)+'))/(ROWS(G3:G'+str(nro_ult_linha)+')*2)*100'
	tot_H = '=(SUM(H3:H'+str(nro_ult_linha)+'))/(ROWS(H3:H'+str(nro_ult_linha)+')*10)*100'
	tot_I = '=(SUM(I3:I'+str(nro_ult_linha)+'))/(ROWS(I3:I'+str(nro_ult_linha)+')*2)*100'
	tot_J = '=(SUM(J3:J'+str(nro_ult_linha)+'))/(ROWS(J3:J'+str(nro_ult_linha)+')*10)*100'
	tot_K = '=(SUM(K3:K'+str(nro_ult_linha)+'))/(ROWS(K3:K'+str(nro_ult_linha)+')*2)*100'
	tot_L = '=(SUM(L3:L'+str(nro_ult_linha)+'))/(ROWS(L3:L'+str(nro_ult_linha)+')*10)*100'
	tot_M = '=(SUM(M3:M'+str(nro_ult_linha)+'))/(ROWS(M3:M'+str(nro_ult_linha)+')*2)*100'
	tot_N = '=(SUM(N3:N'+str(nro_ult_linha)+'))/(ROWS(N3:N'+str(nro_ult_linha)+')*10)*100'
	tot_O = '=(SUM(O3:O'+str(nro_ult_linha)+'))/(ROWS(O3:O'+str(nro_ult_linha)+')*2)*100'
	tot_P = '=(SUM(P3:P'+str(nro_ult_linha)+'))/(ROWS(P3:P'+str(nro_ult_linha)+')*10)*100'
	tot_Q = '=(SUM(Q3:Q'+str(nro_ult_linha)+'))/(ROWS(Q3:Q'+str(nro_ult_linha)+')*2)*100'
	tot_R = '=(SUM(R3:R'+str(nro_ult_linha)+'))/(ROWS(R3:R'+str(nro_ult_linha)+')*10)*100'
	tot_S = '=(SUM(S3:S'+str(nro_ult_linha)+'))/(ROWS(S3:S'+str(nro_ult_linha)+')*2)*100'
	tot_T = '=(SUM(T3:T'+str(nro_ult_linha)+'))/(ROWS(T3:T'+str(nro_ult_linha)+')*10)*100'
	tot_U = '=(SUM(U3:U'+str(nro_ult_linha)+'))/(ROWS(U3:U'+str(nro_ult_linha)+')*2)*100'
	tot_V = '=(SUM(V3:V'+str(nro_ult_linha)+'))/(ROWS(V3:V'+str(nro_ult_linha)+')*10)*100'
	tot_W = '=(SUM(W3:W'+str(nro_ult_linha)+'))/(ROWS(W3:W'+str(nro_ult_linha)+')*2)*100'
	tot_X = '=(SUM(X3:X'+str(nro_ult_linha)+'))/(ROWS(X3:X'+str(nro_ult_linha)+')*10)*100'
	tot_Y = '=(SUM(Y3:Y'+str(nro_ult_linha)+'))/(ROWS(Y3:Y'+str(nro_ult_linha)+')*2)*100'
	tot_Z = '=(SUM(Z3:Z'+str(nro_ult_linha)+'))/(ROWS(Z3:Z'+str(nro_ult_linha)+')*10)*100'
	tot_AA = '=(SUM(AA3:AA'+str(nro_ult_linha)+'))/(ROWS(AA3:AA'+str(nro_ult_linha)+')*2)*100'
	tot_AB = '=(SUM(AB3:AB'+str(nro_ult_linha)+'))/(ROWS(AB3:AB'+str(nro_ult_linha)+')*10)*100'
	tot_AC = '=(SUM(AC3:AC'+str(nro_ult_linha)+'))/(ROWS(AC3:AC'+str(nro_ult_linha)+')*2)*100'
	tot_AD = '=(SUM(AD3:AD'+str(nro_ult_linha)+'))/(ROWS(AD3:AD'+str(nro_ult_linha)+')*10)*100'
	tot_AE = '=(SUM(AE3:AE'+str(nro_ult_linha)+'))/(ROWS(AE3:AE'+str(nro_ult_linha)+')*2)*100'
	tot_AF = '=(SUM(AF3:AF'+str(nro_ult_linha)+'))/(ROWS(AF3:AF'+str(nro_ult_linha)+')*10)*100'
	
	tot_AK = '=AVERAGE(AK3:AK'+str(nro_ult_linha)+')'
	tot_AL = '=AVERAGE(AL3:AL'+str(nro_ult_linha)+')'


	# forma a linha de rodapé (uma lista) com totais:
	rodape = [tot_A, tot_B, tot_C, tot_D, tot_E, tot_F, tot_G, tot_H, tot_I, tot_J, tot_K, tot_L, tot_M, tot_N, tot_O, tot_P, tot_Q, tot_R, tot_S, tot_T, tot_U, tot_V, tot_W, tot_X, tot_Y, tot_Z, tot_AA, tot_AB, tot_AC, tot_AD, tot_AE, tot_AF]


	# appenda na planilha o rodapé com os totais:
	ws.append(rodape)


	# depois de appendar, é necessário colar na mesma linha de rodapé as médias das avaliações:
	ws.cell(row=nro_ult_linha+1, column=37).value = tot_AK
	ws.cell(row=nro_ult_linha+1, column=38).value = tot_AL


	# para insertar uma linha em branco
	#ws.insert_rows(nro_ult_linha+1)


	# Salva planilha:
	wb.save('avaliacao.xlsx')

	print('\n**** Dados salvos na planilha. ****\n')


def lancaDadosDespachante():
	print()

def lancaDadosAutoescola():
	while True:
		print('Data: ', end='')
		try:		
			data = input()
			data = data.replace('-','/')
			data = datetime.datetime.strptime(data, '%d/%m/%Y')
		except:
			print('Data inválida, Favor redigitar...\n')
			lancaDadosAutoescola()
				
		print()
		print('Nro da Pesquisa: ', end='')
		nro_pesquisa = input()
		print()

		print('**********************')
		print('**     CARINHAS     **')
		print('**      0 = :(      **')
		print('**      1 = :|      **')
		print('**      2 = :)      **')
		print('**********************')
		print()

		print('Limpeza        : ', end='')		
		limpezaC = int(input())

		print('Infraestrutura : ', end='')
		infraestruturaC = int(input())

		print('Frota          : ', end='')
		frotaC = int(input())

		print('Informações    : ', end='')
		informacoesC = int(input())

		print('Horário        : ', end='')
		horarioC = int(input())

		print('Preço          : ', end='')
		precoC = int(input())

		print('Condições      : ', end='')
		condicoesC = int(input())

		print('Atendimento    : ', end='')
		atendimentoC = int(input())

		print('Orientações    : ', end='')
		orientacoesC = int(input())

		print('Agilidade      : ', end='')
		agilidadeC = int(input())

		print('Qualidade      : ', end='')
		qualidadeC = int(input())

		print('Prazo          : ', end='')
		prazoC = int(input())

		print('Aulas Práticas : ', end='')
		aulasC = int(input())

		print('Sorteio        : ', end='')
		sorteioC = int(input())

		print('Avaliação      : ', end='')
		avaliacaoC = int(input())

		print('Recomenda      : ', end='')
		recomendaC = int(input())

		print()

		print('***********************')
		print('**       NOTAS       **')
		print('**                   **')
		print('**    (de 0 a 10)    **')
		print('**                   **')
		print('***********************')
		print()

		print('Limpeza        : ', end='')
		limpeza = int(input())

		print('Infraestrutura : ', end='')
		infraestrutura = int(input())

		print('Frota          : ', end='')
		frota = int(input())

		print('Informações    : ', end='')
		informacoes = int(input())

		print('Horário        : ', end='')
		horario = int(input())

		print('Preço          : ', end='')
		preco = int(input())

		print('Condições      : ', end='')
		condicoes = int(input())

		print('Atendimento    : ', end='')
		atendimento = int(input())

		print('Orientações    : ', end='')
		orientacoes = int(input())

		print('Agilidade      : ', end='')
		agilidade = int(input())

		print('Qualidade      : ', end='')
		qualidade = int(input())

		print('Prazo          : ', end='')
		prazo = int(input())

		print('Aulas Práticas : ', end='')
		aulas = int(input())

		print('Sorteio        : ', end='')
		sorteio = int(input())

		print('Avaliação      : ', end='')
		avaliacao = int(input())

		print('Recomenda      : ', end='')
		recomenda = int(input())

		print()

		porcentagem_carinhas = (((limpezaC + infraestruturaC + frotaC + informacoesC + horarioC + precoC + condicoesC + atendimentoC + orientacoesC + agilidadeC + qualidadeC + prazoC + aulasC + sorteioC + avaliacaoC + recomendaC)/32) *100)

		porcentagem_notas    = (((limpeza + infraestrutura + frota + informacoes + horario + preco + condicoes + atendimento + orientacoes + agilidade + qualidade + prazo + aulas + sorteio + avaliacao + recomenda)/160) *100)

		novas_notas = [limpezaC, limpeza, infraestruturaC, infraestrutura, frotaC, frota, informacoesC, informacoes, horarioC, horario, precoC, preco, condicoesC, condicoes, atendimentoC, atendimento, orientacoesC, orientacoes, agilidadeC, agilidade, qualidadeC, qualidade, prazoC, prazo, aulasC, aulas, sorteioC, sorteio, avaliacaoC, avaliacao, recomendaC, recomenda, nro_pesquisa, data, '-', '-', porcentagem_carinhas, porcentagem_notas]

		# novas_notas = [limpezaC, limpeza, infraestruturaC, infraestrutura, frotaC, frota, informacoesC, informacoes, horarioC, horario, precoC, preco, condicoesC, condicoes, atendimentoC, atendimento, orientacoesC, orientacoes, agilidadeC, agilidade, qualidadeC, qualidade, prazoC, prazo, aulasC, aulas, sorteioC, sorteio, avaliacaoC, avaliacao, recomendaC, recomenda, nro_pesquisa, data]

		processa(novas_notas)



# só executa a função acima, se existir o arquivo avaliacao.xlxs:
if os.path.exists('avaliacao.xlsx'):
	print('Planilha existe.')
	print()
	lancaDadosAutoescola()
else:
	print('Arquivos nao existe.')
	# criar planilha em branco





'''


listaDados = [['DATA', 'NOME', 'SERVIÇO', 'VALOR']]

def lancaDados():
	while True:
		data = 'FIM'
		print('Data do Serviço* : ', end='')
		try:		
			data = input().replace('','0').upper().strip(' ')
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
		nome = input().replace('','0').upper()
		nome = nome.strip(' ')

		print('Tipo do Serviço..: ', end='')
		servico = input().replace('','0')
		servico = servico.strip(' ')
		if servico == '0':
			mostraAjuda()
			print('Tipo do Serviço..: ', end='')
			servico = input().replace('','0')

		print('Valor cobrado R$ : ', end='')
		valor = input().replace('','0')
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
		confirma = input().replace('','0')
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

	# Poderia ser se não tivesse o recurso de sortear as colunas:
	# ws.cell(row=qtdeServicos+1, column=3).value = 'VALOR TOTAL: '
	# ws.cell(row=qtdeServicos+1, column=4).value = celTotal


	ws.cell(row=2, column=6).value = 'Exame Médico: '
	ws.cell(row=2, column=7).value = somaExMed

	ws.cell(row=3, column=6).value = 'Exame Psicotécnico: '
	ws.cell(row=3, column=7).value = somaPsi

	ws.cell(row=4, column=6).value = 'Exame Médico/Psicotécnico: '
	ws.cell(row=4, column=7).value = somaMedPsi

	ws.cell(row=5, column=6).value = 'Curso CFC PH: '
	ws.cell(row=5, column=7).value = somaCfc

	ws.cell(row=6, column=6).value = 'Curso Reciclagem: '
	ws.cell(row=6, column=7).value = somaRec

	ws.cell(row=7, column=6).value = 'Curso Renovação: '
	ws.cell(row=7, column=7).value = somaRen

	ws.cell(row=8, column=6).value = 'Aula Simulador: '
	ws.cell(row=8, column=7).value = somaSim

	ws.cell(row=9, column=6).value = 'Admissional: '
	ws.cell(row=9, column=7).value = somaAdm

	ws.cell(row=10, column=6).value = 'Demissional: '
	ws.cell(row=10, column=7).value = somaDem

	ws.cell(row=12, column=6).value = 'VALOR TOTAL: '
	ws.cell(row=12, column=7).value = celTotal

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
planilha = input().replace('','0').lower()

print()
print('\t(- Digite FIM em "DATA" para terminar -)')
print('\t(- Digite 0 em "TIPO" para ver opções -)')
print()

# mostraAjuda()

lancaDados()

#print(listaDados)

# https://openpyxl.readthedocs.io/en/stable/filters.html
# https://openpyxl.readthedocs.io/en/stable/worksheet_tables.html

'''
