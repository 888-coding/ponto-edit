# main.py
import os 
import openpyxl 
from datetime import time 
import colorama 
from colorama import Fore 


#01: Procurar o arquivo 
nome_arquivo = "file.xlsx"
#01.01 Diretorio do programa 
diretorio_atual = os.path.dirname(os.path.realpath(__file__))

#01.02 Abrir arquivo 
wb = openpyxl.load_workbook(nome_arquivo)
#01.03 Abrir nomes das planilhas.
todas_planilhas = wb.sheetnames
#01.END Arquivo encontrado 
print(Fore.GREEN + "Bem vindo ao sistema" + Fore.RESET)
#02: Alterar data
desejaAlterarData = input(Fore.YELLOW + "Deseja corrigir a o mes e ano ? (S/N)  : " + Fore.RESET).upper()
if desejaAlterarData == "S":
	novoMes = int(input("Favor digite o mês : "))
	novoAno = int(input("Favor digite o ano : "))
	sh = wb['DATA']
	sh['B1'].value = novoMes
	sh['B2'].value = novoAno	
	wb.save(nome_arquivo)
	print(Fore.GREEN + "Gravado com sucesso, abra de novo para alterar." + Fore.RESET)


#B2 插入 時間 小時 和 分鐘 
#sheet = arquivo['PEDRO']
#sheet['B2'].value = time(9,00,00)
#arquivo.save(nome_arquivo)

#03: Procurar a planilha do nome 
sheetNome = input("Nome: ").upper()

nome_planilhas = wb.sheetnames

if sheetNome in nome_planilhas:
	print(Fore.GREEN + "Encontrado" + Fore.RESET)
	sh = wb[sheetNome]
	for linha in range(2, 33):
		#lista = ['B', 'C', 'D', 'E']
		campo = 'A' + str(linha)
		print(f"Dia : {sh[campo].value}")
		print("**Entrada")
		# Campo de Entrada : Horas 
		while True: 
			horas01 = input("Horas : ")
			if horas01.isnumeric():
				if int(horas01) >= 0 and int(horas01) <= 100:
					break
				else:
					print(Fore.RED + "Erro!" + Fore.RESET)
			else:
				print(Fore.RED + "Erro!" + Fore.RESET)
		# Campo de Entrada : Minutos 
		while True:
			minutos01 = input("Minutos : ")
			if minutos01.isnumeric():
				if int(minutos01) >= 0 and int(minutos01) <= 100:
					break
				else:
					print(Fore.RED + "Erro!" + Fore.RESET)
			else:
				print(Fore.RED + "Erro!" + Fore.RESET)
		print("**Almoço")
		# Campo de Almoço : Horas
		while True:
			horas02 = input("Horas : ")
			if horas02.isnumeric():
				if int(horas02) >= 0 and int(horas02) <= 100:
					break
				else:
					print(Fore.RED + "Erro!" + Fore.RESET)
			else:
				print(Fore.RED + "Erro!" + Fore.RESET)
		# Campo de Almoço : Minutos 
		while True:
			minutos02 = input("Minutos : ")
			if minutos02.isnumeric():
				if int(minutos02) >= 0 and int(minutos02) <= 100:
					break
				else:
					print(Fore.RED + "Erro!" + Fore.RESET)
			else:
				print(Fore.RED + "Erro!" + Fore.RESET)
		print("**Volta almoço")
		# Campo de Volta de almoço : Horas
		while True:
			horas03 = input("Horas : ")
			if horas03.isnumeric():
				if int(horas03) >= 0 and int(horas03) <= 100:
					break
				else:
					print(Fore.RED + "Erro!" + Fore.RESET)
			else:
				print(Fore.RED + "Erro!" + Fore.RESET)
		# Campo de Volta de almoço : Minutos 
		while True:
			minutos03 = input("Minutos : ")
			if minutos03.isnumeric():
				if int(minutos03) >= 0 and int(minutos03) <= 100:
					break
				else:
					print(Fore.RED + "Erro!" + Fore.RESET)
			else:
				print(Fore.RED + "Erro!" + Fore.RESET)
		print("**Saída")
		# Campo de Saída : Horas 
		while True:
			horas04 = input("Horas : ")
			if horas04.isnumeric():
				if int(horas04) >= 0 and int(horas04) <= 100:
					break
				else:
					print(Fore.RED + "Erro!" + Fore.RESET)
			else:
				print(Fore.RED + "Erro!" + Fore.RESET)
		# Campo de Saída : Minutos
		while True:
			minutos04 = input("Minutos : ")
			if minutos04.isnumeric():
				if int(minutos04) >= 0 and int(minutos04) <= 100:
					break
				else:
					print(Fore.RED + "Erro!" + Fore.RESET)
			else:
				print(Fore.RED + "Erro!" + Fore.RESET)

		print(Fore.GREEN + f"Voce digitou: Entrada: {horas01}:{minutos01}, Almoço: {horas02}:{minutos02}, Volta almoço: {horas03}:{minutos03}, Saída: {horas04}:{minutos04}" + Fore.RESET)
		if int(horas01) != 0:
			campo = 'B' + str(linha)
			sh[campo].value = time(int(horas01),int(minutos01),00)
			campo = 'C' + str(linha)
			sh[campo].value = time(int(horas02),int(minutos02),00)
			campo = 'D' + str(linha)
			sh[campo].value = time(int(horas03),int(minutos03),00)
			campo = 'E' + str(linha)
			sh[campo].value = time(int(horas04),int(minutos04),00)
			wb.save(nome_arquivo)
else:
	print(Fore.RED + "Nome errado" + Fore.RESET)