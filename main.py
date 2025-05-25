# main.py
import os 
import openpyxl 
from datetime import time 
import colorama 
from colorama import Fore 

# 程式開始, 尋找 excel 檔案, 檔案位子
nome_arquivo = "file.xlsx"
diretorio_atual = os.path.dirname(os.path.realpath(__file__))

# 開啟檔案 , 開啟列表
wb = openpyxl.load_workbook(nome_arquivo)
todas_planilhas = wb.sheetnames
print(Fore.GREEN + "Bem vindo ao sistema" + Fore.RESET)

# 改日期
desejaAlterarData = input(Fore.YELLOW + "Deseja corrigir a o mes e ano ? (S/N)  : " + Fore.RESET).upper()
if desejaAlterarData == "S":
	novoMes = int(input("Favor digite o mês : "))
	novoAno = int(input("Favor digite o ano : "))
	sh = wb['DATA']
	sh['B1'].value = novoMes
	sh['B2'].value = novoAno	
	wb.save(nome_arquivo)
	print(Fore.GREEN + "Gravado com sucesso, abra de novo para alterar." + Fore.RESET)


#B2 插入時間(小時和分鐘) 
#sheet = arquivo['PEDRO']
#sheet['B2'].value = time(9,00,00)
#arquivo.save(nome_arquivo)

# 尋找名字 
sheetNome = input("Nome: ").upper()

nome_planilhas = wb.sheetnames

if sheetNome in nome_planilhas:
	# 找到了名字
	print(Fore.GREEN + "Encontrado" + Fore.RESET)
	sh = wb[sheetNome]
	for linha in range(2, 33):
		# 列表分為4行: B,C,D,E 
		# 分別為 上班時間, 吃飯時間, 吃飯回來時間, 下班時間
		# lista = ['B', 'C', 'D', 'E']
		campo = 'A' + str(linha)
		print(f"Dia : {sh[campo].value}")
		print("**Entrada")
		# 進來時間 (小時) 
		while True: 
			horas01 = input("Horas : ")
			if horas01.isnumeric():
				if int(horas01) >= 0 and int(horas01) <= 100:
					break
				else:
					print(Fore.RED + "Erro!" + Fore.RESET)
			else:
				print(Fore.RED + "Erro!" + Fore.RESET)
		# 進來時間 (分鐘) 
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
		# 吃飯時間 (小時)
		while True:
			horas02 = input("Horas : ")
			if horas02.isnumeric():
				if int(horas02) >= 0 and int(horas02) <= 100:
					break
				else:
					print(Fore.RED + "Erro!" + Fore.RESET)
			else:
				print(Fore.RED + "Erro!" + Fore.RESET)
		# 吃飯時間 (分鐘) 
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
		# 吃飯回來時間 (小時)
		while True:
			horas03 = input("Horas : ")
			if horas03.isnumeric():
				if int(horas03) >= 0 and int(horas03) <= 100:
					break
				else:
					print(Fore.RED + "Erro!" + Fore.RESET)
			else:
				print(Fore.RED + "Erro!" + Fore.RESET)
		# 吃飯回來時間 (分鐘) 
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
		# 下班時間 (小時) 
		while True:
			horas04 = input("Horas : ")
			if horas04.isnumeric():
				if int(horas04) >= 0 and int(horas04) <= 100:
					break
				else:
					print(Fore.RED + "Erro!" + Fore.RESET)
			else:
				print(Fore.RED + "Erro!" + Fore.RESET)
		# 下班時間 (分鐘)
		while True:
			minutos04 = input("Minutos : ")
			if minutos04.isnumeric():
				if int(minutos04) >= 0 and int(minutos04) <= 100:
					break
				else:
					print(Fore.RED + "Erro!" + Fore.RESET)
			else:
				print(Fore.RED + "Erro!" + Fore.RESET)

		# 顯示所寫的時間
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
	#沒有找到名字
	print(Fore.RED + "Nome errado" + Fore.RESET)