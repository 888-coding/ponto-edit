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
#01.03 Abrir nomes das planilhas
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


#B2 Inserir um valor de horas:minutos 
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
		print(f"{sh[campo].value}")
		print(Fore.GREEN + "Entrada :" + Fore.RESET)
		horas01 = input("Horas : ")
		minutos01 = input("Minutos : ")
		print(Fore.GREEN + "Almoço :" + Fore.RESET)
		horas02 = input("Horas : ")
		minutos02 = input("Minutos : ")
		print(Fore.GREEN + "Volta almoço :" + Fore.RESET)
		horas03 = input("Horas : ")
		minutos03 = input("Minutos : ")
		print(Fore.GREEN + "Saída :" + Fore.RESET)
		horas04 = input("Horas : ")
		minutos04 = input("Minutos : ")

		print(f"Voce digitou: Entrada: {horas01}:{minutos01}, Almoço: {horas02}:{minutos02}, Volta almoço: {horas03}:{minutos03}, Saída: {horas04}:{minutos04}")
		
else:
	print(Fore.RED + "Nome errado" + Fore.RESET)