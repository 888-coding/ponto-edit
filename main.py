# main.py
import os 
import openpyxl 
import colorama 
from colorama import Fore 

#01: Procurar o arquivo 
nome_arquivo = "file.xlsx"
#01.01 Diretorio do programa 
diretorio_atual = os.path.dirname(os.path.realpath(__file__))

#01.02 Abrir arquivo 
arquivo = openpyxl.load_workbook(nome_arquivo)
#01.03 Abrir nomes das planilhas
todas_planilhas = arquivo.sheetnames
#01.END Arquivo encontrado 
print(Fore.GREEN + "Bem vindo ao sistema" + Fore.RESET)
#02: Alterar data
desejaAlterarData = input(Fore.YELLOW + "Deseja corrigir a o mes e ano ? (S/N)  : " + Fore.RESET).upper()
if desejaAlterarData == "S":
	novoMes = int(input("Favor digite o mês : "))
	novoAno = int(input("Favor digite o ano : "))
	sheet = arquivo['DATA']
	sheet['B1'].value = novoMes
	sheet['B2'].value = novoAno	
	arquivo.save(nome_arquivo)
	print(Fore.GREEN + "Gravado com sucesso, abra de novo para alterar." + Fore.RESET)
