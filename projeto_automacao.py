import pyautogui as pyau
import time

#Bibliotecas
cnpj = {0:"12.345.678/0001-90",1:"98.765.432/0002-10",2:"45.678.901/0003-32",3:"76.543.210/0004-54",4:"11.234.567/0005-76",5:"89.123.450/0006-98"}

rs = {0:"Full Stack", 1:"Strategic Consulting", 2:"Innovation Solutions", 3:"Tech Vision", 4: "Smart Tech", 5:"Byte Works"}

servicos = {1:"Web App", 2:"Consultoria", 3:"Automacao"}

valor = {1:"R$ 400,00", 2:"R$ 300,00", 3:"R$ 850,00"}


print("0: Full Stack")
print("1: Strategic Consulting")
print("2: Innovation Solutions")
print("3: Tech Vision")
print("4: Smart Tech")
print("5: Byte Works")

emitente=int(input('Insira o código do emissor: '))
chave1=str(rs[emitente])
print(chave1)

chave2=str(cnpj[emitente])
print(chave2)

print("0: Full Stack")
print("1: Strategic Consulting")
print("2: Innovation Solutions")
print("3: Tech Vision")
print("4: Smart Tech")
print("5: Byte Works")

destinatario=int(input('Insira o código do destinatário: '))
chave3=str(rs[destinatario])
print(chave3)

chave4=str(cnpj[destinatario])
print(chave4)

print("1: Web App")
print("2: Consultoria")
print("3: Automacao")

servicos2=int(input('Insira o código do serviço: '))
chave5=str(servicos[servicos2])
print(chave5)

chave6=str(servicos[servicos2])
print(chave6)

chave7=str(valor[servicos2])
print(chave7)


def preencher():
    pyau.hotkey('win','d')
    pyau.click(492,240)
    pyau.hotkey('click')
    pyau.press('enter')
    time.sleep(2)
    pyau.hotkey('tab')
    pyau.hotkey('tab')
    pyau.write(chave1)
    pyau.hotkey('tab')
    pyau.write(chave2)
    pyau.hotkey('tab')
    pyau.write(chave3)
    pyau.hotkey('tab')
    pyau.write(chave4)
    pyau.hotkey('tab')
    pyau.write(str(servicos2))
    pyau.hotkey('tab')
    pyau.write(chave6)
    pyau.hotkey('tab')
    pyau.write(chave7)
    pyau.hotkey('tab')
    time.sleep(5)
    pyau.hotkey('enter')
    pyau.hotkey('ctrl','w')
    time.sleep(1)
preencher()

def excel():
    pyau.hotkey('win','r')
    pyau.write('Excel')
    pyau.press('enter')
    time.sleep(5)
    for a in range(3):
        pyau.hotkey('tab')
        time.sleep(1)
    pyau.hotkey('enter')
    time.sleep(1)
    pyau.write(chave1)
    pyau.hotkey('tab')
    time.sleep(1)
    pyau.write(chave2)
    pyau.hotkey('tab')
    time.sleep(1)
    pyau.write(chave3)
    pyau.hotkey('tab')
    time.sleep(1)
    pyau.write(chave4)
    pyau.hotkey('tab')
    time.sleep(1)
    pyau.write(str(servicos2))
    pyau.hotkey('tab')
    time.sleep(1)
    pyau.write(chave6)
    pyau.hotkey('tab')
    time.sleep(1)
    pyau.write(chave7)
    pyau.hotkey('tab')
    pyau.hotkey('tab')
    pyau.hotkey('ctrl','b')
excel()