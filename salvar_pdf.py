'''1- abrir o documento
2- apertar f12
3- ir para pasta de destino
4- renomear o arquivo para o nome do aluno
5- mudar para o formato pdf
6- salvar'''

import pandas as pd
import pyautogui
import time

# ler a tabela Excel
df = pd.read_excel("Controle de Certificados de Ensino - PROEN.xlsx")
print(df.head())
lista_nomes: list = []
lista_cursos: list = []
status = [1, 2]


def checa_estado():
    if status[0] == 1:
        print(status[0])
        print('abrir documento')
        abrir_documento()
    if status[0] == 2:
        print(status[0])
        print('salvar pdf')
        salvar_pdf()


def abrir_documento():

    pyautogui.press('winleft')
    time.sleep(2)
    pyautogui.write('documento_test.docx', interval=0.15)
    time.sleep(2)
    pyautogui.press('enter')
    time.sleep(10)
    pyautogui.press('enter')
    time.sleep(2)
    status.pop(0)  # passa para salvar_pdf
    checa_estado()


def salvar_pdf():
    for nome in lista_nomes:
        for curso in lista_cursos:
            pyautogui.press('f12')
            time.sleep(2)
            pyautogui.write(f'{nome}.pdf')
            time.sleep(2)
            pyautogui.moveTo(507, 424)
            time.sleep(2)
            pyautogui.click(button='left')
            time.sleep(2)
            pyautogui.moveTo(430, 535)
            time.sleep(2)
            pyautogui.click(button='left')
            time.sleep(2)
            pyautogui.moveTo(590, 53)
            time.sleep(2)
            pyautogui.click(button='left')
            time.sleep(3)
            if curso == 'AGRONOMIA':
                pyautogui.write('disparar_email/Agronomia', interval=0.10)
                lista_cursos.pop(0)
                time.sleep(3)
                pyautogui.press('enter')
            elif curso == 'BIOLOGIA (LICENCIATURA)':
                pyautogui.write(
                    'disparar_email/Biologia(Licenciatura)', interval=0.10)
                lista_cursos.pop(0)
                time.sleep(3)
                pyautogui.press('enter')
            elif curso == 'CIÊNCIA E TECNOLOGIA DE ALIMENTOS':
                pyautogui.write(
                    'disparar_email/Ciencia e Tecnologia de Alimentos',
                    interval=0.10)
                lista_cursos.pop(0)
                time.sleep(3)
                pyautogui.press('enter')
            time.sleep(2)
            pyautogui.moveTo(794, 579)
            time.sleep(4)
            pyautogui.click(button='left')
            time.sleep(9)
            pyautogui.hotkey('alt', 'f4')
            lista_nomes.pop(0)
            time.sleep(2)
            mudar_para_proximo()


def mudar_para_proximo():
    pyautogui.moveTo(651, 37)
    time.sleep(1)
    pyautogui.click(button='left')
    time.sleep(1)
    pyautogui.moveTo(963, 64)
    time.sleep(1)
    pyautogui.click(button='left')
    salvar_pdf()


for index, row in df.iterrows():
    nome = row['Nome']
    curso = row['Curso']
    lista_nomes.append(str(nome))
    lista_cursos.append(str(curso))
    checa_estado()
