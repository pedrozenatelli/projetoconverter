import glob
import json
import os
import re
import sys

import numpy as np
import pandas as pd
import requests
from PyQt5 import QtWidgets, uic
from PyQt5.QtCore import *


def replace_images(text, images):
    # Cria uma expressão regular para encontrar todas as ocorrências de {IMAGEM\d}
    pattern = re.compile(r'\{IMAGEM\d\}')

    # Substitui todas as ocorrências pelo valor correspondente da lista de imagens
    for image in images:
        text = pattern.sub(f'<img src="/image/{image}">', text, count=1)

    return text


def download_image(url, name):
    # Verifica se há uma url de uma imagem e salva dentro da variavel url
    if url:
        # Baixa a imagem e salva na pasta images
        image = requests.get(url, stream=True)
        # Salva a imagem na pasta images
        with open(os.path.join(os.getcwd(), 'images', name), 'wb') as f:
            f.write(image.content)
        return name
    else:
        return ''


def cria_json(nome):
    # Define o caminho do arquivo
    caminho_arquivo = os.path.join(os.getcwd(), 'media', f'{nome}.json')

    # Dados a serem escritos
    # Adicione aqui os dados que deseja escrever no arquivo JSON
    data = []

    # Verifica se não existe arquivo com mesmo nome na pasta informada
    if os.path.exists(caminho_arquivo):
        print(f'Arquivo {nome}.json já existe')

        # Se o arquivo já existe apaga o conteudo do arquivo deixando apenas []
        with open(caminho_arquivo, 'w') as file:
            json.dump(data, file)
        return f'{nome}.json'
    else:
        # Cria o arquivo JSON
        with open(caminho_arquivo, 'w') as file:
            json.dump(data, file)

        return f'{nome}.json'


def salvar_json(value, name_arq):
    # Define o caminho do arquivo
    caminho_arquivo = os.path.join(os.getcwd(), 'saida', f'{name_arq}')

    # Abre o arquivo JSON
    with open(caminho_arquivo, 'r') as file:
        data = json.load(file)

    # Adiciona o novo valor
    data.append(value)

    # Salva o arquivo JSON
    with open(caminho_arquivo, 'w', encoding='utf-8') as file:
        json.dump(data, file, indent=4, ensure_ascii=False)


class MyFirstApp(QtWidgets.QMainWindow):

    def __init__(self):
        QtWidgets.QMainWindow.__init__(self)
        self.ui = uic.loadUi("ProjetoJson.ui", self)
        self.setFixedSize(577, 571)  # Definir um tamanho fixo para a janela
        self.btnCriaJson.clicked.connect(self.runjson)
        self.btnConverterNew.clicked.connect(self.converterNew)
        self.btnCleanIntro.clicked.connect(self.clearintro)
        self.btnCleanExit.clicked.connect(self.clearexit)
        self.add_inlist()

    def add_inlist(self):
        way_xlsx = [itt for itt in os.listdir(
            os.path.join(os.getcwd(), 'entrada')) if itt.endswith('.xlsx')]
        way_json = [itt for itt in os.listdir(
            os.path.join(os.getcwd(), 'saida')) if itt.endswith('.json')]
        print(way_xlsx)
        print(way_json)
        self.listWidgetXlsx.clear()
        self.listWidgetJson.clear()
        for icone in way_xlsx:
            self.listWidgetXlsx.addItem(icone)
        for icone in way_json:
            self.listWidgetJson.addItem(icone)

    def evt_update_progress(self, val):
        self.progressBar.setValue(val)

    def evt_update_label(self, val):
        self.labelSaida.setText(val)

    def converterNew(self):
        self.labelSaida.setText("")
        self.progressBar.setValue(0)
        self.add_inlist()

    def runjson(self):
        self.sync_ = ConvJsonThread()
        self.sync_.start()
        self.sync_.update_progress_s.connect(self.evt_update_progress)
        self.sync_.update_rotulo_s.connect(self.evt_update_label)
        self.add_inlist()

    def clearintro(self):
        folder_path = os.path.join(os.getcwd(), 'entrada')
        # Use glob to find all .xlsx files in the folder
        files = glob.glob(os.path.join(folder_path, '*.xlsx'))

        # Loop through the list of filepaths & remove each file.
        for file in files:
            try:
                os.remove(file)
                print(f"File {file} has been removed successfully")
            except Exception as e:
                print(
                    f"Error occurred while deleting file : {file}. Reason: {e}")
        self.add_inlist()

    def clearexit(self):
        folder_path = os.path.join(os.getcwd(), 'saida')
        # Use glob to find all .xlsx files in the folder
        files = glob.glob(os.path.join(folder_path, '*.json'))

        # Loop through the list of filepaths & remove each file.
        for file in files:
            try:
                os.remove(file)
                print(f"File {file} has been removed successfully")
            except Exception as e:
                print(
                    f"Error occurred while deleting file : {file}. Reason: {e}")
        self.add_inlist()


class ConvJsonThread(QThread):

    update_progress_s = pyqtSignal(int)
    update_rotulo_s = pyqtSignal(str)

    def __init__(self):
        super(QThread, self).__init__()

    def run(self):
        self.update_rotulo_s.emit("Iniciando conversão. Aguarde...")

        ccouter = 1

        for itt in os.listdir(os.path.join(os.getcwd(), 'entrada')):
            if itt.endswith('xlsx'):
                # Nome a ser colocado no json
                name_ext = itt.replace('.xlsx', '')
                # caminho completo para o arquivo
                caminho = os.path.join(os.getcwd(), 'entrada', itt)

                # Lendo o arquivo
                df = pd.read_excel(caminho)

                # Cria o arquivo json
                name_arq = cria_json(name_ext)
                print(f"name_arq: {name_arq}")

                # Pega a primeira linha
                primeira_linha = df.iloc[0]

                # Cria a lista com os dados da primeira linha
                lista_t = [item for item in primeira_linha]

                # # Lista nova vazia
                nova_lista_t = []

                # Loop pela lista original
                for item in lista_t:
                    # Verifica se o item é 'A'

                    if item in ['Letra A', 'Letra B', 'Letra C', 'Letra D', 'Letra E']:
                        nova_lista_t.append(item.replace('Letra ', ''))
                    # Verifica se o item é 'B', 'C', 'D' ou 'E'
                    elif item not in ['A', 'B', 'C', 'D', 'E', 'Prova1', 'Prova2', 'Prova3', 'Prova4', 'Prova5', 'Prova6', 'Prova7', 'Prova8', 'Prova9', 'Prova10', 'Prova11', 'Prova12', 'Prova13', 'Prova14', 'Prova15', 'Assunto1', 'Assunto2', 'Assunto3', 'Assunto4', 'Assunto5', 'Assunto6', 'Assunto7', 'Assunto8', 'Assunto9', 'Assunto10', 'Assunto11', 'Assunto12', 'Assunto13', 'Assunto14', 'Assunto15', 'IMAGEM1', 'IMAGEM2', 'IMAGEM3', 'IMAGEM4', 'IMAGEMALTA', 'IMAGEMALTB', 'IMAGEMALTC', 'IMAGEMALTD', 'IMAGEMALTE']:
                        nova_lista_t.append(item)

                nova_lista_t.append('Provas')
                nova_lista_t.append('Assuntos')
                nova_lista_t.append('Imagens')

                # # Atribui a nova lista à lista_t
                lista_t = nova_lista_t

                # Imprime a lista
                print(lista_t)
                print(
                    "*****************************************************************")
                # print(len(lista_t))
                itenslen = len(list(df.itertuples()))
                # Iterando sobre as linhas do DataFrame
                for row in df.itertuples():
                    print(f"row: {row}")
                    # print(f"row: {row[2]}")
                    collector = {}
                    if row.Index == 0:
                        continue
                    else:
                        # Codigo
                        collector[lista_t[0]] = row[1]
                        # Pergunta
                        collector[lista_t[1]] = row[2].replace(
                            '\n', ' ').replace('\xa0', ' ')
                        # Comentario do professor
                        collector[lista_t[2]] = row[3]
                        # Tipo
                        collector[lista_t[3]] = row[4]
                        # Materia
                        collector[lista_t[4]] = row[5]
                        # Disciplica
                        collector[lista_t[5]] = row[6]
                        # Banca
                        collector[lista_t[6]] = row[7]
                        # Órgão
                        collector[lista_t[7]] = row[8]
                        # Ano
                        collector[lista_t[8]] = row[9]
                        # Alternativa A
                        collector[lista_t[9]] = row[15] if not pd.isnull(
                            row[15]) else ''
                        # Alternativa B
                        collector[lista_t[10]] = row[16] if not pd.isnull(
                            row[16]) else ''
                        # Alternativa C
                        collector[lista_t[11]] = row[17] if not pd.isnull(
                            row[17]) else ''
                        # Alternativa D
                        collector[lista_t[12]] = row[18] if not pd.isnull(
                            row[18]) else ''
                        # Alternativa E
                        collector[lista_t[13]] = row[19] if not pd.isnull(
                            row[19]) else ''
                        # Alternativa Correta
                        collector[lista_t[14]] = row[20] if not pd.isnull(
                            row[20]) else ''
                        # Melhor Comentário
                        collector[lista_t[15]] = row[21] if not pd.isnull(
                            row[21]) else ''
                        # Provas
                        provas_collect = []
                        for i in range(22, 37):
                            if not pd.isnull(row[i]):
                                provas_collect.append(row[i])
                        collector[lista_t[16]] = provas_collect
                        # Assuntos
                        assuntos_collect = []
                        for i in range(37, 52):
                            if not pd.isnull(row[i]):
                                assuntos_collect.append(row[i])
                        collector[lista_t[17]] = assuntos_collect
                        # Imagens
                        imagens_collect = []
                        for i in range(52, 57):
                            if not pd.isnull(row[i]):
                                name_image = download_image(
                                    row[i], f'{row[1]}_{i}.png')
                                imagens_collect.append(name_image)
                        collector[lista_t[18]] = imagens_collect
                        # Substiui as imagens no texto
                        collector[lista_t[1]] = replace_images(
                            collector[lista_t[1]], imagens_collect)

                    # Salva o json
                    salvar_json(collector, name_arq)
                    print(f"ccouter: {ccouter}")
                    print(f"itenslen: {itenslen}")
                    self.update_progress_s.emit(int((ccouter*100)/itenslen))
                    ccouter += 1
                self.update_progress_s.emit(int((itenslen*100)/itenslen))
                self.update_rotulo_s.emit(
                    f"{ccouter} linhas de {itenslen} convertidas...")
        self.update_rotulo_s.emit("Conversão finalizada com sucesso")


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    mainWindow = MyFirstApp()
    mainWindow.show()
    app.exec_()
