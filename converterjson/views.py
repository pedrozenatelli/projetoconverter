import json
import os
import re
from tempfile import NamedTemporaryFile
from zipfile import ZipFile

import pandas as pd
import requests
from django.conf import settings
from django.core.files.base import ContentFile
from django.core.files.storage import FileSystemStorage, default_storage
from django.http import FileResponse, Http404, HttpResponseRedirect
from django.shortcuts import redirect, render
from django.urls import reverse
from pandas import isnull, read_excel


def download_file(request, file_name):
    # Aqui, estou assumindo que você armazena os arquivos em uma pasta chamada 'media'
    file_path = os.path.join(settings.MEDIA_ROOT, file_name)

    if os.path.exists(file_path):
        response = FileResponse(open(file_path, 'rb'))
        response['Content-Disposition'] = f'attachment; filename="{file_name}"'
        return response
    else:
        # O arquivo não foi encontrado
        raise Http404


def clear_json_files_in_saida():
    # Definindo o caminho para a pasta 'saida'
    saida_directory = os.path.join(os.getcwd(), 'saida')

    # Iterando sobre todos os arquivos na pasta 'saida'
    for filename in os.listdir(saida_directory):
        # Se o arquivo terminar com '.json', exclua-o
        if filename.endswith('.json'):
            file_path = os.path.join(saida_directory, filename)
            try:
                os.remove(file_path)
            except Exception as e:
                print(f"Error deleting file {file_path}: {e}")


def conversion_success(request):
    converted_files = request.session.get('converted_files', [])

    # Se você quiser limpar os arquivos da sessão após acessá-los (opcional)
    if 'converted_files' in request.session:
        del request.session['converted_files']

    return render(request, 'conversion_success.html', {'converted_files': converted_files})


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
        with open(os.path.join(os.getcwd(), 'saida', 'images', name), 'wb') as f:
            f.write(image.content)
        return name
    else:
        return ''


def cria_json(nome, value):
    # Define o caminho do arquivo
    caminho_arquivo = os.path.join(os.getcwd(), 'saida', f'{nome}.json')

    # Se você quiser adicionar uma verificação de existência e criar um novo nome:
    # counter = 1
    # while os.path.exists(caminho_arquivo):
    #     nome = f"{nome}_{counter}"
    #     caminho_arquivo = os.path.join(os.getcwd(), 'saida', f'{nome}.json')
    #     counter += 1

    # Salva o arquivo JSON
    with open(caminho_arquivo, 'w', encoding='utf-8') as file:
        json.dump(value, file, indent=4, ensure_ascii=False)

    # Retornar apenas o nome do arquivo sem a extensão
    return f"{nome}.json"


def salvar_json(value, name_arq):
    # Define o caminho do arquivo
    caminho_arquivo = os.path.join(os.getcwd(), 'saida', f'{name_arq}')
    # # Abre o arquivo JSON
    # with open(caminho_arquivo, 'r') as file:
    #     data = json.load(file)

    # # Adiciona o novo valor
    # data.append(value)

    # # Salva o arquivo JSON
    # with open(caminho_arquivo, 'w', encoding='utf-8') as file:
    #     json.dump(data, file, indent=4, ensure_ascii=False)


def start(request):
    if request.method == 'POST':

        uploaded_files = request.FILES.getlist('documents')
        errors = []
        converted_files = []

        for uploaded_file in uploaded_files:
            collector_full = []
            if uploaded_file.name.endswith('.xlsx'):
                fs = FileSystemStorage()
                name = fs.save(uploaded_file.name, uploaded_file)
                # caminho absoluto para o arquivo
                file_path = fs.path(name)
                try:
                    # Verifica se é possível ler o arquivo como um arquivo do Excel
                    df = pd.read_excel(file_path)
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
                        elif item not in ['A', 'B', 'C', 'D', 'E', 'Prova1', 'Prova2', 'Prova3', 'Prova4', 'Prova5', 'Prova6', 'Prova7', 'Prova8', 'Prova9', 'Prova10', 'Prova11', 'Prova12', 'Prova13', 'Prova14', 'Prova15', 'Assunto1', 'Assunto2', 'Assunto3', 'Assunto4', 'Assunto5', 'Assunto6', 'Assunto7', 'Assunto8', 'Assunto9', 'Assunto10', 'Assunto11', 'Assunto12', 'Assunto13', 'Assunto14', 'Assunto15', 'IMAGEM1', 'IMAGEM2', 'IMAGEM3', 'IMAGEM4', 'IMAGEMRESP1', 'IMAGEMRESP2', 'IMAGEMRESP3', 'IMAGEMALTA', 'IMAGEMALTB', 'IMAGEMALTC', 'IMAGEMALTD', 'IMAGEMALTE'] and not pd.isnull(item):
                            nova_lista_t.append(item)

                    nova_lista_t.append('Provas')
                    nova_lista_t.append('Assuntos')
                    nova_lista_t.append('Imagens_Perg')
                    nova_lista_t.append('Imagens_Resp')
                    nova_lista_t.append('Imagens_Alt')

                    # Atribui a nova lista à lista_t
                    lista_t = nova_lista_t
                    itenslen = len(list(df.itertuples()))
                    # Iterando sobre as linhas do DataFrame

                    for row in df.itertuples():
                        collector = {}
                        if row.Index == 0:
                            continue
                        else:
                            # Codigo
                            collector[lista_t[0]] = int(
                                row[1]) if not pd.isnull(
                                row[1]) else ''
                            # Pergunta
                            collector[lista_t[1]] = row[2].replace(
                                '\n', ' ').replace('\xa0', ' ') if not pd.isnull(
                                row[2]) else ''
                            # Comentario do professor
                            collector[lista_t[2]] = row[3] if not pd.isnull(
                                row[3]) else ''
                            # Tipo
                            collector[lista_t[3]] = row[4] if not pd.isnull(
                                row[4]) else ''
                            # Materia
                            collector[lista_t[4]] = row[5] if not pd.isnull(
                                row[5]) else ''
                            # Disciplica
                            collector[lista_t[5]] = row[6] if not pd.isnull(
                                row[6]) else ''
                            # Banca
                            collector[lista_t[6]] = row[7] if not pd.isnull(
                                row[7]) else ''
                            # Órgão
                            collector[lista_t[7]] = row[8] if not pd.isnull(
                                row[8]) else ''
                            # Ano
                            collector[lista_t[8]] = row[9] if not pd.isnull(
                                row[9]) else ''
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
                                    provas_collect.append(
                                        row[i])
                            collector[lista_t[16]] = provas_collect
                            # Assuntos
                            assuntos_collect = []
                            for i in range(37, 52):
                                if not pd.isnull(row[i]):
                                    assuntos_collect.append(
                                        row[i])
                            collector[lista_t[17]] = assuntos_collect
                            # Imagens Pergunta
                            imagens_perg = []
                            for i in range(52, 56):
                                if not pd.isnull(row[i]):
                                    imagens_perg.append(
                                        row[i])

                            collector[lista_t[18]] = imagens_perg

                            # Imagens Resposta
                            imagens_resp = []
                            for i in range(56, 59):
                                if not pd.isnull(row[i]):
                                    imagens_resp.append(
                                        row[i])

                            collector[lista_t[19]] = imagens_resp

                            # Imagens Alternativa
                            imagens_alt = []
                            for i in range(59, 64):
                                if not pd.isnull(row[i]):
                                    imagens_alt.append(
                                        row[i])

                            collector[lista_t[20]] = imagens_alt

                            # Insere tudo dentro de collector full
                            collector_full.append(collector)

                    # cria um arquivo json em branco com o mesmo nome base
                    base_name = os.path.splitext(name)[0]  # remove a extensão
                    base_novo = base_name.split("_")[0]

                    json_path = cria_json(base_novo, collector_full)

                    # Ajuste o caminho conforme sua configuração.
                    converted_files.append(f"/saida/{json_path}")
                    # Armazena os links dos arquivos JSON na sessão
                    request.session['converted_files'] = converted_files

                    # Exclui arquivo inutil
                    fs.delete(name)

                except Exception as e:
                    # remova o arquivo se ele não for válido
                    fs.delete(name)
                    errors.append(f'Arquivo {uploaded_file.name} inválido.')
                    continue  # continue para o próximo arquivo na lista

            else:
                errors.append(f'Arquivo {uploaded_file.name} inválido.')

        if errors:
            return render(request, 'start.html', {'error': ' '.join(errors)})

        # Redireciona para a página de sucesso após a conversão
        return redirect('conversion_success')
    else:
        # Limpa todos os json da pagina
        clear_json_files_in_saida()
        return render(request, 'start.html')
