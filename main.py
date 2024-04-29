import pandas as pd
import PySimpleGUI as sg

def salvar_arquivo(dados_filtrados):
    nome_arquivo = sg.popup_get_file('Salvar como:', save_as=True, default_extension=".xlsx")
    if nome_arquivo:
        dados_filtrados.to_excel(nome_arquivo, index=False)
        sg.popup("Arquivo salvo com sucesso!")

# Carregar o arquivo Excel
caminho_do_arquivo = 'Base_de_dados.xlsx'
dados_excel = pd.read_excel(caminho_do_arquivo)

# Layout da interface
layout = [
    [sg.Text("Coluna de Filtro:"), sg.Combo(dados_excel.columns, key='-COLUNA-', enable_events=True)],
    [sg.Text("Filtrar por:"), sg.Combo([], key='-FILTRO-', enable_events=True)],
    [sg.Button("Filtrar"), sg.Button("Sair")]
]

# Criar a janela
janela = sg.Window('Filtro e Salvamento de Dados', layout)

# Loop de eventos
while True:
    evento, valores = janela.read()
    if evento == sg.WINDOW_CLOSED or evento == 'Sair':
        break
    elif evento == '-COLUNA-':
        coluna_selecionada = valores['-COLUNA-']
        filtro_unico = dados_excel[coluna_selecionada].unique()
        janela['-FILTRO-'].update(values=filtro_unico)
    elif evento == 'Filtrar':
        filtro = valores['-FILTRO-']
        coluna = valores['-COLUNA-']
        dados_filtrados = dados_excel[dados_excel[coluna].str.contains(filtro, na=False)]
        salvar_arquivo(dados_filtrados)

# Fechar a janela
janela.close()
