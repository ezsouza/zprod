import re
import pandas as pd
import tkinter as tk
from tkinter import *
from tkinter import ttk
from ttkthemes import ThemedStyle
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from msedge.selenium_tools import Edge, EdgeOptions

# executar login no SIAC
def login():
    global driver
    # Create an EdgeOptions object
    edge_options = EdgeOptions()
    # Create an Edge WebDriver instance
    driver = Edge(executable_path='C:/WebDriver/msedgedriver.exe', options=edge_options)  # Replace with your path
    # coletar entrada de usuário e senha
    user = user_entry.get()
    password = password_entry.get()
    # acessar página de login''""
    driver.get('http://10.99.101.9/siac/do/Login?method=Pesquisar')
    # inserir usuário e senha
    driver.find_element(by=By.ID, value='codUsuario').send_keys(user)
    driver.find_element(by=By.ID, value='senhaUsuario').send_keys(password)
    # clicar no botão de login
    driver.find_element(by=By.ID, value='okButton').click()

# acessar página de inserção de produtos
def insert_product():
    # link da página de pesquisa/inserção
    driver.get('http://10.99.101.9/siac/do/produtos/ProdutosSubmit?method=Pesquisar')
    # clicar no botão de inserir
    driver.find_element(By.XPATH, "//span[text()='Inserir']").click()

def submit_form():
    selected = var.get()
    for sec_option in sec_options:
        if selected == sec_option[0]:
            sec_value = sec_option[1]
            break
    Select(driver.find_element(by=By.ID, value='codigoEstruturaMercadologica')).select_by_value(
        str(sec_value))

def submit_form_combo():
    selected = var_b.get()
    for sec_option_combo in sec_options_combo:
        if selected == sec_option_combo[0]:
            sec_value_combo = sec_option_combo[1]
            break
    Select(driver.find_element(by=By.ID, value='codigoEstruturaMercadologica')).select_by_value(
        str(sec_value_combo))

# informações de cadastro do produto
def product_information():
    # coletar informações de cadastro
    internal_code = internalcode_entry.get()
    description_name = description_entry.get()
    shortdescription_name = shortdescription_entry.get()
    # inserir informações de cadastro
    driver.find_element(By.ID, 'codigoInterno').send_keys(internal_code)
    driver.find_element(By.ID, 'descricaoProduto').send_keys(description_name.upper())
    driver.find_element(By.ID, 'descricaoResumidaProduto').send_keys(shortdescription_name.upper())
    # marcar os checkbox de cadastro
    driver.find_element(By.ID, 'indHabilitado').click()
    driver.find_element(By.ID, 'indProducaoPropria').click()
    driver.find_element(By.ID, 'indicProdutoCardapio').click()
    driver.find_element(By.ID, 'indicadorMarcaPropria').click()

def get_tax_value_initial(value): # Função para retornar o valor numérico correspondente ao valor inicial
    for initial, num_value in option_taxation: # Percorre a lista de tuplas
        if value.startswith(initial): # Verifica se o valor inicial corresponde
            return num_value # Retorna o valor numérico correspondente
    return None  # Retorna None se não houver correspondência

def get_cst_value_initial(value): # Função para retornar o valor numérico correspondente ao valor inicial
    for initial, num_value in option_cst: # Percorre a lista de tuplas
        if value.startswith(initial): # Verifica se o valor inicial corresponde
            return num_value # Retorna o valor numérico correspondente
    return None  # Retorna None se não houver correspondência

def product_profile():
    # coletar informações do perfil selecionado
    selected_item = profile_var.get()
    # coleta o index do item selecionado
    selected_index = list_type.index(selected_item)
    # coleta os valores equivalentes ao index do tipo de perfil selecionado
    ncm = str(data_excel.iloc[selected_index, init_column + 1])
    cest = str(data_excel.iloc[selected_index, init_column + 2])
    tax = data_excel.iloc[selected_index, init_column + 3]
    cfop = str(data_excel.iloc[selected_index, init_column + 4]).split('.')[0]
    tax_federal = str(data_excel.iloc[selected_index, init_column + 5])
    tax_state = str(data_excel.iloc[selected_index, init_column + 6])
    aliquot_pis = str(data_excel.iloc[selected_index, init_column + 7])
    cst_pis = data_excel.iloc[selected_index, init_column + 8]
    aliquot_cofins = str(data_excel.iloc[selected_index, init_column + 9])
    cst_cofins = data_excel.iloc[selected_index, init_column + 10]
    # coleta o valor numérico correspondente ao valor da inicial no excel
    tax_value = str(get_tax_value_initial(tax))
    cst_pis_value = str(get_cst_value_initial(cst_pis))
    cst_cofins_value = str(get_cst_value_initial(cst_cofins))
    # seleciona origem do produto (0 = Nacional)
    Select(driver.find_element(by=By.ID, value='codigoOrigemProduto')).select_by_value('0')
    # enviar valor do ncm
    driver.find_element(by=By.ID, value='codigoNCM').send_keys(ncm)
    # enviar valor do cest
    driver.find_element(by=By.ID, value='cest').send_keys(cest)
    # pressiona o botão da aba perfil
    driver.find_element(By.ID, 'divabaTributacao').click()
    # pressiona o botão inserir
    driver.find_element(by=By.XPATH, value="//span[text()='Inserir']").click()
    # marcar o checkbox permite digitação de desconto e vende pelo código interno
    driver.find_element(By.ID, 'indicadorPermiteDigitacaoDesconto').click()
    driver.find_element(By.ID, 'indicVendeCodInterno').click()
    # acessa o menu de tributação
    driver.find_element(By.ID, 'divabaTributacao').click()
    # envia o valor coletado da tributação de acordo com o perfil selecionado
    Select(driver.find_element(by=By.ID, value='tributacao')).select_by_value(str(tax_value))
    # envia o valor do cfop
    driver.find_element(by=By.ID, value='CFOP').send_keys(cfop)
    # envia o valor da tributação federal
    driver.find_element(by=By.ID, value='percTributacao').send_keys(str('00' + tax_federal))
    # envia o valor da tributação estadual
    driver.find_element(by=By.ID, value='percTribEst').send_keys(str(tax_state))
    # acessar a aba de pis e cofins
    driver.find_element(By.ID, 'divabaPisCofins').click()
    # envia o valor da aliquota do pis
    driver.find_element(by=By.ID, value='percPis').send_keys(str(aliquot_pis))
    # envia o valor da aliquota do cofins
    driver.find_element(by=By.ID, value='percCofins').send_keys(str(aliquot_cofins))
    # envia o valor do cst do pis
    Select(driver.find_element(by=By.ID, value='cstPis')).select_by_value(str(cst_pis_value))
    # envia o valor do cst do cofins
    Select(driver.find_element(by=By.ID, value='cstCofins')).select_by_value(str(cst_cofins_value))
    # marcar o checkbox de produto ativo
    driver.find_element(By.ID, 'indicProdutoAtivo').click()
    # selecionar todas as lojas
    select_stores = Select(driver.find_element(by=By.ID, value='lojasSelecionadasTrib'))
    options = select_stores.options
    for option in options:
        option.click()
    # confirmar perfil do produto
    driver.find_element(by=By.XPATH, value="//span[text()='Confirmar']").click()

def product_price():
    driver.find_element(By.ID, 'divabaPrecos').click()
    sell_date = date_entry.get()
    date_regex = re.compile(r'\b\d{2}/\d{2}/\d{4}\b')
    match = date_regex.match(sell_date)
    if match:
        date_entry.config(style='White.TEntry')
        driver.find_element(By.ID, "dataInicioValidade").send_keys(sell_date)
    else:
        date_entry.config(style='Red.TEntry')

    sell_price = price_entry.get()
    try:
        sell_price = float(sell_price)
        price_entry.config(style='White.TEntry')
        driver.find_element(By.ID, 'valorPrecoVenda').send_keys(str(sell_price))
    except ValueError:
        price_entry.config(style='Red.TEntry')
    driver.find_element(By.ID, 'valorPrecoCusto').send_keys('0,01')
    sao_paulo = driver.find_element(By.ID, 'tr7')
    driver.execute_script('arguments[0].click();', sao_paulo)
    driver.find_element(By.XPATH,
        "//a[@href=\"javascript:executarSubmit(document.produtoForm, 'Adicionar Preço/Loja');\"]").click()

def combo_information():
    combo_internalcode = combo_internalcode_entry.get()
    combo_description = combo_description_entry.get()
    combo_shortdescription = combo_shortdescription_entry.get()
    driver.find_element(By.ID, 'codigoInterno').send_keys(combo_internalcode)
    driver.find_element(By.ID, 'descricaoProduto').send_keys(combo_description.upper())
    driver.find_element(By.ID, 'descricaoResumidaProduto').send_keys(combo_shortdescription.upper())
    driver.find_element(By.ID, 'indHabilitado').click()
    # seleciona origem do produto (0 = Nacional)
    Select(driver.find_element(by=By.ID, value='codigoOrigemProduto')).select_by_value('0')
    # seleciona o valor do tipo de produto (3 = Cesta)
    Select(driver.find_element(by=By.ID, value='codigoTipoProduto')).select_by_value('3')

def combo_profile():
    # coletar informações do perfil selecionado
    selected_item_combo = profile_combo_var.get()
    # coleta o index do item selecionado
    selected_index_combo = list_type.index(selected_item_combo)
    # coleta os valores equivalentes ao index do tipo de perfil selecionado
    ncm = str(data_excel.iloc[selected_index_combo, init_column + 1])
    cest = str(data_excel.iloc[selected_index_combo, init_column + 2])
    # seleciona origem do produto (0 = Nacional)
    Select(driver.find_element(by=By.ID, value='codigoOrigemProduto')).select_by_value('0')
    # enviar valor do ncm
    driver.find_element(by=By.ID, value='codigoNCM').send_keys(ncm)
    # enviar valor do cest
    driver.find_element(by=By.ID, value='cest').send_keys(cest)
    # marcar os checkbox de cadastro
    driver.find_element(By.ID, 'indProducaoPropria').click()
    driver.find_element(By.ID, 'indicProdutoCardapio').click()
    driver.find_element(By.ID, 'indicadorMarcaPropria').click()

def combo_store():
    driver.find_element(by=By.ID, value='divabaCesta').click()
    select_stores = Select(driver.find_element(by=By.ID, value='lojaCestaDisponivelSel'))
    options = select_stores.options
    for option in options:
        option.click()
    driver.find_element(By.XPATH,
        '//*[@id="abaCestaLojas"]/table/tbody/tr/td/table/tbody/tr[3]/td[2]/a[1]').click()
    driver.find_element(by=By.ID, value='divabaCestaComponente').click()

def save():
    driver.find_element(By.XPATH, '//*[@id="mainContent"]/form/div/table/tbody/tr/td[4]/table/tbody/tr/td[1]/a/span[2]').click()

def exec_product():
    insert_product()
    product_information()
    submit_form()
    product_profile()
    product_price()

def exec_combo():
    insert_product()
    combo_information()
    submit_form_combo()
    combo_profile()
    combo_store()
    send_entry_values()
    driver.find_element(By.ID, 'divabaCestaComponentePrecos').click()
    sao_paulo_stores = driver.find_element(By.ID, 'tr7')
    driver.execute_script('arguments[0].click();', sao_paulo_stores)


# carregar o arquivo excel
path_excel = r'\\SJSRVSLC\b\Publica Nova\CARDÁPIO\Planilha de Tributação.xlsx'
# ler o arquivo Excel
data_excel = pd.read_excel(path_excel)
# definir a linha inicial [index 1] = [linha 2]
init_row = 1
# definir a coluna inicial [index 0] = [coluna 1]
init_column = 0
# definir as colunas do arquivo excel [linha, coluna]
column_type = data_excel.iloc[0:, 0].dropna() # dropna() = remover linhas vazias
# converter as colunas em listas
list_type = column_type.tolist()
# definir as opções de taxação
option_taxation = [
    ('F', 1), ('I', 2), ('N', 3), ('T0', 4), ('T1', 5), ('T2', 6),
    ('T3', 7), ('T4', 8), ('T5', 9), ('T6', 10), ('T7', 11)
]
# definir as opções de cst
option_cst = [
    ('1 - Operação Tributável (base de cálculo = valor da operação alíquota normal (cumulativo/não cumulativo))', 1),
    ('2 - Operação Tributável (base de cálculo = valor da operação (alíquota diferenciada))', 2),
    ('4 - Operação Tributável (tributação monofásica (alíquota zero))', 4),
    ('5 - Operação Tributável (Substituição Tributária)', 5),
    ('6 - Operação Tributável (alíquota zero)', 6),
    ('7 - Operação Isenta da Contribuição', 7),
    ('8 - Operação Sem Incidência da Contribuição', 8),
    ('9 - Operação com Suspensão da Contribuição', 9),
    ('49 - Outras Operações de Saída', 49)
]
# inicializar o programa (GUI)
root = tk.Tk()
# nomear a janela
root.title('ZProd')
# definir estilo da janela
style = ThemedStyle(root)
style.set_theme('adapta')
# definir tamanho da janela (largura x altura)
root.resizable(False, False)
# criar um título centralizado na janela
title_label = ttk.Label(root, text='◄ ZProd ►', font=('Times New Roman',
                        18, 'bold'), anchor='center', justify='center')
title_label.pack(side='top', fill='both')
# criar um notebook para separar as abas do sistema
notebook = ttk.Notebook(root)
notebook.pack(fill='both', expand=True)
# difinir as abas do sistema
login_frame = ttk.Frame(notebook)
product_frame = ttk.Frame(notebook)
combo_frame = ttk.Frame(notebook)
# adicionar as abas ao notebook e nomeá-las
notebook.add(login_frame, text='Login')
notebook.add(product_frame, text='Produtos')
notebook.add(combo_frame, text='Combos')
# label e insert do usuário
user_label = ttk.Label(login_frame, text='Usuário: ')
user_label.grid(row=0, column=0, padx=10, pady=10)
user_entry = ttk.Entry(login_frame)
user_entry.grid(row=0, column=1, padx=10, pady=10)
# label e insert da senha
password_label = ttk.Label(login_frame, text='Senha: ')
password_label.grid(row=1, column=0, padx=10, pady=10)
password_entry = ttk.Entry(login_frame, show='*')
password_entry.grid(row=1, column=1, padx=10, pady=10)
# botão para executar a função de login
login_button = ttk.Button(login_frame, text='Login', command=login)
login_button.grid(row=2, column=0, columnspan=2, padx=10, pady=10, sticky='e')
# executar a função de login ao pressionar a tecla Enter
password_entry.bind('<Return>', lambda event: login())
# label e insert do código interno
internalcode_label = ttk.Label(product_frame, text='Código Interno: ')
internalcode_label.grid(row=2, column=0, padx=10, pady=10)
internalcode_entry = ttk.Entry(product_frame)
internalcode_entry.grid(row=2, column=1, padx=10, pady=10)
# label e insert da descrição
description_label = ttk.Label(product_frame, text='Descrição: ')
description_label.grid(row=3, column=0, padx=10, pady=10)
description_entry = ttk.Entry(product_frame)
description_entry.grid(row=3, column=1, padx=10, pady=10)
# label e insert da descrição reduzida
shortdescription_label = ttk.Label(product_frame, text='Descrição Reduzida: ')
shortdescription_label.grid(row=4, column=0, padx=10, pady=10)
shortdescription_entry = ttk.Entry(product_frame)
shortdescription_entry.grid(row=4, column=1, padx=10, pady=10)
# opções de seção
sec_options = [
    ('Selecione', 0), ('0001 Gyu-Don', 2), ('0002 Curry', 3), ('0003 Especial', 4),
    ('0004 Mega', 5), ('0005 Combos', 6), ('0006 Kids', 7), ('0007 Bebidas', 8),
    ('0008 Teishoku', 9), ('0009 Acompanhamentos', 10), ('0010 Mercado', 11),
    ('0011 Lamen', 12), ('0012 Quiosque', 13), ('0013 SobreMesa', 15), ('0014 Promocao', 16),
    ('0015 Salmao', 17), ('0016 Ifood', 18), ('0017 NewWave Bowl', 19), ('0018 Executivos', 20),
    ('0019 Tonkatsu', 21), ('0020 Hayashi', 22), ('0021 Shogayakidon', 23), ('0022 Yakissoba', 24),
    ('0023 Brinquedo', 25), ('0024 Yakiniku', 26), ('0025 Viagem', 27), ('0026 Teriyaki', 28)
]
# variável para armazenar a opção selecionada
var = tk.StringVar()
# definir a opção padrão
var.set(sec_options[0][0])
# label e dropdown para selecionar a seção
section_label = ttk.Label(product_frame, text='Seção: ')
section_label.grid(row=0, column=0, padx=10, pady=10)
section_dropdown = ttk.OptionMenu(product_frame, var, *[sec_option[0] for sec_option in sec_options])
section_dropdown.grid(row=0, column=1, padx=10, pady=10)
# label e insert da data de início
date_label = ttk.Label(product_frame, text='Data de Início (DD/MM/AAAA): ')
date_label.grid(row=5, column=0, padx=10, pady=10)
date_entry = ttk.Entry(product_frame)
date_entry.grid(row=5, column=1, padx=10, pady=10)
# label e insert do valor do produto
price_label = ttk.Label(product_frame, text='Valor do Produto (Apenas Números): ')
price_label.grid(row=6, column=0, padx=10, pady=10)
price_entry = ttk.Entry(product_frame)
price_entry.grid(row=6, column=1, padx=10, pady=10)
# valor padrão para o perfil do item
profile_var = tk.StringVar(value=list_type[0])
# label e dropdown para selecionar o perfil do item
profile_label = ttk.Label(product_frame, text='Perfil do Item: ')
profile_label.grid(row=1, column=0, padx=10, pady=10)
profile_dropdown = ttk.OptionMenu(product_frame, profile_var, *list_type)
profile_dropdown.grid(row=1, column=1, padx=10, pady=10)
# botão para executar a função de inserir produto
execute_button = ttk.Button(product_frame, text='Executar', command=exec_product)
execute_button.grid(row=7, column=1, padx=10, pady=10, sticky='we')
# botão para salvar a configuração do produto
save_button = ttk.Button(product_frame, text='Salvar', command=save)
save_button.grid(row=7, column=2, padx=10, pady=10, sticky='we')
# -------------------------
# -------- COMBO ----------
# -------------------------
# título para informações gerais do combo
geral_title = ttk.Label(combo_frame, text='Geral', font=('Times New Roman', 14, 'bold'))
geral_title.grid(row=0, column=0, columnspan=2, pady=10)
# label e insert do código interno
combo_internalcode_label = ttk.Label(combo_frame, text='Código Interno:')
combo_internalcode_label.grid(row=0, column=2, padx=10, pady=10)
combo_internalcode_entry = ttk.Entry(combo_frame)
combo_internalcode_entry.grid(row=0, column=3, padx=10, pady=10)
# label e insert da descrição
combo_description_label = ttk.Label(combo_frame, text='Descrição:')
combo_description_label.grid(row=1, column=0, padx=10, pady=10)
combo_description_entry = ttk.Entry(combo_frame)
combo_description_entry.grid(row=1, column=1, padx=10, pady=10)
# label e insert da descrição reduzida
combo_shortdescription_label = ttk.Label(combo_frame, text='Descrição Reduzida:')
combo_shortdescription_label.grid(row=1, column=2, padx=10, pady=10)
combo_shortdescription_entry = ttk.Entry(combo_frame)
combo_shortdescription_entry.grid(row=1, column=3, padx=10, pady=10)
# função para selecionar a seção do combo
def submit_form_combo():
    selected = var_b.get() # pega a opção selecionada
    for sec_option_combo in sec_options_combo: # percorre a lista de opções
        if selected == sec_option_combo[0]: # se a opção selecionada for igual a opção da lista
            sec_value_combo = sec_option_combo[1] # pega o valor da opção
            break # para o loop
    Select(driver.find_element(by=By.ID, value='codigoEstruturaMercadologica')).select_by_value(
        str(sec_value_combo)) # seleciona a seção
# opções de seção
sec_options_combo = [
    ('Selecione', 0), ('0001 Gyu-Don', 2), ('0002 Curry', 3), ('0003 Especial', 4),
    ('0004 Mega', 5), ('0005 Combos', 6), ('0006 Kids', 7), ('0007 Bebidas', 8),
    ('0008 Teishoku', 9), ('0009 Acompanhamentos', 10), ('0010 Mercado', 11),
    ('0011 Lamen', 12), ('0012 Quiosque', 13), ('0013 SobreMesa', 15), ('0014 Promocao', 16),
    ('0015 Salmao', 17), ('0016 Ifood', 18), ('0017 NewWave Bowl', 19), ('0018 Executivos', 20),
    ('0019 Tonkatsu', 21), ('0020 Hayashi', 22), ('0021 Shogayakidon', 23), ('0022 Yakissoba', 24),
    ('0023 Brinquedo', 25), ('0024 Yakiniku', 26), ('0025 Viagem', 27), ('0026 Teriyaki', 28)
]
# variável B para armazenar a opção selecionada
var_b = tk.StringVar()
# definir a opção padrão
var_b.set(sec_options_combo[0][0])
# label e dropdown para selecionar a seção
combo_section_label = ttk.Label(combo_frame, text='Seção:')
combo_section_label.grid(row=2, column=0, padx=10, pady=10)
combo_dropdown = ttk.OptionMenu(combo_frame, var_b, *[sec_option_combo[0] for sec_option_combo in sec_options_combo])
combo_dropdown.grid(row=2, column=1, padx=10, pady=10)
# valor padrão para o perfil do item
profile_combo_var = tk.StringVar(value=list_type[0])
# label e dropdown para selecionar o perfil do item
profile_combo_label = ttk.Label(combo_frame, text='Perfil do Item: ')
profile_combo_label.grid(row=2, column=2, padx=10, pady=10)
profile_combo_dropdown = ttk.OptionMenu(combo_frame, profile_combo_var, *list_type)
profile_combo_dropdown.grid(row=2, column=3, padx=10, pady=10)
# título para componentes do combo
components_title = ttk.Label(combo_frame, text='Componentes', font=('Times New Roman', 14, 'bold'))
components_title.grid(row=4, column=0, columnspan=2, pady=10)
# dicionário para armazenar os campos de entrada
entry_fields = {} 

# função para atualizar os campos de entrada
def update_entry_fields(*args):
    num_fields = int(selected_option.get()) # pegar o valor da opção selecionada
    for entry in entry_fields.values(): # percorrer os valores do dicionário
        entry.destroy() # destruir o campo de entrada
    entry_fields.clear() # limpar o dicionário
    for widget in combo_frame.grid_slaves(): # percorrer os widgets do frame
        if int(widget.grid_info()['row']) >= 6: # se o widget estiver na linha 6 ou superior
            widget.destroy() # destruir o widget
    for i in range(num_fields): # percorrer o número de campos de entrada
        row = 6 + (i // 2)  # dividir i por 2 para obter o número da linha a partir da linha 6
        col = (i % 2) * 2  # use o resto para obter o número da coluna mutiplicado por 2 para obter o label e o entry
        entry = ttk.Entry(combo_frame) # criar o campo de entrada
        entry.grid(row=row, column=col+1, padx=10, pady=10) # colocar o campo de entrada
        labelcode = ttk.Label(combo_frame, text='Código ' + str(i + 1)) # criar o label
        labelcode.grid(row=row, column=col, padx=10, pady=10) # posicionar o label
        entry_fields[labelcode] = entry # adicionar o label e o campo de entrada ao dicionário

# envia valores dos campos de entrada
def send_entry_values(*args):
    for label, entry in entry_fields.items(): # percorrer os itens do dicionário
        value = entry.get() # pegar o valor do campo de entrada
        internal_code_combo = driver.find_element(By.ID, 'codigoInternoCesta') # pegar o campo de entrada do código interno
        internal_code_combo.clear() # limpar o campo de entrada
        internal_code_combo.send_keys(value) # enviar o valor do campo de entrada
        driver.find_element(By.ID, 'quantidadeProdutoCesta').send_keys('1') # enviar o valor 1 para o campo de quantidade
        add_internal_code_combo = driver.find_element(By.XPATH, '//*[@id="abaCestaComponente"]/table/tbody/tr[1]/td[5]/a[1]') # pegar o botão adicionar
        add_internal_code_combo.click() # clicar no botão adicionar

# botão para enviar os valores dos campos de entrada
selected_option = tk.StringVar(combo_frame)
# definir a opção padrão
selected_option.set('1')
# menu de opções de 1 a 8 campos de entrada
option_menu = ttk.OptionMenu(combo_frame, selected_option, '0', '1', '2', '3', '4', '5', '6', '7', '8')
option_menu.grid(row=5, column=1, padx=10, pady=10)
# label para o menu de opções
option_menu_label = ttk.Label(combo_frame, text='Itens: ')
option_menu_label.grid(row=5, column=0,padx=10, pady=10)
# chamar a função update_entry_fields sempre que a opção selecionada mudar
selected_option.trace("w", update_entry_fields)
# executar cadastro de combo
save_button = ttk.Button(combo_frame, text='Executar', command=exec_combo)
save_button.grid(row=5, column=3, padx=10, pady=10)
# manter execução do programa
root.mainloop()