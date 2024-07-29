from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from time import sleep
from datetime import datetime
import openpyxl
import os
import schedule

def iniciar_driver():
    chrome_options = Options()

    arguments = ['--lang=pt-BR', '--window-size=1200,1400', '--incognito']

    for argument in arguments:
        chrome_options.add_argument(argument)

    caminho_padrao_para_download = 'D:\\Adilson\\Python\\2024\\monitorar_precos_destrava'

    chrome_options.add_experimental_option("prefs", {
        'download.default_directory': caminho_padrao_para_download,
        'download.directory_upgrade': True,
        'download.prompt_for_download': False,
        "profile.default_content_setting_values.notifications": 2, 
        "profile.default_content_setting_values.automatic_downloads": 1,
    })

    driver = webdriver.Chrome(options=chrome_options)
    return driver

def layout():
    choice_one = '''\n###############
1 - Macbook
2 - Iphone
###############'''
    choice_two = '''\nVocê deseja pesquisar outro produto após 30 minutos?
1 - Sim
2 - Não
'''
    return choice_one, choice_two

def product_choice():
    while True:
        print(layout()[0])
        choice = input('\nDigite o número da sua escolha: ').strip()

        if choice == '1' or choice == 'um':
            product_name = 'Macbook'
            break
        elif choice == '2' or choice == 'dois':
            product_name = 'Iphone'
            break
        else:
            os.system('cls' if os.name == 'nt' else 'clear')
            print('Opção inválida, tente novamente!!!\n')
            pass

    while True:
        print(f'\nProduto escolhido: {product_name}')
        print(layout()[1])
        choice_permanent = input('Digite o número da sua escolha: ').strip()
        if choice_permanent == '1' or choice_permanent == 'um':
            choice_permanent = 'Sim'
            break
        elif choice_permanent == '2' or choice_permanent == 'dois':
            choice_permanent = 'Nao'
            break
        else:
            os.system('cls' if os.name == 'nt' else 'clear')
            print('Opção inválida, tente novamente!!!\n')
            pass

    url_site = f'https://www.buscape.com.br/search?q={product_name}'
    os.system('cls' if os.name == 'nt' else 'clear')

    print(f'Produto escolhido: {product_name}')
    print(f'Deseja escolher outro produto após 30 minutos: {choice_permanent}')

    return url_site, choice, choice_permanent
    
# ○ Verificar o preço atual. # ○ Guardar o valor do preço(somente o valor numérico, não em texto)
def extract_name_and_price(url_site, choice):
    driver = iniciar_driver()
    driver.get(url_site)
    sleep(10)
    driver.execute_script("window.scrollTo(0, 200);")
    sleep(2)

    if choice == '1' or choice == 'um':
        xpath_product = '//*[@id="product-card-5673969::name"]'
        product_name = driver.find_element(By.XPATH, xpath_product)
        product_name = product_name.text

        xpath_price = '//*[@id="__next"]/main/div[2]/div[7]/div[1]/div/a/div[2]/div[2]/div[2]/p'
        price = driver.find_element(By.XPATH, xpath_price)
        sleep(3)
        price = price.text
    elif choice == '2' or choice == 'dois':
        xpath_product = '//*[@id="product-card-12474754::name"]'
        product_name = driver.find_element(By.XPATH, xpath_product)
        product_name = product_name.text

        xpath_price = '//*[@id="__next"]/main/div[2]/div[7]/div[1]/div/a/div[2]/div[2]/div[2]/p'
        price = driver.find_element(By.XPATH, xpath_price)
        sleep(3)
        price = price.text

    # tratar campo preço
    price = price.replace('R$', '')  
    price = price.replace('.', '')  
    price = price.replace(',', '.')
    price = price.replace(' ', '')
    price = float(price)

    os.system('cls' if os.name == 'nt' else 'clear')
    print(product_name)
    print(price)

    return driver, product_name, price

def create_workbook(product_name, price, url):

    filename = 'acompanhamento_de_precos.xlsx'
    date_now = datetime.now().strftime('%Y-%m-%d')

    if os.path.exists(filename):
        workbook = openpyxl.load_workbook(filename)
        worksheet = workbook.active
    else:
        workbook = openpyxl.Workbook()
        worksheet = workbook.active
        worksheet.title = 'Acompanhamento de Preços'
        worksheet.append(['Produto', 'Data atual', 'Valor', 'Link'])

    # Adiciona os dados
    worksheet.append([product_name, date_now, price, url])

    # Salva a planilha
    workbook.save(filename)

    print('Dados salvos na planilha "acompanhamento_de_precos.xlsx" com sucesso!\n')


def main(url_site, choice):
    driver, product_name, price = extract_name_and_price(url_site, choice)
    create_workbook(product_name, price, url_site)
    driver.quit()

    
if __name__ == '__main__':
    os.system('cls' if os.name == 'nt' else 'clear')
    url_site, choice, choice_permanent = product_choice()
    while True:
        if choice_permanent == 'Nao':
            url_site = url_site
            main(url_site, choice)
            # agendamento para rodar a cada 30 minutos
            schedule.every(30).minutes.do(main(url_site, choice))
            while True:
                schedule.run_pending()
                sleep(1)
        elif choice_permanent == 'Sim':
            main(url_site, choice)
            # agendamento para rodar a cada 30 minutos usando while + sleep
            sleep(1800)
            url_site, choice, choice_permanent = product_choice()