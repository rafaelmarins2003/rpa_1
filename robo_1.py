import os
import csv
import pandas as pd
import logging
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from tabulate import tabulate
from time import sleep
from unidecode import unidecode
from datetime import date, timedelta, datetime
from _Lib import _config, _msg_erro_slack, _msg_slack, _envio_email, _sql_update, \
    _sql_select_valores_sql_pd, _sql_insert_many


# -----------------------------------------------------------------------------
# CONFIGURAÇÃO DO LOGGER
# -----------------------------------------------------------------------------
logger = logging.getLogger("robo_1")
logger.setLevel(logging.INFO)

# Formato das mensagens: data/hora, nível de log, nome do logger, mensagem
formatter = logging.Formatter("%(asctime)s [%(levelname)s] %(name)s - %(message)s")

# Define o arquivo de log (você pode mudar o path se preferir)
# Local
log_file_path = os.path.join(os.getcwd(), "logs", "robo_1_novo_v3.log")

# VM
# log_file_path = r"D:\logs\robo_1_novo_v2.log"


file_handler = logging.FileHandler(log_file_path, mode="a", encoding="utf-8")
file_handler.setFormatter(formatter)

logger.addHandler(file_handler)
# -----------------------------------------------------------------------------

# BANCO 1
LOGIN = 'login'
# SENHA = 'senha'

meses = {
    1: "Janeiro",
    2: "Fevereiro",
    3: "Marco",
    4: "Abril",
    5: "Maio",
    6: "Junho",
    7: "Julho",
    8: "Agosto",
    9: "Setembro",
    10: "Outubro",
    11: "Novembro",
    12: "Dezembro"
}

headers = ['DATA', 'DESCRIÇÃO', 'VALOR', 'TRANSAÇÃO']

dicio_clientes_excecao = {
    'NOME_EMPRESA': 'ID_SLACK_USUARIO'
}


def get_exe_directory():
    return os.path.join(os.getcwd(), "temp")


def iniciar_driver():
    try:
        # Local Linux
        chrome_options = Options()
        chrome_options.binary_location = "/usr/bin/chromium-browser"
        service = Service("/usr/bin/chromedriver")
        chrome_options.add_experimental_option("detach", True)
        chrome_options.add_argument("--timeout=900000")
        chrome_options.add_argument("--set-script-timeout=900000")
        chrome_options.add_argument("--disable-webgl")
        chrome_options.add_argument("--disable-extensions")
        # chrome_options.add_argument("--start-maximized")
        chrome_options.add_experimental_option("prefs", {
            "download.default_directory": get_exe_directory(),
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
        })
        driver = webdriver.Chrome(service=service, options=chrome_options)
        driver.set_window_size(1280, 950)

        # Local Windows
        # chrome_options = Options()
        # chrome_options.add_experimental_option("detach", True)
        # chrome_options.add_argument("--timeout=900000")
        # chrome_options.add_argument("--set-script-timeout=900000")
        # chrome_options.add_argument("--disable-webgl")
        # chrome_options.add_argument("--disable-extensions")
        # # chrome_options.add_argument("--start-maximized")
        # chrome_options.add_argument("--disable-gpu")  # Ajuda no Windows
        # chrome_options.add_argument("--window-size=1400,950")
        # chrome_options.add_argument("--no-sandbox")  # Útil para rodar em servidores
        # chrome_options.add_experimental_option("prefs", {
        #     "download.default_directory": get_exe_directory(),
        #     "download.prompt_for_download": False,
        #     "download.directory_upgrade": True,
        #     "safebrowsing.enabled": True
        # })
        # servico = Service(ChromeDriverManager().install())
        # driver = webdriver.Chrome(service=servico, options=chrome_options)
        # driver.set_window_size(1280, 950)

        # VM Windows
        # _options = Options()
        # _options.binary_location = r"D:\90\chrome.exe"
        # _options.add_argument("--timeout=120000")
        # _options.add_argument("--set-script-timeout=60000")
        # _options.add_argument("--disable-webgl")
        # _options.add_argument("--disable-extensions")
        # _options.add_argument("--disable-gpu")  # Ajuda no Windows
        # _options.add_argument("--no-sandbox")  # Útil para rodar em servidores
        # _options.add_experimental_option("prefs", {
        #     "download.default_directory": get_exe_directory(),
        #     "download.prompt_for_download": False,
        #     "download.directory_upgrade": True,
        #     "safebrowsing.enabled": True
        # })
        #
        # driver = webdriver.Chrome(options=_options)
        # driver.set_window_size(1280, 950)
    except Exception as e:
        logger.error(f"FUNÇÃO: iniciar_driver()\nERRO: {e}")
        ## _msg_erro_slack(f"<robo_1_novo_v2> - FUNÇÃO: iniciar_driver()\nERRO: {e}")
        raise

    return driver


def format_date(date_obj):
    if not date_obj:
        return
    try:
        return date_obj.strftime('%d/%m/%Y')
    except:
        try:
            return datetime.strptime(date_obj, '%d/%m/%Y')
        except Exception as e:
            logger.error(f"FUNÇÃO: format_date()\nERRO: {e}")
            ## _msg_erro_slack(f"<robo_1_novo_v2> - FUNÇÃO: format_date()\nERRO: {e}")
            raise


def format_value(value):
    if not value:
        return
    value = value.replace('.', '').replace(',', '.')
    try:
        return float(value)
    except:
        raise


def identificar_campos_login_senha(driver):
    try:
        campo_login = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, 'cpfcnpj')))
        campo_login.click()

        sleep(1)
        campo_login.send_keys(LOGIN)

        sleep(2)

        campo_senha = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, 'senha_home')))
        campo_senha.click()
    except Exception as e:
        logger.error(f"FUNÇÃO: identificar_campos_login_senha()\nERRO: {e}")
        ## _msg_erro_slack(f"<robo_1_novo_v2> - FUNÇÃO: identificar_campos_login_senha()\nERRO: {e}")
        raise


def clicar_shift_act(driver, i):
    try:
        sleep(1)
        tecla = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, '[data-name="shift"]')))
        tecla.click()
        logger.info(f'clicado: {i}')
        sleep(1)

    except Exception as e:
        logger.error(f"FUNÇÃO: clicar_shift_act()\nERRO: {e}")
        ## _msg_erro_slack(f"<robo_1_novo_v2> - FUNÇÃO: clicar_shift_act()\nERRO: {e}")
        raise


def clicar_shift_dct(driver, i):
    try:
        sleep(1)
        tecla = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="senha_home_keyboard"]/div[3]/button[41]')))
        tecla.click()
        logger.info(f'clicado: {i}')
        sleep(1)
    except Exception as e:
        logger.error(f"FUNÇÃO: clicar_shift_dct()\nERRO: {e}")
        ## _msg_erro_slack(f"<robo_1_novo_v2> - FUNÇÃO: clicar_shift_dct()\nERRO: {e}")
        raise


def clicar_letras(driver, i):
    try:
        sleep(1)
        tecla = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, f'[data-name="{i}"]')))
        tecla.click()
        logger.info(f'clicado: {i}')
        sleep(1)
    except Exception as e:
        logger.error(f"FUNÇÃO: clicar_letras()\nERRO: {e}")
        ## _msg_erro_slack(f"<robo_1_novo_v2> - FUNÇÃO: clicar_letras()\nERRO: {e}")
        raise


def botao_login(driver):
    try:
        botao_entrar = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="senha_home_keyboard"]/div[2]/button[53]')))
        botao_entrar.click()
        logger.info('Entrando no portal...')
        sleep(10)
    except Exception as e:
        logger.error(f"FUNÇÃO: botao_login()\nERRO: {e}")
        ## _msg_erro_slack(f"<robo_1_novo_v2> - FUNÇÃO: botao_login()\nERRO: {e}")
        raise


def fazer_login(driver):
    driver.get("https://portal.banco1.com.br/")

    identificar_campos_login_senha(driver)

    sleep(2)
    sequencia_clicks_senha = [
        'shift_act', 'S', 'shift_dct', 'e', 'n', 'h', 'a', '.'
    ]

    try:
        for i in sequencia_clicks_senha:

            if i == 'shift_act':
                clicar_shift_act(driver, i)

            elif i == 'shift_dct':
                clicar_shift_dct(driver, i)

            else:
                clicar_letras(driver, i)

        botao_login(driver)

    except Exception as e:
        logger.error(f"FUNÇÃO: fazer_login()\nERRO: {e}")
        ## _msg_erro_slack(f"<robo_1_novo_v2> - FUNÇÃO: fazer_login()\nERRO: {e}")
        raise e
    return driver


def botao_inicio(driver):
    try:
        logger.info('Inicio consultar info')
        WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="menu"]/div/a[1]')))
        sleep(5)
        logger.info('Selecionando botao inicio')
        inicio = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="menu"]/div/a[1]')))
        inicio.click()
        sleep(6)
    except Exception as e:
        logger.error(f"FUNÇÃO: botao_inicio()\nERRO: {e}")
        ## _msg_erro_slack(f"<robo_1_novo_v2> - FUNÇÃO: botao_inicio()\nERRO: {e}")
        raise


def localizar_tabela(driver):
    try:
        logger.info('Localizando tabela...')
        try:
            tabela = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                (By.CLASS_NAME, 'table-content')
            ))

        except:
            tabela = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                (By.XPATH, '//*[@id="main"]/div[5]/div/div/table/tbody')
            ))
        logger.info('Localizando linhas...')
        lista_linhas = tabela.find_elements(By.TAG_NAME, 'tr')
    except Exception as e:
        logger.error(f"FUNÇÃO: localizar_tabela()\nERRO: {e}")
        ## _msg_erro_slack(f"<robo_1_novo_v2> - FUNÇÃO: localizar_tabela()\nERRO: {e}")
        raise
    return lista_linhas


def quantidade_paginas(driver):
    try:
        try:
            WebDriverWait(driver,10).until(EC.visibility_of_element_located((By.XPATH, './/nav//button[last()-1]')))
            ultima_pagina = driver.find_element(By.XPATH, './/nav//button[last()-1]').text
        except:
            sleep(2)
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, './/nav//button[last()-1]')))
            ultima_pagina = driver.find_element(By.XPATH, './/nav//button[last()-1]').text
    except Exception as e:
        logger.error(f"FUNÇÃO: quantidade_paginas()\nERRO: {e}")
        ## _msg_erro_slack(f"<robo_1_novo_v2> - FUNÇÃO: quantidade_paginas()\nERRO: {e}")
        raise
    return ultima_pagina


def select_identificador_saldo(nome_conta):
    try:
        data = date.strftime(date.today() - timedelta(days=2), "%Y-%m-%d")
        select = f"""select CONCAT(a.cliente, a.dataregistro, a.descricao, a.valor, a.transacao) as identificador from [Digital].[dbo].[t_extratos_1] as a
                                WHERE dataregistro >= '{data}' AND a.cliente = '{nome_conta}' AND a.transacao = 'S'"""

        select_conferencia = _sql_select_valores_sql_pd(select, 'Digital')
        if select_conferencia is None:
            logger.error(f"FUNÇÃO: select_identificador_saldo()\nERRO: no resultado do select")
            ## _msg_erro_slack(f"<robo_1_novo_v2> - FUNÇÃO: select_identificador_saldo()\nERRO: no resultado do select")
            raise
    except Exception as e:
        logger.error(f"FUNÇÃO: select_identificador_saldo()\nERRO: {e}")
        ## _msg_erro_slack(f"<robo_1_novo_v2> - FUNÇÃO: select_identificador_saldo()\nERRO: {e}")
        raise
    return select_conferencia['identificador'].astype(str).tolist()


def coloca_pagina_atual(driver, i, pag):
    try:
        if i == 0:
            return
        elif pag != 1:
            try:
                logger.info('Colocado na pagina certa!')
                clicar_pagina = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
                    (By.XPATH, f".//nav//button[normalize-space(text())='{pag}']")
                ))
                clicar_pagina.click()
                sleep(2)
            except:
                sleep(5)
                logger.info('Colocado na pagina certa! Tentativa 2')
                clicar_pagina = WebDriverWait(driver, 30).until(EC.presence_of_element_located(
                    (By.XPATH, f'//*[@id="main"]/div[5]/div/div/nav/button[{pag+1}]')
                ))
                clicar_pagina.click()
            return
        else:
            return
    except Exception as e:
        logger.error(f"FUNÇÃO: coloca_pagina_atual()\nERRO: {e}")
        ## _msg_erro_slack(f"<robo_1_novo_v2> - FUNÇÃO: coloca_pagina_atual()\nERRO: {e}")
        raise


def clicar_na_linha(driver, linha, j):
    try:
        logger.info('localizando botão selecionar...')
        sleep(1)
        try:
            try:
                botao = linha.find_element(By.XPATH, './/td[last()]//button')
                botao.click()
            except:
                sleep(3)
                botao = WebDriverWait(linha, 10).until(
                EC.element_to_be_clickable(
                    (By.XPATH, f'.//td[last()]//button'))
                )
                botao.click()
        except:
            botao = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, f'//*[@id="main"]/div[5]/div/div/table/tbody/tr[{j + 1}]/td[9]/span/button'))
            )
            botao.click()
    except Exception as e:
        logger.error(f"FUNÇÃO: clicar_na_linha()\nERRO: {e}")
        ## _msg_erro_slack(f"<robo_1_novo_v2> - FUNÇÃO: clicar_na_linha()\nERRO: {e}")
        raise


def nome_empresa(driver):
    try:
        textos = WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.XPATH, "//div[@class='userInfo-ultimoLogin']"))
        )
        texto_completo = textos.get_attribute("innerText")
        nome_conta = texto_completo.split('(')[-1].replace(')', '')
        logger.info(f"Texto encontrado: {nome_conta}")
        return nome_conta
    except Exception as e:
        logger.error(f"FUNÇÃO: nome_empresa()\nERRO: {e}")
        ## _msg_erro_slack(f"<robo_1_novo_v2> - FUNÇÃO: nome_empresa()\nERRO: {e}")
        raise


def saldo(driver):
    try:
        WebDriverWait(driver, 30).until(EC.visibility_of_element_located(
            (By.ID, 'saldo_disponivel')
        ))
        elemento_saldo = WebDriverWait(driver, 30).until(EC.presence_of_element_located(
            (By.ID, 'saldo_disponivel')
        )).text
        saldo = float(elemento_saldo.replace('R$', '').replace('.', '').replace(',', '.').replace(' ', ''))
        return saldo
    except Exception as e:
        logger.error(f"FUNÇÃO: saldo()\nERRO: {e}")
        ## _msg_erro_slack(f"<robo_1_novo_v2> - FUNÇÃO: saldo()\nERRO: {e}")
        raise


def se_dados_for_carregando(driver):
    try:
        driver.refresh()
        sleep(10)
        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '//*[@id="search"]')))
        pesquisar = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="search"]')))
        pesquisar.click()
    except Exception as e:
        logger.error(f"FUNÇÃO: se_dados_for_carregando()\nERRO: {e}")
        ## _msg_erro_slack(f"<robo_1_novo_v2> - FUNÇÃO: se_dados_for_carregando()\nERRO: {e}")
        raise


def botao_pesquisar(driver):
    try:
        botao = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
            (By.ID, 'search')
        ))
        botao.click()
    except Exception as e:
        logger.error(f"FUNÇÃO: botao_pesquisar()\nERRO: {e}")
        ## _msg_erro_slack(f"<robo_1_novo_v2> - FUNÇÃO: botao_pesquisar()\nERRO: {e}")
        raise


def mensagem_slack(mensagem_formatada, saldo_diario, dados_mensagem, headers, saldo_repetido):
    try:
        dataregistrosaldo = str(format_date(date.today()))
        dados_mensagem.append(['----------', '-------------------------', '-------------', '----------'])
        dados_mensagem.append([dataregistrosaldo, 'SALDO DISPONÍVEL', f"R$ {saldo_diario}", 'Saldo'])
        mensagem_formatada += f'```{tabulate(dados_mensagem, headers=headers)}```'

        # _msg_slack('notificacao_contas_escrow', mensagem_formatada)
    except Exception as e:
        logger.error(f"FUNÇÃO: mensagem_slack()\nERRO: {e}")
        ## _msg_erro_slack(f"<robo_1_novo_v2> - FUNÇÃO: mensagem_slack()\nERRO: {e}")
        raise
    return mensagem_formatada


def mensagem_empresas_especificas_slack(nome_conta, dicio_clientes_excecao, mensagem_formatada):
    try:
        if nome_conta in dicio_clientes_excecao.keys():
            canal = dicio_clientes_excecao[str(nome_conta)]
            # _msg_slack(canal=canal, msg=mensagem_formatada)
    except Exception as e:
        logger.error(f"FUNÇÃO: mensagem_empresas_especificas_slack()\nERRO: {e}")
        ## _msg_erro_slack(f"<robo_1_novo_v2> - FUNÇÃO: mensagem_empresas_especificas_slack()\nERRO: {e}")
        raise


def insert_saldo(valor_saldo, nome_conta):
    try:
        logger.info('Tentando inserir dados de saldo na tabela...')
        identificador_saldo = select_identificador_saldo(nome_conta)
        data_atual = date.today().strftime('%Y-%m-%d')
        identificador_atual = nome_conta + data_atual + 'SALDO DISPONIVEL' + f"{valor_saldo:.2f}" + 'S'
        if identificador_atual in identificador_saldo:
            logger.info('Saldo já existente na tabela.')
            return True
        else:
            insert_saldo = f"""
            INSERT INTO t_extratos_1 (banco, cliente, dataregistro, descricao, valor, mes, ano, transacao)
            VALUES ('BANCO 1', '{nome_conta}', '{data_atual}',
                    'SALDO DISPONIVEL', {valor_saldo}, '{meses[int(data_atual[5:7])]}',
                    {int(data_atual[:4])}, 'S')
            """
            _sql_update(insert_saldo, 'Digital')
            logger.info('Dados de saldo inseridos na tabela com sucesso.')
            return False
    except Exception as e:
        logger.error(f"FUNÇÃO: insert_saldo()\nERRO: {e}")
        ## _msg_erro_slack(f"<robo_1_novo_v2> - FUNÇÃO: insert_saldo()\nERRO: {e}")
        raise


def foto_tela_extratos(driver, nome_conta):
    try:

        if os.path.exists(os.path.join('print_extratos', f"{nome_conta} EXTRATO.png")):
            os.remove(os.path.join('print_extratos', f"{nome_conta} EXTRATO.png"))
            sleep(5)
            body = driver.find_element(By.TAG_NAME, value='body')
            body.click()
            body.send_keys(Keys.END)
            driver.execute_script("document.body.style.zoom='80%'")
            sleep(2)
            driver.save_screenshot(os.path.join('print_extratos', f"{nome_conta} EXTRATO.png"))

            driver.execute_script("document.body.style.zoom='100%'")

        else:
            sleep(5)
            body = driver.find_element(By.TAG_NAME, value='body')
            body.click()
            body.send_keys(Keys.END)
            driver.execute_script("document.body.style.zoom='80%'")
            sleep(2)
            driver.save_screenshot(os.path.join('print_extratos', f"{nome_conta} EXTRATO.png"))
            driver.execute_script("document.body.style.zoom='100%'")

        img_path = os.path.join('print_extratos', f"{nome_conta} EXTRATO.png")

        de = _config("email_cobr")
        para = 'email@gmail.com'
        logger.info(f'Enviando email para {para}')

        assunto = f'(1) Movimento na Conta: {nome_conta}'
        mensagem = "Segue em anexo o print do extrato!"
        _envio_email(p_from=de,
                       p_to=para,
                       p_assunto=assunto,
                       p_text=mensagem,
                       p_html="",
                       p_filename=img_path,
                       p_password="")
        logger.info(f'Email com extrato de {nome_conta} enviado')
    except Exception as e:
        logger.error(f"FUNÇÃO: foto_tela_extratos()\nERRO: {e}")
        # _msg_erro_slack(f"<robo_1_novo_v2> - FUNÇÃO: foto_tela_extratos()\nERRO: {e}")
        raise


def botao_10dias(driver):
    try:
        botao = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
            (By.ID, 'select_search')
        ))
        select = Select(botao)

        # Selecionar a opção com value="5"
        select.select_by_value('60')
    except Exception as e:
        logger.error(f"FUNÇÃO: botao_5dias()\nERRO: {e}")
        ## _msg_erro_slack(f"<robo_1_novo_v2> - FUNÇÃO: botao_5dias()\nERRO: {e}")
        raise


def remove_duplicatas():
    delete = """WITH Duplicados AS (
    SELECT
        [cliente],
        [dataregistro],
        [descricao],
        [valor],
        [mes],
        [ano],
        [transacao],
        [historico],
        ROW_NUMBER() OVER (
            PARTITION BY [cliente],
                         [dataregistro],
                         [descricao],
                         [valor],
                         [mes],
                         [ano],
                         [transacao],
                         [historico]
            ORDER BY (SELECT NULL)
        ) AS RN
    FROM [Digital].[dbo].[t_extratos_1]
    WHERE transacao <> 'S'
    )
    DELETE FROM Duplicados
    WHERE RN > 1;"""
    _sql_update(delete, 'Digital')


def t_contas_escrow(driver, nome_conta):
    try:
        nome_conta = str(unidecode(nome_conta))
        num_conta = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
            (By.XPATH, '//*[@id="main"]/div[2]/div[1]/div[2]/div/span')
        )).text
        logger.info(f'Numero conta: {num_conta}')
        identificador_db = _sql_select_valores_sql_pd(
            'SELECT numero_documento as concatenado FROM t_contas_escrow WHERE banco',
            'Digital')
        select = f"""SELECT Id_cedente, cnpj FROM tb_cedente WHERE Nome_Cedente LIKE '{' '.join(nome_conta.split()[0:2])}%'"""
        id_cedente_e_cnpj = _sql_select_valores_sql_pd(select, 'DW')
        if len(id_cedente_e_cnpj) == 0:
            select = f"""SELECT Id_cedente, cnpj FROM tb_cedente WHERE Nome_Cedente LIKE '{' '.join([nome_conta.split()[0], nome_conta.split()[1][0]] if len(nome_conta.split()) > 1 else [nome_conta.split()[0]])}%'"""
            id_cedente_e_cnpj = _sql_select_valores_sql_pd(select, 'DW')
            if len(id_cedente_e_cnpj) == 0:
                select = f"""SELECT Id_cedente, cnpj FROM tb_cedente WHERE Nome_Cedente LIKE '{nome_conta.split()[0]}%'"""
                id_cedente_e_cnpj = _sql_select_valores_sql_pd(select, 'DW')
                if len(id_cedente_e_cnpj) == 0:
                    logger.error(f"FUNÇÃO: t_contas_escrow()\nERRO: Não foi possivel identificar conta: {nome_conta}")
        # print(id_cedente_e_cnpj.iterrows())
        for idx, row in id_cedente_e_cnpj.iterrows():
            id_cedente = row['Id_cedente']
            cnpj = row['cnpj']
            identificador_py = cnpj
            if identificador_py in identificador_db:
                continue
            else:
                insert = ("INSERT INTO t_contas_escrow (nome_conta, numero_conta, numero_documento, ativo, id_cedente)"
                          f"VALUES ('{nome_conta}', '{num_conta}', '{cnpj}', 1, {id_cedente})")
                _sql_update(insert, 'Digital')

                delete_2 = """WITH Duplicados AS (
                        SELECT
                            numero_documento, id_cedente,
                            ROW_NUMBER() OVER (
                                PARTITION BY numero_documento, id_cedente
                                ORDER BY (SELECT NULL)
                            ) AS RN
                        FROM [Digital].[dbo].[t_contas_escrow]
                        )
                        DELETE FROM Duplicados
                        WHERE RN > 1;"""
                _sql_update(delete_2, 'Digital')
    except Exception as e:
        logger.error(f"FUNÇÃO: t_contas_escrow()\nERRO: {e}")
        # _msg_erro_slack(f"<robo_1_novo_v2> - FUNÇÃO: t_contas_escrowl()\nERRO: {e}")
        raise


def get_identificador():
    data_menos_10dias = (datetime.today() - timedelta(days=10)).strftime('%Y-%m-%d %H:%M:%S')
    select = (f"SELECT CONCAT(cliente, dataregistro, descricao,"
              f"valor, mes, ano, transacao, historico, documento,"
              f"conta_corrente, valor_bloqueado, num_movimento,"
              f"cod_hist, saldo, cod_agencia, cpf_contraparte,"
              f"nome_contraparte, cgc_cpf) as nome FROM t_extratos_1 WHERE datainsert >= '{data_menos_10dias}'")
    # conn12 = {"BDIP": "172.17.0.12", "BDLOGIN": "", "BDPASSWORD": "@!@#$%)(*&^"}

    casos_antigos_raw = _sql_select_valores_sql_pd(select, 'Digital')
    casos_antigos = [caso['nome'] for _, caso in casos_antigos_raw.iterrows()]
    return casos_antigos


def csv_diario(registro):
    nome_arquivo = f"1_{date.today().strftime('%d-%m-%Y')}.csv"
    caminho_arquivo = os.path.join("csvs", nome_arquivo)

    # Define as colunas correspondentes à ordem da tupla
    colunas = [
        "banco",
        "cliente",
        "data_registro",
        "descricao",
        "valor",
        "mes",
        "ano",
        "transacao",
        "historico",
        "documento",
        "conta_corrente",
        "valor_bloqueado",
        "num_movimento",
        "cod_hist",
        "saldo",
        "cod_agencia",
        "cpf_contraparte",
        "nome_contraparte",
        "cgc_cpf"
    ]

    # Verifica se o arquivo já existe
    arquivo_existe = os.path.isfile(caminho_arquivo)

    # Abre o arquivo no modo 'append' se existir ou 'write' se não existir
    modo_abertura = 'a' if arquivo_existe else 'w'

    with open(caminho_arquivo, modo_abertura, newline='', encoding='utf-8') as csvfile:
        writer = csv.writer(csvfile)

        # Se o arquivo não existir, escreve o cabeçalho primeiro
        if not arquivo_existe:
            writer.writerow(colunas)

        # Escreve a tupla de registro
        writer.writerow(registro)

    print(f"Registro inserido em: {caminho_arquivo}")


def tabela_xlsx(driver, nome_conta):

    if len(os.listdir('temp')) > 0:
        for i in os.listdir('temp'):
            os.remove(os.path.join('temp', i))
        logger.info('arquivos no diretorio temp limpos')

    botao_download = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
            (By.CLASS_NAME, 'dashboard-control--xlsx ')
        ))
    botao_download.click()
    sleep(7)
    df = pd.read_excel(os.path.join('temp', os.listdir('temp')[0]), skiprows=8, dtype={"Conta Corrente": str,
                                                                            "Código Histórico": str,
                                                                            "Código Agência": str,
                                                                            "CPF Contraparte": str})

    if df.iloc[0,0] == 'Não foram encontrados resultados para essa conta':
        botao_inicio(driver)
        return []

    df = df.astype('object')
    df.where(df.notna(), None, inplace=True)
    df['Saldo'] = df['Saldo'].apply(lambda x: format_value(x))
    df['Valor Bloqueado'] = df['Valor Bloqueado'].apply(lambda x: format_value(x))
    df['Data Movimentação'] = df['Data Movimentação'].apply(lambda x: format_date(x))
    df['Valor'] = df['Valor'].apply(lambda x: format_value(x))


    dados_mensagem = []
    lista_linhas = []
    repetidos = get_identificador()
    for idx, row in df.iterrows():
        # Dados
        conta_corrente =row['Conta Corrente']
        saldo = row['Saldo']
        valor_bloqueado =row['Valor Bloqueado']
        documento = row['Número Documento']
        dataregistro = row['Data Movimentação']
        num_movimento = row['Número do Movimento']
        transacao = row['Natureza']
        cod_hist = row['Código Histórico']
        valor = row['Valor']
        cod_agencia = row['Código Agência']
        descricao = row['Histórico Descrição']
        historico = row['Complemento']
        cgc_cpf = row['CGC CPF']
        cpf_contraparte = row['CPF Contraparte']
        nome_contraparte = row['Nome Contraparte']

        mes = meses.get(int(row['Data Movimentação'].strftime('%m') if row['Data Movimentação'] else None))
        ano = int(row['Data Movimentação'].strftime('%Y') if row['Data Movimentação'].strftime('%Y') else None)

        # Dados mensagem slack
        data_mensagem_raw = dataregistro.strftime('%Y-%m-%d').split('-')
        data_mensagem = f'{data_mensagem_raw[2]}/{data_mensagem_raw[1]}/{data_mensagem_raw[0]}'
        transacao_mensagem = 'Debito' if transacao == 'D' else 'Crédito'
        dados_mensagem.append([data_mensagem, descricao, f"R$ {valor}", transacao_mensagem])

        # Dados insert banco
        linha_insert = ('BANCO 1', nome_conta, dataregistro, descricao,
                        valor, mes, ano, transacao, historico, documento,
                        conta_corrente, valor_bloqueado, num_movimento,
                        cod_hist, saldo, cod_agencia, cpf_contraparte,
                        nome_contraparte, cgc_cpf)
        dado = ''.join([str(item if item is not None else '') for item in (
            nome_conta, dataregistro.strftime('%Y-%m-%d'), descricao, f"{valor:.2f}", mes, ano, transacao,
            historico, documento, conta_corrente, f"{valor_bloqueado:.2f}", num_movimento,
            cod_hist, f"{saldo:.2f}", cod_agencia, cpf_contraparte, nome_contraparte, cgc_cpf
        )])
        if dado not in repetidos:
            # print(f'dado: {dado}')
            # print(f'repetidos: {repetidos}')
            lista_linhas.append(linha_insert)
            csv_diario(linha_insert)

    insert = ("INSERT INTO t_extratos_1_backup (banco, cliente, dataregistro, descricao,"
              "valor, mes, ano, transacao, historico, documento,"
              "conta_corrente, valor_bloqueado, num_movimento,"
              "cod_hist, saldo, cod_agencia, cpf_contraparte,"
              "nome_contraparte, cgc_cpf) VALUES"
              "(%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)")
    print(f'lista_linhas: {lista_linhas}')
    _sql_insert_many(insert, 'Digital', lista_linhas)
    remove_duplicatas()
    botao_inicio(driver)

    for i in os.listdir('temp'):
        os.remove(os.path.join('temp', i))
    logger.info('arquivos no diretorio temp limpos')

    return dados_mensagem


def rpa_1_rafael(driver, headers, dicio_clientes_excecao):

    for i in os.listdir('temp'):
        os.remove(os.path.join('temp', i))

    fazer_login(driver)
    botao_inicio(driver)

    qnt_paginas = quantidade_paginas(driver)
    for i in range(int(qnt_paginas)):   # Itera sobre cada página
        pag = i + 1
        if pag != 1:
            coloca_pagina_atual(driver, i, pag)
            sleep(5)
        lista_linhas = localizar_tabela(driver)
        for j, linha in enumerate(lista_linhas):    # Itera sobre cada linha da página

            logger.info(f'Página - {pag}')
            coloca_pagina_atual(driver, j, pag)
            logger.info(f'Clicando na linha - {j+1}')
            try:
                clicar_na_linha(driver, linha, j)
            except:
                continue
            sleep(3)
            nome_conta = nome_empresa(driver)
            mensagem_formatada = f"*BANCO 1* - {nome_conta}:\n"
            botao_10dias(driver)
            sleep(2)
            botao_pesquisar(driver)
            saldo_diario = saldo(driver)
            foto_tela_extratos(driver, nome_conta)
            dados_mensagem = tabela_xlsx(driver, nome_conta)
            saldo_repetido = insert_saldo(saldo_diario, nome_conta)

            if len(dados_mensagem) == 0 and saldo_repetido:
                logger.info('dados repetidos, ignorando...')
                continue
            try:
                logger.info(
                    f"Mensagem: {mensagem_formatada}, Saldo: {saldo_diario}, dados: {dados_mensagem}, headers: {headers}, saldo_rep: {saldo_repetido}")
                mensagem_final = mensagem_slack(mensagem_formatada, saldo_diario, dados_mensagem, headers, saldo_repetido)
            except Exception as e:
                logger.error(f"FUNÇÃO: rpa_1_rafael()\nERRO: {e}")
                ## _msg_erro_slack(f"<robo_1_novo_v2> - FUNÇÃO: rpa_1_rafael()\nERRO: {e}")
                raise
            logger.info('Mensagens enviadas para o slack')

    sleep(5)
    driver.quit()


if __name__ == '__main__':
    print("RPA 1 ( ͡° ͜ʖ ͡°)")
    # while True:
    hora = datetime.now().hour

    minuto = datetime.now().minute

    dia_semana = date.today().weekday()

    if ((hora in (8, 10, 13, 14, 16, 17, 22) and minuto == 0) or (hora == 15 and minuto == 15) or (hora == 8 and minuto == 50)) and dia_semana < 5:
    # if True:
        driver = iniciar_driver()
        sleep(3)
        try:
            rpa_1_rafael(driver, headers, dicio_clientes_excecao)
            remove_duplicatas()
            print("RPA 1 ( ͡° ͜ʖ ͡°)")
        except Exception as e:
            logger.error(f"ERRO: {e}")
            # _msg_erro_slack(f"<robo_1_novo_v2> - ERRO: {e}")
        finally:
            sleep(5)
            driver.quit()
            del driver