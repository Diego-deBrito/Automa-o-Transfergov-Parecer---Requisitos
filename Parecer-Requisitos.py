import pandas as pd
import time
import os
import json
from selenium import webdriver
from bs4 import BeautifulSoup
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime


# ==============================================================================
# Funções existentes (sem alterações)
# ==============================================================================

def conectar_navegador_existente():
    options = webdriver.ChromeOptions()
    options.debugger_address = "localhost:9222"
    try:
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
        print("✅ Conectado ao navegador existente!")
        return driver
    except Exception as erro:
        print(
            f"❌ Erro ao conectar ao navegador: {erro}. Certifique-se de que o Chrome foi iniciado com o modo de depuração ativado.")
        exit()


def esperar_elemento(driver, xpath, tempo=1):
    try:
        return WebDriverWait(driver, tempo).until(EC.element_to_be_clickable((By.XPATH, xpath)))
    except:
        return None


def esperar_elemento_JSPATH(driver, jspath, tempo=5):
    try:
        return WebDriverWait(driver, tempo).until(EC.element_to_be_clickable((By.CSS_SELECTOR, jspath)))
    except:
        return None


def extrair_data(texto):
    if not texto: return None, None
    formatos_data = ["%d/%m/%Y %H:%M:%S", "%d/%m/%Y", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d"]
    for formato in formatos_data:
        try:
            data = datetime.strptime(texto, formato)
            return data, data.strftime("%d/%m/%Y %H:%M:%S")
        except ValueError:
            continue
    return None, None


def navegar_menu_principal(driver, instrumento):
    try:
        print("🧭 Navegando para a proposta:", instrumento)
        driver.get("https://discricionarias.transferegov.sistema.gov.br/voluntarias/Principal/Principal.do")
        time.sleep(2)  # Espera a página principal carregar

        esperar_elemento(driver, "/html/body/div[1]/div[3]/div[1]/div[1]/div[1]/div[3]", tempo=5).click()
        esperar_elemento(driver, "/html/body/div[1]/div[3]/div[2]/div[1]/div[1]/ul/li[3]/a", tempo=5).click()
        time.sleep(2)

        campo_pesquisa = esperar_elemento(driver,
                                          "/html/body/div[3]/div[15]/div[3]/div/div/form/table/tbody/tr[1]/td[2]/input",
                                          tempo=5)
        if not campo_pesquisa:
            print("⚠️ Campo de pesquisa não encontrado.")
            return False

        campo_pesquisa.clear()
        campo_pesquisa.send_keys(instrumento)
        esperar_elemento(driver,
                         "/html/body/div[3]/div[15]/div[3]/div/div/form/table/tbody/tr[1]/td[2]/span/input").click()
        time.sleep(2)

        link_proposta = esperar_elemento(driver, "/html/body/div[3]/div[15]/div[3]/div[3]/table/tbody/tr/td[1]/div/a",
                                         tempo=3)
        if not link_proposta:
            print(f"⚠️ Proposta '{instrumento}' não encontrada na lista.")
            return False
        link_proposta.click()
        time.sleep(2)

        abas = driver.window_handles
        if len(abas) > 1:
            driver.switch_to.window(abas[-1])

        plano_trabalho = esperar_elemento(driver, "//*[@id='div_997366806']/span/span", tempo=10)
        if not plano_trabalho:
            print("⚠️ Aba 'Plano de Trabalho' não encontrada.")
            return False
        plano_trabalho.click()
        time.sleep(1)

        subaba_pareceres = esperar_elemento(driver, "/html/body/div[3]/div[15]/div[1]/div/div[2]/a[10]/div/span/span",
                                            tempo=10)
        if not subaba_pareceres:
            print("⚠️ Subaba 'Pareceres' não encontrada.")
            return False
        subaba_pareceres.click()
        time.sleep(2)

        print("✅ Navegação concluída!")
        return True
    except Exception as e:
        print(f"❌ Erro ao navegar até proposta: {e}")
        return False


def verificar_pareceres(driver):
    try:
        print("👉 Verificando pareceres...")
        html = driver.page_source
        soup = BeautifulSoup(html, 'html.parser')

        # Verificar se há técnicos conhecidos
        acao = "Técnico Analisar" if not any(
            nome.lower() in html.lower() for nome in NOMES_TECNICOS) else "Nenhuma ação necessária"

        # Extrair a data mais recente de cada tabela, considerando apenas os técnicos listados
        proposta_data = ""
        plano_data = ""
        ajuste_data = ""
        todas_datas = []  # Para armazenar todas as datas para comparação

        # Tabela de Pareceres de Proposta
        try:
            elemento = esperar_elemento(driver, '/html/body/div[3]/div[15]/div[3]/div[2]/table/thead/tr/th[3]', tempo=1)
            if elemento:
                responsaveis = driver.find_elements(By.XPATH,
                                                    '/html/body/div[3]/div[15]/div[3]/div[2]/table/tbody/tr/td[3]')
                datas = driver.find_elements(By.XPATH, '/html/body/div[3]/div[15]/div[3]/div[2]/table/tbody/tr/td[1]')
                if responsaveis and datas and len(responsaveis) == len(datas):
                    lista_datas_validas = []
                    for resp_elem, data_elem in zip(responsaveis, datas):
                        responsavel = resp_elem.text.strip()
                        data_texto = data_elem.text.strip()
                        if any(nome.lower() in responsavel.lower() for nome in NOMES_TECNICOS):
                            data, data_formatada = extrair_data(data_texto)
                            if data:
                                lista_datas_validas.append(data)
                                print(f"✅ Data válida da proposta: {data_formatada} (Responsável: {responsavel})")
                            else:
                                print(f"⚠️ Data inválida: {data_texto}")
                    if lista_datas_validas:
                        data_mais_recente = max(lista_datas_validas)
                        proposta_data = data_mais_recente.strftime('%d/%m/%Y %H:%M:%S')
                        print(f"🕐 Data mais recente de parecer de proposta: {proposta_data}")
                    else:
                        print("⛔ Nenhuma data válida encontrada para responsáveis conhecidos.")
                else:
                    print("⚠️ Inconsistência no número de colunas: Responsáveis vs Datas")
            else:
                print("⛔ Elemento de cabeçalho da coluna Responsável não encontrado.")
        except Exception as e:
            print(f"❗ Erro ao processar pareceres de proposta via Selenium: {e}")
            proposta_data = ""

        # 🔁 Substituindo análise da tabela 'tituloListagem' via Selenium
        try:
            plano_data = ""
            elemento = esperar_elemento(driver,
                                        '/html/body/div[3]/div[15]/div[3]/div[4]/div/div/form/div[1]/table/thead/tr/th[3]',
                                        tempo=1)
            if elemento:
                responsaveis_plano = driver.find_elements(By.XPATH,
                                                          '/html/body/div[3]/div[15]/div[3]/div[4]/div/div/form/div[1]/table/tbody/tr/td[3]')
                datas_plano = driver.find_elements(By.XPATH,
                                                   '/html/body/div[3]/div[15]/div[3]/div[4]/div/div/form/div[1]/table/tbody/tr/td[1]')
                if responsaveis_plano and datas_plano and len(responsaveis_plano) == len(datas_plano):
                    lista_datas_validas_plano = []
                    for resp_elem, data_elem in zip(responsaveis_plano, datas_plano):
                        responsavel = resp_elem.text.strip()
                        data_texto = data_elem.text.strip()
                        if any(nome.lower() in responsavel.lower() for nome in NOMES_TECNICOS):
                            data, data_formatada = extrair_data(data_texto)
                            if data:
                                lista_datas_validas_plano.append(data)
                                print(
                                    f"✅ Data válida do Plano de Trabalho: {data_formatada} (Responsável: {responsavel})")
                            else:
                                print(f"⚠️ Data inválida no Plano de Trabalho: {data_texto}")
                    if lista_datas_validas_plano:
                        plano_data_dt = max(lista_datas_validas_plano)
                        plano_data = plano_data_dt.strftime("%d/%m/%Y %H:%M:%S")
                        print(f"🕐 Data mais recente do Plano de Trabalho: {plano_data}")
                    else:
                        print("⛔ Nenhuma data válida encontrada para responsáveis conhecidos no Plano de Trabalho.")
                else:
                    print("⚠️ Inconsistência no Plano de Trabalho: Responsáveis vs Datas")
            else:
                print("⛔ Elemento da coluna Responsável no Plano de Trabalho não encontrado.")
        except Exception as e:
            print(f"❗ Erro ao extrair dados do Plano de Trabalho via Selenium: {e}")
            plano_data = ""

        # Tabela de Pareceres das Solicitações de Ajuste
        tabela_ajuste = soup.find("table", {"id": "tblParecerAjuste"})
        if tabela_ajuste:
            linhas = tabela_ajuste.find_all("tr")[1:]  # Ignora header
            print(f"📋 Encontradas {len(linhas)} linhas na tabela tblParecerAjuste")
            for linha in linhas:
                colunas = linha.find_all("td")
                if colunas and len(colunas) > 2:
                    responsavel = colunas[0].text.strip()
                    data_str = colunas[1].text.strip()
                    print(f"🔍 Linha: Responsável={responsavel}, Data={data_str}")
                    if any(nome.lower() in responsavel.lower() for nome in NOMES_TECNICOS):
                        data, data_formatada = extrair_data(data_str)
                        if data:
                            if not ajuste_data or (data > datetime.strptime(ajuste_data,
                                                                            "%d/%m/%Y %H:%M:%S" if ":" in ajuste_data else "%d/%m/%Y")):
                                ajuste_data = data_formatada
                            todas_datas.append(data_formatada)
                            print(
                                f"✅ Data válida encontrada para Ajuste: {data_formatada} (Responsável: {responsavel})")
                        else:
                            print(f"⚠️ Data inválida: {data_str}")
                else:
                    print(f"⚠️ Linha inválida: {colunas}")
        else:
            print("⚠️ Tabela tblParecerAjuste não encontrada")

        # Definir strings para o Excel
        proposta_str = proposta_data if proposta_data else "Nenhum parecer localizado"
        plano_str = plano_data if plano_data else "Nenhum parecer localizado"
        ajuste_str = ajuste_data if ajuste_data else "Nenhum parecer localizado"

        # Exibir no console os dados a serem inseridos
        print(f"📝 Dados a serem inseridos: Lista Pareceres de Proposta={proposta_str}, "
              f"Lista Pareceres do Plano de Trabalho={plano_str}, "
              f"Lista Pareceres das Solicitações de Ajuste={ajuste_str}")

        # ✅ Substituindo o cálculo de data mais recente por Selenium puro + funções existentes
        try:
            data_mais_recente = None
            # Consulta na primeira coluna
            elemento = esperar_elemento(driver, '/html/body/div[3]/div[15]/div[3]/div[2]/table/thead/tr/th[1]', tempo=1)
            if elemento:
                datas_coluna_1 = driver.find_elements(By.XPATH,
                                                      '/html/body/div[3]/div[15]/div[3]/div[2]/table/tbody/tr/td[1]')
                lista_datas_1 = []
                for data_element in datas_coluna_1:
                    texto = data_element.text.strip()
                    if texto:
                        data, _ = extrair_data(texto)
                        if data:
                            lista_datas_1.append(data)
                if lista_datas_1:
                    data_mais_recente = max(lista_datas_1)
                    print(f"🕐 Data mais recente na primeira coluna: {data_mais_recente.strftime('%d/%m/%Y %H:%M:%S')}")
                else:
                    print("⛔ Nenhuma data válida encontrada na primeira coluna.")
            else:
                print("⛔ Elemento da primeira coluna não encontrado.")

            # Consulta na segunda coluna (se houver dados, sobrescreve)
            elemento = esperar_elemento(driver,
                                        '/html/body/div[3]/div[15]/div[3]/div[4]/div/div/form/div[1]/table/thead/tr/th[1]',
                                        tempo=1)
            if elemento:
                datas_coluna_2 = driver.find_elements(By.XPATH,
                                                      '/html/body/div[3]/div[15]/div[3]/div[4]/div/div/form/div[1]/table/tbody/tr/td[1]')
                lista_datas_2 = []
                for data_element in datas_coluna_2:
                    texto = data_element.text.strip()
                    if texto:
                        data, _ = extrair_data(texto)
                        if data:
                            lista_datas_2.append(data)
                if lista_datas_2:
                    data_coluna_2 = max(lista_datas_2)
                    if not data_mais_recente or data_coluna_2 > data_mais_recente:
                        data_mais_recente = data_coluna_2
                    print(f"🕐 Data mais recente na segunda coluna: {data_coluna_2.strftime('%d/%m/%Y %H:%M:%S')}")
                else:
                    print("⛔ Nenhuma data válida encontrada na segunda coluna.")
            else:
                print("⛔ Elemento da segunda coluna não encontrado.")
        except Exception as e:
            print(f"❗ Erro ao buscar as datas mais recentes via Selenium: {e}")
            data_mais_recente = None

        return acao, proposta_str, plano_str, ajuste_str, data_mais_recente

    except Exception as e:
        print(f"⚠️ Erro ao verificar pareceres: {e}")
        return "Erro ao verificar", "Erro ao extrair", "Erro ao extrair", "Erro ao extrair", None


def buscar_data_mais_recente(driver, xpath, descricao="", descripcion=None):
    try:
        elemento = WebDriverWait(driver, 2).until(
            EC.presence_of_element_located((By.XPATH, xpath))
        )
        texto = elemento.text.strip()
        data, _ = extrair_data(texto)
        if data:
            data_formatada = data.strftime("%d/%m/%Y %H:%M:%S")
            print(f"✅ [{descricao}] Data extraída: {data_formatada}")
            return data_formatada
        else:
            print(f"⚠️ [{descricao}] Data inválida no texto: {texto}")
            return ""
    except Exception as e:
        print(f"⚠️ [{descripcion}] XPath não encontrado ou erro ao buscar: {e}")
        return ""


def buscar_status(driver, xpath, descricao="", descripcion=None):
    try:
        elemento = WebDriverWait(driver, 2).until(
            EC.presence_of_element_located((By.XPATH, xpath))
        )
        texto = elemento.text.strip()
        print(f"✅ [{descricao}] Status extraído: {texto}")
        return texto
    except Exception as e:
        print(f"⚠️ [{descripcion}] Status não encontrado ou erro ao buscar: {e}")
        return ""


def verificar_requisitos(driver, data_pareceres, proposta_str, plano_str, ajuste_str):
    try:
        print("👉 Verificando requisitos (XPath + Selenium)...")

        # 🧭 Passo 1: Acessar aba Requisitos
        requisitos_menu = esperar_elemento(driver, "/html/body/div[3]/div[15]/div[1]/div/div[1]/a[3]/div/span/span",
                                           tempo=10)
        if requisitos_menu:
            requisitos_menu.click()
        else:
            print("❌ Não foi possível acessar a aba 'Requisitos'")
            return "Erro ao acessar aba Requisitos", "", "", "", "", "", ""

        time.sleep(1)
        # 🧭 Passo 2: Acessar subaba Requisitos
        subaba_requisitos = esperar_elemento_JSPATH(driver,
                                                    "a[id='menu_link_2144784112_100344749'] div[class='inactiveTab'] span span",
                                                    tempo=10)
        if subaba_requisitos:
            subaba_requisitos.click()
        else:
            print("❌ Não foi possível acessar a subaba 'Requisitos'")
            return "Erro ao acessar subaba Requisitos", "", "", "", "", "", ""

        time.sleep(2)

        # ✅ Coletar os dados diretamente com XPaths
        certidoes_data = buscar_data_mais_recente(driver,
                                                  "/html/body/div[3]/div[16]/div[2]/div[2]/form/div[1]/div[1]/table/tbody/tr[1]/td[2]",
                                                  "Certidões")

        declaracoes_data = buscar_data_mais_recente(driver,
                                                    "/html/body/div[3]/div[16]/div[2]/div[2]/form/div[1]/div[2]/table/tbody/tr[1]/td[2]",
                                                    "Declarações")

        comprovantes_data = buscar_data_mais_recente(driver,
                                                     "/html/body/div[3]/div[16]/div[2]/div[2]/form/div[1]/div[3]/table/tbody/tr[1]/td[2]",
                                                     "Comprovantes de Execução")

        outros_data = buscar_data_mais_recente(driver,
                                               "/html/body/div[3]/div[16]/div[2]/div[2]/form/div[1]/div[4]/table/tbody/tr[1]/td[2]",
                                               "Outros")

        historico_data = buscar_data_mais_recente(driver,
                                                  "/html/body/div[3]/div[16]/div[2]/div[2]/form/div[1]/div[5]/table/tbody/tr[1]/td[3]",
                                                  "Histórico (Data)")

        historico_evento = buscar_status(driver,
                                         "/html/body/div[3]/div[16]/div[2]/div[2]/form/div[1]/div[5]/table/tbody/tr[1]/td[1]",
                                         "Histórico (Evento)")

        # ✅ Verificar se consta algum documento inserido (Certidões, Declarações, Comprovantes ou Outros com datas válidas)
        documentos_encontrados = any([
            certidoes_data and certidoes_data != "Nenhuma certidão localizada",
            declaracoes_data and declaracoes_data != "Nenhuma declaração localizada",
            comprovantes_data and comprovantes_data != "Nenhum comprovante localizado",
            outros_data and outros_data != "Nenhum outro documento localizado"
        ])

        # ✅ Verificar se consta alguma diligência nos pareceres
        diligencia_nos_pareceres = any([
            "diligência" in (proposta_str or "").lower(),
            "diligência" in (plano_str or "").lower(),
            "diligência" in (ajuste_str or "").lower()
        ])

        # 🎯 Aplicar as novas regras para definir a ação necessária
        acao = None
        if historico_evento.strip() == "Análise Registrada - Não Atendido":
            acao = "Nenhuma Ação Necessária"
        elif historico_evento.strip() == "Análise Registrada - Atendido":
            acao = "Parceria Celebrada"
        elif not documentos_encontrados:
            # Regra: Se nenhuma data for encontrada na aba Requisitos (exceto Histórico),
            # a ação é "Entidade Pendência de Documentação"
            acao = "Entidade Pendência de Documentação"
        elif (
                (("diligência" in (proposta_str or "").lower()) or ("diligência" in (plano_str or "").lower()))
                and documentos_encontrados
        ):
            acao = "Técnico Analisar"
        elif not diligencia_nos_pareceres and documentos_encontrados:
            acao = "Técnico Analisar"
        elif not diligencia_nos_pareceres and not documentos_encontrados:
            acao = "Técnico Inserir 1ª Diligência"
        else:
            # Regra: Se a única data válida for de Histórico (Data), a ação deve ser
            # "Entidade Pendência de Documentação" em vez de "Técnico Analisar"
            acao = "Entidade Pendência de Documentação"

        # ✅ Ajustar a lógica para sobrescrever "Técnico Analisar" com base na Regra
        if acao == "Técnico Analisar" and historico_data and not documentos_encontrados:
            acao = "Entidade Pendência de Documentação"

        # ✅ Preencher os campos com mensagens padrão caso estejam vazios
        certidoes_str = certidoes_data if certidoes_data else "Nenhuma certidão localizada"
        declaracoes_str = declaracoes_data if declaracoes_data else "Nenhuma declaração localizada"
        comprovantes_str = comprovantes_data if comprovantes_data else "Nenhum comprovante localizado"
        outros_str = outros_data if outros_data else "Nenhum outro documento localizado"

        print(f"📝 Dados extraídos: Certidões={certidoes_str}, Declarações={declaracoes_str}, "
              f"Comprovantes={comprovantes_str}, Outros={outros_str}, "
              f"Histórico (Data)={historico_data}, Histórico (Evento)={historico_evento}, Ação={acao}")

        return acao, certidoes_str, declaracoes_str, comprovantes_str, outros_str, historico_data, historico_evento

    except Exception as e:
        print(f"❗ Erro ao verificar requisitos via XPath: {e}")
        return "Erro ao verificar", "", "", "", "", "", ""


def ler_checkpoint(caminho_checkpoint):
    if os.path.exists(caminho_checkpoint):
        try:
            with open(caminho_checkpoint, 'r') as f:
                return json.load(f).get('ultimo_indice', 0)
        except (json.JSONDecodeError, IOError):
            return 0
    return 0


def salvar_checkpoint(caminho_checkpoint, indice):
    try:
        with open(caminho_checkpoint, 'w') as f:
            json.dump({'ultimo_indice': indice}, f)
        print(f"✅ Checkpoint salvo: próximo início no índice {indice}")
    except IOError:
        print("⚠️ Erro ao salvar checkpoint.")


def salvar_resultado(df, caminho_saida):
    df.to_excel(caminho_saida, index=False)
    print(f"📁 Resultado salvo em: {caminho_saida}")


NOMES_TECNICOS = [
    "Aline da Silva Mesquita Santos",
    "Josely Pereira de Aquino de Lima",
    "Julio Cesar Santos de Paula",
    "Fabricia de Morais Ramos",
    "Loane Arcebispo Avilino",
    "Samara Nogueira de Souza",
    "João Victor Cavalcante Teodoro"
]


# ==============================================================================
# Funções modificadas
# ==============================================================================

def ler_entrada_excel(caminho):
    try:
        # Carregar a planilha, especificando a aba
        df = pd.read_excel(caminho, sheet_name='Propostas 2025', dtype=str)
        print(f"📄 Planilha lida. Contagem inicial de linhas (brutas): {len(df)}")

        # Verificar se as colunas obrigatórias existem
        colunas_obrigatorias = ['Nº Proposta', 'Instrumento']
        for col in colunas_obrigatorias:
            if col not in df.columns:
                raise ValueError(f"Coluna '{col}' não encontrada na planilha!")

        # Filtrar linhas com 'Nº Proposta' não nulo, não vazio e remover duplicatas
        df = df[df['Nº Proposta'].notna() & (df['Nº Proposta'].str.strip() != '')].drop_duplicates(
            subset=['Nº Proposta']).reset_index(drop=True)
        print(f"📄 Após filtro de 'Nº Proposta' e remoção de duplicatas: {len(df)} linhas")

        # Garantir que a coluna 'Instrumento' não tenha valores nulos
        df['Instrumento'] = df['Instrumento'].fillna('').str.strip()
        print(f"📄 Coluna 'Instrumento' preenchida e limpa.")

        # Criar DataFrame de saída com as colunas na ordem especificada
        colunas_saida = [
            'Nº Proposta',
            'Instrumento',
            'Situacional',  # Adicionada a coluna Situacional
            'Técnico Responsável pela Formalização',
            'Lista Pareceres de Proposta',
            'Lista Pareceres do plano de Trabalho',
            'Lista Pareceres das Solicitações de Ajuste do Plano de Trabalho',
            'Certidões',
            'Declarações',
            'Comprovantes de Execução',
            'Outros',
            'Histórico (Data)',
            'Histórico (Evento)',
            'Projeto Básico / Termo de Referência (Somente Convênio)',
            'Ação Necessária'
        ]
        df_saida = pd.DataFrame(columns=colunas_saida)

        # Copiar as colunas de entrada para a saída, se existirem
        for col in ['Nº Proposta', 'Instrumento', 'Situacional', 'Técnico Responsável pela Formalização']:
            if col in df.columns:
                df_saida[col] = df[col]
            else:
                print(f"⚠️ Coluna {col} não encontrada na planilha. Preenchendo com vazio.")
                df_saida[col] = ''

        # Inicializar colunas criadas como vazias para as linhas restantes
        for col in colunas_saida:
            if col not in df_saida.columns:
                df_saida[col] = ''

        # Garantir que o número de linhas em df_saida seja igual ao de df
        df_saida = df_saida.reindex(range(len(df))).reset_index(drop=True)
        print(f"📄 DataFrame de saída criado com {len(df_saida)} linhas")

        return df_saida

    except Exception as e:
        print(f"❌ Erro ao ler planilha: {e}")
        raise


def preencher_linha(df, indice, driver, instrumento):
    """
    Função auxiliar que contém a lógica de processamento para uma única linha (proposta).
    Reutilizada pelo processamento completo e pelo reprocessamento.
    """
    # ================= INÍCIO DAS NOVAS REGRAS =================
    # Listas de situacionais para cada ação, conforme as novas regras
    situacionais_em_celebracao = [
        "Parecer e Termo para Correção",
        "Parecer e Termo Corrigido (Necessidade de Ajuste)",
        "Parecer para Assinatura",
        "Termo Disponibilizado Usuário Externo",
        "Termo Disponibilizado Secretário(a)",
        "Processo Enviado ao GAB - Para Publicação"
    ]
    situacionais_parceria_celebrada = [
        "Aguardando Registro Transferegov CGFP",
        "Cláusula Suspensiva 2025",
        "Enviado à CGAP"
    ]
    situacionais_proposta_rejeitada = [
        "Alteração de Beneficiário",
        "Proposta Rejeitada"
    ]
    situacionais_tecnico_analisar = [
        "Enviar Link Declarações",
        "Enviado Link Declarações"
    ]
    # Lista de situações que FORÇAM a extração de dados (lógica original)
    situacionais_continuar_extracao = [
        "Requisitos para Celebração Atendido",
        "Pendência Requisitos para Celebração"
    ]

    # Obter o valor de Situacional para a linha atual
    situacional = str(df.loc[indice, "Situacional"]).strip() if 'Situacional' in df.columns else ""

    # Inicializar variáveis de controle
    acao_final = None
    executar_extracao_web = True  # Flag para decidir se a navegação web é necessária

    # Aplicar as novas regras baseadas no 'Situacional'
    if situacional in situacionais_em_celebracao:
        acao_final = "Em Celebração"
        executar_extracao_web = False
    elif situacional in situacionais_parceria_celebrada:
        acao_final = "Parceria Celebrada"
        executar_extracao_web = False
    elif situacional in situacionais_proposta_rejeitada:
        acao_final = "Proposta Rejeitada"
        executar_extracao_web = False
    elif situacional in situacionais_tecnico_analisar:
        acao_final = "Técnico Analisar"
        executar_extracao_web = False
    elif situacional in situacionais_continuar_extracao:
        executar_extracao_web = True  # Força a extração para estes casos
    # Se 'situacional' não estiver em nenhuma lista, a extração ocorrerá por padrão

    # ================= FIM DAS NOVAS REGRAS =================

    # A lógica de extração só será executada se necessário
    if not executar_extracao_web:
        print(f"👉 Ação definida diretamente pelo Situacional '{situacional}': {acao_final}")
        # Preenche a linha com a ação definida e marca o resto como não aplicável
        df.loc[indice, "Ação Necessária"] = acao_final
        df.loc[indice, "Lista Pareceres de Proposta"] = "N/A (Definido pelo Situacional)"
        df.loc[indice, "Lista Pareceres do plano de Trabalho"] = "N/A (Definido pelo Situacional)"
        df.loc[
            indice, "Lista Pareceres das Solicitações de Ajuste do Plano de Trabalho"] = "N/A (Definido pelo Situacional)"
        df.loc[indice, "Certidões"] = "N/A (Definido pelo Situacional)"
        df.loc[indice, "Declarações"] = "N/A (Definido pelo Situacional)"
        df.loc[indice, "Comprovantes de Execução"] = "N/A (Definido pelo Situacional)"
        df.loc[indice, "Outros"] = "N/A (Definido pelo Situacional)"
        df.loc[indice, "Histórico (Data)"] = "N/A (Definido pelo Situacional)"
        df.loc[indice, "Histórico (Evento)"] = "N/A (Definido pelo Situacional)"

    else:
        # Se a extração for necessária, executa o fluxo original
        print(f"👉 Situacional '{situacional}' requer análise. Iniciando extração de dados...")
        if not navegar_menu_principal(driver, instrumento):
            # Preenche a linha com informações de falha na navegação
            df.loc[indice, "Ação Necessária"] = "Instrumento não encontrado"
            df.loc[indice, "Lista Pareceres de Proposta"] = "N/A"
            df.loc[indice, "Lista Pareceres do plano de Trabalho"] = "N/A"
            df.loc[indice, "Lista Pareceres das Solicitações de Ajuste do Plano de Trabalho"] = "N/A"
            df.loc[indice, "Certidões"] = "N/A"
            df.loc[indice, "Declarações"] = "N/A"
            df.loc[indice, "Comprovantes de Execução"] = "N/A"
            df.loc[indice, "Outros"] = "N/A"
            df.loc[indice, "Histórico (Data)"] = "N/A"
            df.loc[indice, "Histórico (Evento)"] = "N/A"
        else:
            # Lógica de extração de dados se a navegação for bem-sucedida
            _, proposta_str, plano_str, ajuste_str, data_pareceres = verificar_pareceres(driver)

            acao_requisitos, certidoes_str, declaracoes_str, comprovantes_str, outros_str, historico_data, historico_evento = verificar_requisitos(
                driver, data_pareceres, proposta_str, plano_str, ajuste_str
            )

            # Converter as strings de data para objetos datetime para comparação
            data_proposta, _ = extrair_data(
                proposta_str) if proposta_str and proposta_str != "Nenhum parecer localizado" else (None, None)
            data_plano, _ = extrair_data(plano_str) if plano_str and plano_str != "Nenhum parecer localizado" else (
            None, None)
            data_certidoes, _ = extrair_data(
                certidoes_str) if certidoes_str and certidoes_str != "Nenhuma certidão localizada" else (None, None)
            data_declaracoes, _ = extrair_data(
                declaracoes_str) if declaracoes_str and declaracoes_str != "Nenhuma declaração localizada" else (
            None, None)
            data_comprovantes, _ = extrair_data(
                comprovantes_str) if comprovantes_str and comprovantes_str != "Nenhum comprovante localizado" else (
            None, None)
            data_outros, _ = extrair_data(
                outros_str) if outros_str and outros_str != "Nenhum outro documento localizado" else (None, None)
            data_historico, _ = extrair_data(historico_data) if historico_data and historico_data != "" else (
            None, None)

            # Determinar a data mais recente de pareceres
            datas_pareceres = [data_proposta, data_plano]
            datas_pareceres_validas = [d for d in datas_pareceres if d is not None]
            data_pareceres_mais_recente = max(datas_pareceres_validas) if datas_pareceres_validas else None

            # Determinar a data mais recente de requisitos
            datas_requisitos = [data_certidoes, data_declaracoes, data_comprovantes, data_outros]
            datas_requisitos_validas = [d for d in datas_requisitos if d is not None]
            data_requisitos_mais_recente = max(datas_requisitos_validas) if datas_requisitos_validas else None

            # Aplicar as regras de data para determinar a Ação Necessária
            if not datas_pareceres_validas and not datas_requisitos_validas:
                acao_final = "Entidade Pendência de Documentação"
            elif not datas_pareceres_validas and datas_requisitos_validas:
                acao_final = "Técnico Analisar"
            elif datas_pareceres_validas and not datas_requisitos_validas:
                acao_final = "Entidade Pendência de Documentação"
            elif data_pareceres_mais_recente and data_requisitos_mais_recente:
                if data_pareceres_mais_recente > data_requisitos_mais_recente:
                    acao_final = "Entidade Pendência de Documentação"
                elif data_requisitos_mais_recente > data_pareceres_mais_recente:
                    acao_final = "Técnico Analisar"
                else:
                    if historico_evento.strip() == "Complementação Solicitada":
                        acao_final = "Entidade Pendência de Documentação"
                    elif historico_evento.strip() == "Enviado para Verificação":
                        acao_final = "Técnico Analisar"
                    else:
                        acao_final = acao_requisitos
            else:
                acao_final = acao_requisitos

            # Preenche o DataFrame com os dados extraídos
            df.loc[indice, "Lista Pareceres de Proposta"] = proposta_str
            df.loc[indice, "Lista Pareceres do plano de Trabalho"] = plano_str
            df.loc[indice, "Lista Pareceres das Solicitações de Ajuste do Plano de Trabalho"] = ajuste_str
            df.loc[indice, "Certidões"] = certidoes_str
            df.loc[indice, "Declarações"] = declaracoes_str
            df.loc[indice, "Comprovantes de Execução"] = comprovantes_str
            df.loc[indice, "Outros"] = outros_str
            df.loc[indice, "Histórico (Data)"] = historico_data
            df.loc[indice, "Histórico (Evento)"] = historico_evento
            df.loc[indice, "Ação Necessária"] = acao_final

    return df


def rodar_processamento_completo(driver, caminho_entrada, caminho_saida, caminho_checkpoint, filtro_instrumento):
    """
    Executa o script do início, lendo a planilha de entrada e processando apenas as propostas que atendem ao filtro.
    """
    print(f"\n--- MODO DE PROCESSAMENTO COMPLETO (Filtro: {filtro_instrumento}) ---")

    try:
        df = ler_entrada_excel(caminho_entrada)
    except Exception:
        print("❌ Encerrando o script devido a erro na leitura do arquivo de entrada.")
        return

    # Aplicar o filtro na coluna 'Instrumento', se não for 'Todos'
    if filtro_instrumento.lower() != 'todos':
        df = df[df['Instrumento'].str.lower() == filtro_instrumento.lower()].reset_index(drop=True)
        print(f"📄 Após aplicar filtro '{filtro_instrumento}' na coluna 'Instrumento': {len(df)} linhas")

    total_propostas = len(df)
    if total_propostas == 0:
        print(f"⚠️ Nenhuma proposta encontrada para o filtro '{filtro_instrumento}'. Encerrando.")
        return

    ultimo_indice = ler_checkpoint(caminho_checkpoint)
    print(f"🛠️ Iniciando a partir do índice {ultimo_indice} (checkpoint)")

    for indice, linha in df.iterrows():
        if indice < ultimo_indice:
            continue

        print(f"\n📦 Processando proposta {indice + 1}/{total_propostas} (Instrumento: {linha['Instrumento']})")

        instrumento = str(linha.get("Nº Proposta", "")).strip()
        if not instrumento:
            print("⚠️ Nº da proposta não encontrado na linha de entrada. Pulando.")
            continue

        # Chama a função que contém a lógica de processamento
        df = preencher_linha(df, indice, driver, instrumento)

        # Salva o resultado e o checkpoint
        salvar_resultado(df, caminho_saida)
        salvar_checkpoint(caminho_checkpoint, indice + 1)

    print("\n✅ Processamento completo de todas as propostas filtradas concluído!")

    if os.path.exists(caminho_checkpoint):
        os.remove(caminho_checkpoint)
        print("🗑️ Arquivo de checkpoint removido.")


def reprocessar_falhas(driver, caminho_saida, filtro_instrumento):
    """
    Verifica a planilha de saída, encontra linhas em branco ('Ação Necessária' vazia)
    e reprocessa apenas essas linhas que atendem ao filtro de Instrumento.
    """
    print(f"\n--- MODO DE REPROCESSAMENTO DE FALHAS (Filtro: {filtro_instrumento}) ---")

    if not os.path.exists(caminho_saida):
        print(f"⚠️ Arquivo de saída '{caminho_saida}' não encontrado. Não há o que reprocessar.")
        print("Execute o processamento completo primeiro.")
        return

    df = pd.read_excel(caminho_saida, dtype=str).fillna('')

    # Aplicar o filtro na coluna 'Instrumento', se não for 'Todos'
    if filtro_instrumento.lower() != 'todos':
        df = df[df['Instrumento'].str.lower() == filtro_instrumento.lower()].reset_index(drop=True)
        print(f"📄 Após aplicar filtro '{filtro_instrumento}' na coluna 'Instrumento': {len(df)} linhas")

    # Identifica as linhas onde a 'Ação Necessária' está em branco
    linhas_para_reprocessar = df[df['Ação Necessária'].str.strip() == '']

    if linhas_para_reprocessar.empty:
        print("✅ Nenhuma falha encontrada para o filtro especificado. A planilha de saída está completamente preenchida.")
        return

    total_falhas = len(linhas_para_reprocessar)
    print(f"🔍 {total_falhas} propostas com falha encontradas. Iniciando reprocessamento...")

    contador = 0
    for indice, linha in linhas_para_reprocessar.iterrows():
        contador += 1
        print(f"\n📦 Reprocessando falha {contador}/{total_falhas} (Linha {indice + 2} do Excel)")

        instrumento = str(linha.get("Nº Proposta", "")).strip()
        if not instrumento:
            print("⚠️ Nº da proposta não encontrado nesta linha. Pulando.")
            continue

        # Chama a função que contém a lógica de processamento
        df = preencher_linha(df, indice, driver, instrumento)

        # Salva o resultado imediatamente após processar a linha
        salvar_resultado(df, caminho_saida)

    print("\n✅ Reprocessamento de falhas concluído!")


def main():
    """
    Função principal que oferece ao usuário a escolha do modo de execução e o filtro de Instrumento.
    """
    # 📥 Caminhos de entrada, saída e checkpoint
    caminho_entrada = r'C:\Users\diego.brito\Downloads\robov1\output\Situacional Propostas SNEAELIS.xlsm'
    caminho_saida = r'C:\Users\diego.brito\Downloads\robov1\output\resultado_pr.xlsx'
    caminho_checkpoint = r'C:\Users\diego.brito\Downloads\robov1\output\checkpoint.json'

    while True:
        print("\n=====================================")
        print("  ROBÔ DE ANÁLISE DE PROPOSTAS SNEA")
        print("=====================================")
        print("Escolha o modo de execução:")
        print("[1] Rodar processamento completo (a partir do zero ou de um checkpoint)")
        print("[2] Reprocessar apenas as falhas (preencher linhas em branco na saída)")
        print("[3] Sair")

        escolha = input("Digite sua escolha (1, 2 ou 3): ").strip()

        if escolha in ['1', '2']:
            # Solicitar o filtro para a coluna 'Instrumento'
            print("\n📌 Informe o filtro para a coluna 'Instrumento' (ex.: 'Convênio', 'Contrato de Repasse') ou 'Todos' para processar tudo:")
            filtro_instrumento = input("Filtro: ").strip()
            if not filtro_instrumento:
                filtro_instrumento = 'Todos'  # Valor padrão se o usuário não informar nada
            print(f"✅ Filtro selecionado: {filtro_instrumento}")

            driver = conectar_navegador_existente()
            if escolha == '1':
                rodar_processamento_completo(driver, caminho_entrada, caminho_saida, caminho_checkpoint, filtro_instrumento)
            elif escolha == '2':
                reprocessar_falhas(driver, caminho_saida, filtro_instrumento)

            print("👋 Encerrando a conexão com o navegador.")
            driver.quit()
            break  # Sai do loop while após a conclusão

        elif escolha == '3':
            print("👋 Encerrando o programa.")
            break

        else:
            print("❌ Escolha inválida. Por favor, digite 1, 2 ou 3.")


if __name__ == "__main__":
    main()