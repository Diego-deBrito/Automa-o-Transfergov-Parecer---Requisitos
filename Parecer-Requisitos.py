# -*- coding: utf-8 -*-
"""
Este script é um robô de automação (RPA) projetado para interagir com o portal Transferegov.
Seu objetivo principal é extrair o status de propostas governamentais, analisando
documentos, pareceres e requisitos para determinar a "Ação Necessária" em cada caso.

Funcionalidades:
- Conecta-se a uma sessão do Google Chrome já em execução no modo de depuração.
- Navega pela plataforma Transferegov para encontrar propostas específicas.
- Extrai dados de datas e responsáveis das seções de "Pareceres" e "Requisitos".
- Aplica um conjunto de regras de negócio para determinar o status processual.
- Preenche uma planilha Excel com os dados extraídos e a ação recomendada.
- Oferece modos de execução para processamento completo, reprocessamento de falhas
  e correção de dados na planilha de saída.

Pré-requisitos:
- Python 3.x
- Bibliotecas: pandas, selenium, beautifulsoup4, webdriver-manager
- Google Chrome instalado.
- Uma instância do Google Chrome deve ser iniciada em modo de depuração na porta 9222.
  Exemplo de comando (Windows):
  "C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="C:\ChromeDebug"
"""

import pandas as pd
import time
import os
import json
import unicodedata
from selenium import webdriver
from bs4 import BeautifulSoup
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime

# --- Configuração Global ---
# Lista de nomes de técnicos para filtrar os pareceres relevantes.
# A análise de datas de pareceres será focada apenas nos emitidos por estes profissionais.
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
# Funções Utilitárias e de Interação com o Navegador
# ==============================================================================

def conectar_navegador_existente():
    """
    Conecta-se a uma instância do Google Chrome em execução com a porta de depuração remota ativada.

    Returns:
        webdriver.Chrome: Objeto do driver do Selenium conectado ao navegador.
        None: Se a conexão falhar, o script é encerrado.
    """
    options = webdriver.ChromeOptions()
    options.debugger_address = "localhost:9222"
    try:
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
        print("[SUCCESS] Conectado com sucesso ao navegador existente.")
        return driver
    except Exception as erro:
        print(f"[ERROR] Falha ao conectar ao navegador: {erro}.")
        print("[INFO] Verifique se o Chrome foi iniciado com o modo de depuração ativado na porta 9222.")
        exit()


def esperar_elemento(driver, xpath, tempo=1):
    """
    Aguarda explicitamente até que um elemento seja clicável na página.

    Args:
        driver (webdriver.Chrome): A instância do driver do Selenium.
        xpath (str): O seletor XPath do elemento a ser aguardado.
        tempo (int): O tempo máximo de espera em segundos.

    Returns:
        WebElement: O elemento encontrado.
        None: Se o elemento não for encontrado dentro do tempo limite.
    """
    try:
        return WebDriverWait(driver, tempo).until(EC.element_to_be_clickable((By.XPATH, xpath)))
    except:
        return None


def esperar_elemento_JSPATH(driver, jspath, tempo=5):
    """
    Aguarda explicitamente até que um elemento (via seletor CSS) seja clicável.

    Args:
        driver (webdriver.Chrome): A instância do driver do Selenium.
        jspath (str): O seletor CSS do elemento.
        tempo (int): O tempo máximo de espera em segundos.

    Returns:
        WebElement: O elemento encontrado.
        None: Se o elemento não for encontrado.
    """
    try:
        return WebDriverWait(driver, tempo).until(EC.element_to_be_clickable((By.CSS_SELECTOR, jspath)))
    except:
        return None


def extrair_data(texto_data):
    """
    Converte uma string de data em um objeto datetime, testando múltiplos formatos.

    Args:
        texto_data (str): A string contendo a data (ex: "dd/mm/aaaa HH:MM:SS").

    Returns:
        tuple: Uma tupla contendo (objeto datetime, string de data formatada) ou (None, None) se a conversão falhar.
    """
    if not texto_data:
        return None, None
    
    formatos_suportados = ["%d/%m/%Y %H:%M:%S", "%d/%m/%Y", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d"]
    for formato in formatos_suportados:
        try:
            data_obj = datetime.strptime(texto_data, formato)
            return data_obj, data_obj.strftime("%d/%m/%Y %H:%M:%S")
        except ValueError:
            continue
    return None, None


# ==============================================================================
# Funções de Lógica de Negócio e Web Scraping
# ==============================================================================

def navegar_menu_principal(driver, num_proposta):
    """
    Navega através do menu do Transferegov para encontrar e abrir uma proposta específica.

    Args:
        driver (webdriver.Chrome): A instância do driver do Selenium.
        num_proposta (str): O número da proposta a ser pesquisada.

    Returns:
        bool: True se a navegação até a página de pareceres for bem-sucedida, False caso contrário.
    """
    try:
        print(f"[INFO] Navegando para a proposta: {num_proposta}")
        driver.get("https://discricionarias.transferegov.sistema.gov.br/voluntarias/Principal/Principal.do")
        time.sleep(2)  # Aguarda estática para carregamento inicial da página.

        # Sequência de cliques para navegar pelos menus
        if not (elemento_menu := esperar_elemento(driver, "/html/body/div[1]/div[3]/div[1]/div[1]/div[1]/div[3]", tempo=5)):
            print("[WARNING] Elemento do menu principal não encontrado.")
            return False
        elemento_menu.click()

        if not (elemento_submenu := esperar_elemento(driver, "/html/body/div[1]/div[3]/div[2]/div[1]/div[1]/ul/li[3]/a", tempo=5)):
            print("[WARNING] Elemento do submenu não encontrado.")
            return False
        elemento_submenu.click()
        time.sleep(2)

        # Pesquisa pela proposta
        if not (campo_pesquisa := esperar_elemento(driver, "/html/body/div[3]/div[15]/div[3]/div/div/form/table/tbody/tr[1]/td[2]/input", tempo=5)):
            print("[WARNING] Campo de pesquisa não encontrado.")
            return False
        campo_pesquisa.clear()
        campo_pesquisa.send_keys(num_proposta)

        if not (botao_pesquisar := esperar_elemento(driver, "/html/body/div[3]/div[15]/div[3]/div/div/form/table/tbody/tr[1]/td[2]/span/input")):
            print("[WARNING] Botão de pesquisa não encontrado.")
            return False
        botao_pesquisar.click()
        time.sleep(2)

        # Abre a proposta encontrada e muda o foco para a nova aba
        if not (link_proposta := esperar_elemento(driver, "/html/body/div[3]/div[15]/div[3]/div[3]/table/tbody/tr/td[1]/div/a", tempo=3)):
            print(f"[WARNING] Proposta '{num_proposta}' não encontrada na lista de resultados.")
            return False
        link_proposta.click()
        time.sleep(2)

        if len(driver.window_handles) > 1:
            driver.switch_to.window(driver.window_handles[-1])

        # Navega para a aba de pareceres
        if not (plano_trabalho := esperar_elemento(driver, "//*[@id='div_997366806']/span/span", tempo=10)):
            print("[WARNING] Aba 'Plano de Trabalho' não encontrada.")
            return False
        plano_trabalho.click()
        time.sleep(1)

        if not (subaba_pareceres := esperar_elemento(driver, "/html/body/div[3]/div[15]/div[1]/div/div[2]/a[10]/div/span/span", tempo=10)):
            print("[WARNING] Sub-aba 'Pareceres' não encontrada.")
            return False
        subaba_pareceres.click()
        time.sleep(2)

        print("[SUCCESS] Navegação até a página de pareceres concluída.")
        return True
    except Exception as e:
        print(f"[ERROR] Ocorreu um erro inesperado durante a navegação: {e}")
        return False


def verificar_pareceres(driver):
    """
    Extrai as datas dos pareceres da proposta e do plano de trabalho,
    considerando apenas os técnicos listados em NOMES_TECNICOS.

    Args:
        driver (webdriver.Chrome): A instância do driver do Selenium.

    Returns:
        tuple: Contém (ação, data_proposta, data_plano, data_ajuste, data_mais_recente_geral).
               Retorna strings de erro em caso de falha.
    """
    try:
        print("[INFO] Iniciando verificação de pareceres...")
        html = driver.page_source
        soup = BeautifulSoup(html, 'html.parser')
        
        proposta_data = ""
        plano_data = ""
        datas_pareceres_validas = []

        # --- Extração de Pareceres da Proposta ---
        try:
            tabela_proposta = soup.find("div", {"id": "divPareceresProposta"})
            if tabela_proposta:
                linhas = tabela_proposta.find("tbody").find_all("tr")
                for linha in linhas:
                    celulas = linha.find_all("td")
                    if len(celulas) >= 3:
                        data_texto = celulas[0].text.strip()
                        responsavel = celulas[2].text.strip()
                        
                        if any(nome.lower() in responsavel.lower() for nome in NOMES_TECNICOS):
                            data_obj, data_formatada = extrair_data(data_texto)
                            if data_obj:
                                datas_pareceres_validas.append(data_obj)
                                print(f"  [PARECER PROPOSTA] Data encontrada: {data_formatada} (Responsável: {responsavel})")
        except Exception as e:
            print(f"[WARNING] Não foi possível extrair dados de pareceres da proposta: {e}")

        # --- Extração de Pareceres do Plano de Trabalho ---
        try:
            tabela_plano = soup.find("div", {"id": "divPareceresPlanoTrabalho"})
            if tabela_plano:
                linhas = tabela_plano.find("tbody").find_all("tr")
                for linha in linhas:
                    celulas = linha.find_all("td")
                    if len(celulas) >= 3:
                        data_texto = celulas[0].text.strip()
                        responsavel = celulas[2].text.strip()
                        
                        if any(nome.lower() in responsavel.lower() for nome in NOMES_TECNICOS):
                            data_obj, data_formatada = extrair_data(data_texto)
                            if data_obj:
                                datas_pareceres_validas.append(data_obj)
                                print(f"  [PARECER PLANO] Data encontrada: {data_formatada} (Responsável: {responsavel})")
        except Exception as e:
            print(f"[WARNING] Não foi possível extrair dados de pareceres do plano de trabalho: {e}")

        # Determina a data mais recente entre todos os pareceres válidos
        data_mais_recente_obj = None
        if datas_pareceres_validas:
            data_mais_recente_obj = max(datas_pareceres_validas)
            proposta_data = data_mais_recente_obj.strftime('%d/%m/%Y %H:%M:%S')
            print(f"[INFO] Data mais recente de parecer (Proposta/Plano): {proposta_data}")

        proposta_str = proposta_data if proposta_data else "Nenhum parecer localizado"
        
        # O campo 'plano_str' e 'ajuste_str' podem ser preenchidos se necessário,
        # mas a lógica atual consolida tudo em 'proposta_str'
        plano_str = "Nenhum parecer localizado" 
        ajuste_str = "Nenhum parecer localizado"

        return "Ação a Definir", proposta_str, plano_str, ajuste_str, data_mais_recente_obj

    except Exception as e:
        print(f"[ERROR] Falha crítica ao verificar pareceres: {e}")
        return "Erro ao verificar", "Erro ao extrair", "Erro ao extrair", "Erro ao extrair", None


def verificar_requisitos(driver):
    """
    Navega para a aba de 'Requisitos', extrai as datas de envio de documentos
    e o status do histórico mais recente.

    Args:
        driver (webdriver.Chrome): A instância do driver do Selenium.

    Returns:
        tuple: Contém os dados extraídos (certidoes, declaracoes, etc.) e a ação derivada.
               Retorna strings de erro em caso de falha.
    """
    try:
        print("[INFO] Iniciando verificação de requisitos...")
        
        # --- Navegação para a aba e sub-aba de Requisitos ---
        if not (requisitos_menu := esperar_elemento(driver, "/html/body/div[3]/div[15]/div[1]/div/div[1]/a[3]/div/span/span", tempo=10)):
            print("[ERROR] Não foi possível acessar a aba 'Requisitos'.")
            return "Erro de Navegação", "", "", "", "", "", ""
        requisitos_menu.click()
        time.sleep(1)

        if not (subaba_requisitos := esperar_elemento_JSPATH(driver, "a[id='menu_link_2144784112_100344749'] div[class='inactiveTab'] span span", tempo=10)):
            print("[ERROR] Não foi possível acessar a sub-aba 'Requisitos'.")
            return "Erro de Navegação", "", "", "", "", "", ""
        subaba_requisitos.click()
        time.sleep(2)

        # --- Lógica de Extração de Datas dos Documentos ---
        palavras_ignorar = [
            "atestado", "capacidade", "técnica", "projeto", "técnico", 
            "pedagógico", "ptp", "plano", "trabalho", "planilha", "custos"
        ]

        def extrair_data_valida_categoria(xpath_tabela, categoria_nome):
            """Função aninhada para extrair a data mais recente de uma tabela de documentos."""
            try:
                linhas = driver.find_elements(By.XPATH, xpath_tabela + "/tbody/tr")
                datas_validas = []
                for linha in linhas:
                    tds = linha.find_elements(By.TAG_NAME, "td")
                    if len(tds) >= 2:
                        nome_arquivo = tds[0].text.strip().lower()
                        data_arquivo_str = tds[1].text.strip()
                        
                        # Ignora documentos que não são de conformidade (ex: planos de trabalho)
                        if not any(palavra in nome_arquivo for palavra in palavras_ignorar):
                            if (data_dt := extrair_data(data_arquivo_str)[0]):
                                datas_validas.append(data_dt)

                if datas_validas:
                    data_recente = max(datas_validas)
                    data_formatada = data_recente.strftime('%d/%m/%Y %H:%M:%S')
                    print(f"  [{categoria_nome}] Data mais recente encontrada: {data_formatada}")
                    return data_recente, data_formatada
                return None, f"Nenhum(a) {categoria_nome.lower()} localizado(a)"
            except Exception:
                return None, f"Erro ao extrair {categoria_nome.lower()}"

        # Extrai data para cada categoria de documento
        data_certidoes_obj, certidoes_str = extrair_data_valida_categoria(
            "/html/body/div[3]/div[16]/div[2]/div[2]/form/div[1]/div[1]/table", "Certidões")
        data_declaracoes_obj, declaracoes_str = extrair_data_valida_categoria(
            "/html/body/div[3]/div[16]/div[2]/div[2]/form/div[1]/div[2]/table", "Declarações")
        data_comprovantes_obj, comprovantes_str = extrair_data_valida_categoria(
            "/html/body/div[3]/div[16]/div[2]/div[2]/form/div[1]/div[3]/table", "Comprovantes de Execução")
        data_outros_obj, outros_str = extrair_data_valida_categoria(
            "/html/body/div[3]/div[16]/div[2]/div[2]/form/div[1]/div[4]/table", "Outros")

        # --- Extração do Histórico ---
        historico_data = ""
        historico_evento = ""
        try:
            historico_evento = driver.find_element(By.XPATH, "/html/body/div[3]/div[16]/div[2]/div[2]/form/div[1]/div[5]/table/tbody/tr[1]/td[1]").text.strip()
            historico_data_str = driver.find_element(By.XPATH, "/html/body/div[3]/div[16]/div[2]/div[2]/form/div[1]/div[5]/table/tbody/tr[1]/td[3]").text.strip()
            _, historico_data = extrair_data(historico_data_str)
            print(f"  [HISTÓRICO] Evento: '{historico_evento}', Data: {historico_data}")
        except Exception:
            print("[WARNING] Não foi possível extrair dados do histórico.")
        
        # Consolida todas as datas de requisitos para encontrar a mais recente
        datas_requisitos_validas = [d for d in [data_certidoes_obj, data_declaracoes_obj, data_comprovantes_obj, data_outros_obj] if d]
        data_mais_recente_requisitos = max(datas_requisitos_validas) if datas_requisitos_validas else None
        
        return data_mais_recente_requisitos, certidoes_str, declaracoes_str, comprovantes_str, outros_str, historico_data, historico_evento

    except Exception as e:
        print(f"[ERROR] Falha crítica ao verificar requisitos: {e}")
        return None, "Erro", "Erro", "Erro", "Erro", "Erro", "Erro"

# ==============================================================================
# Funções de Gerenciamento de Arquivos e Processamento em Lote
# ==============================================================================

def ler_checkpoint(caminho):
    """Lê o último índice processado de um arquivo de checkpoint."""
    if os.path.exists(caminho):
        try:
            with open(caminho, 'r') as f:
                return json.load(f).get('ultimo_indice', 0)
        except (json.JSONDecodeError, IOError):
            return 0
    return 0


def salvar_checkpoint(caminho, indice):
    """Salva o índice da próxima proposta a ser processada."""
    try:
        with open(caminho, 'w') as f:
            json.dump({'ultimo_indice': indice}, f)
        print(f"[INFO] Checkpoint salvo. Próximo início no índice {indice}.")
    except IOError:
        print("[WARNING] Não foi possível salvar o checkpoint.")


def salvar_resultado(df, caminho_saida):
    """Salva o DataFrame final em um arquivo Excel."""
    try:
        df.to_excel(caminho_saida, index=False)
        print(f"[SUCCESS] Resultado salvo com sucesso em: {caminho_saida}")
    except Exception as e:
        print(f"[ERROR] Falha ao salvar o arquivo de resultado: {e}")


def ler_entrada_excel(caminho):
    """
    Lê e prepara a planilha de entrada, filtrando e estruturando os dados.

    Args:
        caminho (str): O caminho para o arquivo Excel de entrada.

    Returns:
        pd.DataFrame: DataFrame preparado para o processamento.
    """
    try:
        df = pd.read_excel(caminho, sheet_name='Propostas 2025', dtype=str)
        print(f"[INFO] Planilha lida. Total de linhas: {len(df)}.")

        colunas_obrigatorias = ['Nº Proposta', 'Instrumento']
        if not all(col in df.columns for col in colunas_obrigatorias):
            raise ValueError("As colunas 'Nº Proposta' e 'Instrumento' são obrigatórias.")

        # Limpeza e preparação dos dados
        df.dropna(subset=['Nº Proposta'], inplace=True)
        df = df[df['Nº Proposta'].str.strip() != ''].drop_duplicates(subset=['Nº Proposta']).reset_index(drop=True)
        print(f"[INFO] Após limpeza e remoção de duplicatas: {len(df)} linhas.")

        # Estrutura do DataFrame de saída
        colunas_saida = [
            'Nº Proposta', 'Instrumento', 'Técnico Responsável pela Formalização',
            'Lista Pareceres de Proposta', 'Lista Pareceres do plano de Trabalho',
            'Certidões', 'Declarações', 'Comprovantes de Execução', 'Outros',
            'Histórico (Data)', 'Histórico (Evento)',
            'Fator Pendente para Celebração (Exceto Documentação)',
            'Situacional (Documentação)', 'Ação Necessária (Automação)'
        ]
        
        # Cria um novo DataFrame com a estrutura correta, preservando dados existentes
        df_saida = pd.DataFrame(columns=colunas_saida)
        for col in df_saida.columns:
            if col in df.columns:
                df_saida[col] = df[col]
            else:
                df_saida[col] = ''
        
        return df_saida.fillna('')

    except FileNotFoundError:
        print(f"[ERROR] Arquivo de entrada não encontrado em: {caminho}")
        raise
    except Exception as e:
        print(f"[ERROR] Erro ao ler e preparar a planilha de entrada: {e}")
        raise


def preencher_linha(df, indice, driver, num_proposta):
    """
    Orquestra o processo de extração e análise para uma única proposta (linha do DataFrame).

    Args:
        df (pd.DataFrame): O DataFrame principal.
        indice (int): O índice da linha a ser processada.
        driver (webdriver.Chrome): A instância do driver do Selenium.
        num_proposta (str): O número da proposta a ser analisada.

    Returns:
        pd.DataFrame: O DataFrame atualizado com os resultados da linha.
    """
    # --- Regras de Negócio baseadas no "Situacional" (pré-análise) ---
    situacionais_em_celebracao = [
        "Parecer e Termo para Correção", "Parecer e Termo Corrigido (Necessidade de Ajuste)",
        "Parecer para Assinatura", "Termo Disponibilizado Usuário Externo",
        "Termo Disponibilizado Secretário(a)", "Processo Enviado ao GAB - Para Publicação"
    ]
    situacionais_parceria_celebrada = [
        "Aguardando Registro Transferegov CGFP", "Cláusula Suspensiva 2025", "Enviado à CGAP"
    ]
    situacionais_proposta_rejeitada = ["Alteração de Beneficiário", "Proposta Rejeitada"]
    situacionais_tecnico_analisar = ["Enviar Link Declarações", "Enviado Link Declarações"]
    
    situacional = df.loc[indice, "Situacional (Documentação)"].strip()
    acao_final = None

    if situacional in situacionais_em_celebracao:
        acao_final = "Em Celebração"
    elif situacional in situacionais_parceria_celebrada:
        acao_final = "Parceria Celebrada"
    elif situacional in situacionais_proposta_rejeitada:
        acao_final = "Proposta Rejeitada"
    elif situacional in situacionais_tecnico_analisar:
        acao_final = "Técnico Analisar"

    # Se uma ação foi definida pela pré-análise, pula a extração web
    if acao_final:
        print(f"[INFO] Ação definida como '{acao_final}' com base no situacional '{situacional}'. Pulando extração web.")
        df.loc[indice, "Ação Necessária (Automação)"] = acao_final
        for col in ['Lista Pareceres de Proposta', 'Certidões', 'Declarações', 'Histórico (Data)']:
             df.loc[indice, col] = "N/A (Definido pelo Situacional)"
        return df

    # --- Processo de Extração Web ---
    if not navegar_menu_principal(driver, num_proposta):
        df.loc[indice, "Ação Necessária (Automação)"] = "Instrumento não encontrado"
        return df

    _, proposta_str, plano_str, _, data_pareceres = verificar_pareceres(driver)
    
    data_requisitos, certidoes_str, declaracoes_str, comprovantes_str, outros_str, historico_data, historico_evento = verificar_requisitos(driver)

    # --- Lógica de Decisão da Ação Necessária ---
    # Compara a data do último parecer do técnico com a data do último documento enviado pela entidade.
    if data_pareceres and data_requisitos:
        if data_pareceres > data_requisitos:
            acao_final = "Entidade Pendência de Documentação"
        elif data_requisitos > data_pareceres:
            acao_final = "Técnico Analisar"
        else: # Datas são iguais, desempate pelo evento do histórico
            if "Complementação Solicitada" in historico_evento:
                acao_final = "Entidade Pendência de Documentação"
            else:
                acao_final = "Técnico Analisar"
    elif data_requisitos and not data_pareceres:
        acao_final = "Técnico Analisar"
    elif data_pareceres and not data_requisitos:
        acao_final = "Entidade Pendência de Documentação"
    else: # Nenhum documento ou parecer encontrado
        acao_final = "Entidade Pendência de Documentação"
        
    # Preenche o DataFrame com os resultados
    df.loc[indice, "Lista Pareceres de Proposta"] = proposta_str
    df.loc[indice, "Lista Pareceres do plano de Trabalho"] = plano_str
    df.loc[indice, "Certidões"] = certidoes_str
    df.loc[indice, "Declarações"] = declaracoes_str
    df.loc[indice, "Comprovantes de Execução"] = comprovantes_str
    df.loc[indice, "Outros"] = outros_str
    df.loc[indice, "Histórico (Data)"] = historico_data
    df.loc[indice, "Histórico (Evento)"] = historico_evento
    df.loc[indice, "Ação Necessária (Automação)"] = acao_final

    return df


def rodar_processamento_completo(driver, paths, filtro_instrumento):
    """
    Executa o script completo, lendo a planilha de entrada e processando as propostas filtradas.
    """
    print(f"\n--- MODO: PROCESSAMENTO COMPLETO | FILTRO: {filtro_instrumento} ---")

    try:
        df = ler_entrada_excel(paths['entrada'])
    except Exception:
        print("[ERROR] Encerrando script devido a erro na leitura do arquivo de entrada.")
        return

    if filtro_instrumento.lower() != 'todos':
        df = df[df['Instrumento'].str.lower() == filtro_instrumento.lower()].reset_index(drop=True)
        print(f"[INFO] Após filtro '{filtro_instrumento}': {len(df)} linhas para processar.")

    if df.empty:
        print("[INFO] Nenhuma proposta encontrada para o filtro especificado. Encerrando.")
        return

    ultimo_indice = ler_checkpoint(paths['checkpoint'])
    print(f"[INFO] Iniciando a partir do índice {ultimo_indice} (via checkpoint).")

    tempos = []
    inicio_total = time.time()
    total_propostas = len(df)

    for idx, linha in df.iterrows():
        if idx < ultimo_indice:
            continue

        inicio_proposta = time.time()
        print(f"\n[INFO] Processando proposta {idx + 1}/{total_propostas} | Nº: {linha['Nº Proposta']}")
        
        num_proposta = str(linha.get("Nº Proposta", "")).strip()
        if not num_proposta:
            print("[WARNING] Nº da proposta não encontrado nesta linha. Pulando.")
            continue

        df = preencher_linha(df, idx, driver, num_proposta)

        # Cálculo de tempo e estimativa
        tempo_gasto = time.time() - inicio_proposta
        tempos.append(tempo_gasto)
        media_tempo = sum(tempos) / len(tempos)
        restantes = total_propostas - (idx + 1)
        estimado = (restantes * media_tempo) / 60
        print(f"[TIMER] Tempo da proposta: {tempo_gasto:.2f}s | Média: {media_tempo:.2f}s | Estimativa restante: {estimado:.1f} min")

        salvar_checkpoint(paths['checkpoint'], idx + 1)
        # Salva o resultado a cada X iterações para segurança
        if (idx + 1) % 10 == 0:
            salvar_resultado(df, paths['saida'])


    salvar_resultado(df, paths['saida'])
    tempo_total = (time.time() - inicio_total) / 60
    print(f"\n[SUCCESS] Processamento completo concluído! Tempo total: {tempo_total:.1f} min")

    if os.path.exists(paths['checkpoint']):
        os.remove(paths['checkpoint'])
        print("[INFO] Arquivo de checkpoint removido após conclusão.")


def reprocessar_falhas(driver, paths, filtro_instrumento):
    """
    Verifica a planilha de saída, encontra linhas com falhas e as reprocessa.
    """
    print(f"\n--- MODO: REPROCESSAMENTO DE FALHAS | FILTRO: {filtro_instrumento} ---")

    if not os.path.exists(paths['saida']):
        print(f"[WARNING] Arquivo de saída '{paths['saida']}' não encontrado. Execute o processamento completo primeiro.")
        return

    df = pd.read_excel(paths['saida'], dtype=str).fillna('')

    if filtro_instrumento.lower() != 'todos':
        df = df[df['Instrumento'].str.lower() == filtro_instrumento.lower()]

    # Identifica falhas (células vazias em 'Ação Necessária')
    indices_para_reprocessar = df[df['Ação Necessária (Automação)'].str.strip() == ''].index

    if indices_para_reprocessar.empty:
        print("[SUCCESS] Nenhuma falha encontrada para o filtro especificado.")
        return

    total_falhas = len(indices_para_reprocessar)
    print(f"[INFO] {total_falhas} propostas com falha encontradas. Iniciando reprocessamento...")
    
    for count, idx in enumerate(indices_para_reprocessar):
        linha = df.loc[idx]
        print(f"\n[INFO] Reprocessando falha {count + 1}/{total_falhas} | Linha do Excel: {idx + 2}")
        
        num_proposta = str(linha.get("Nº Proposta", "")).strip()
        if not num_proposta:
            print("[WARNING] Nº da proposta não encontrado. Pulando.")
            continue
            
        df = preencher_linha(df, idx, driver, num_proposta)
        salvar_resultado(df, paths['saida']) # Salva após cada reprocessamento

    print("\n[SUCCESS] Reprocessamento de falhas concluído!")

# ==============================================================================
# Função Principal e Interface de Usuário
# ==============================================================================

def main():
    """
    Função principal que gerencia o fluxo de execução e a interação com o usuário.
    """
    # Defina os caminhos para os arquivos de entrada, saída e checkpoint.
    # Recomenda-se usar caminhos relativos ou variáveis de ambiente para portabilidade.
    base_path = r'C:\Users\diego.brito\Downloads\robov1\output'
    paths = {
        'entrada': os.path.join(base_path, 'Situacional Propostas SNEAELIS.xlsm'),
        'saida': os.path.join(base_path, 'resultado_analise_propostas.xlsx'),
        'checkpoint': os.path.join(base_path, 'checkpoint.json')
    }

    while True:
        print("\n" + "="*50)
        print("    ROBÔ DE ANÁLISE DE PROPOSTAS - TRANSFEREGOV")
        print("="*50)
        print("Escolha o modo de execução:")
        print("[1] Rodar processamento completo (do início ou checkpoint)")
        print("[2] Reprocessar apenas as falhas")
        print("[3] Sair")

        escolha = input("Digite sua escolha (1, 2 ou 3): ").strip()

        if escolha in ['1', '2']:
            print("\nInforme o tipo de instrumento para filtrar (ex: 'Convênio')")
            filtro_instrumento = input("Ou digite 'Todos' para processar todos: ").strip()
            if not filtro_instrumento:
                filtro_instrumento = 'Todos'
            print(f"[INFO] Filtro selecionado: {filtro_instrumento}")
            
            driver = conectar_navegador_existente()
            if escolha == '1':
                rodar_processamento_completo(driver, paths, filtro_instrumento)
            elif escolha == '2':
                reprocessar_falhas(driver, paths, filtro_instrumento)
            
            print("[INFO] Processamento finalizado. O navegador permanecerá aberto.")

        elif escolha == '3':
            print("[INFO] Encerrando o programa.")
            break
        else:
            print("[WARNING] Escolha inválida. Por favor, digite 1, 2 ou 3.")

if __name__ == "__main__":
    main()
