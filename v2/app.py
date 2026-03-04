import time
import sys
import os
import subprocess
import urllib.request
import pyautogui

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from config import empresas_alvo, MES, ANO, NOME_RELATORIO, PASTA_DOWNLOAD, URL_ESPERADA_CCEE, ELEMENTOS_VALIDACAO


# =============================================================================
# INICIALIZAÇÃO
# =============================================================================

def iniciar_chrome_debug():
    """Abre o Chrome com porta de debug. Substitui o iniciar_robo.bat."""
    caminho_chrome = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
    pasta_perfil = r"C:\selenium\chrome_perfil"

    try:
        urllib.request.urlopen("http://127.0.0.1:9222/json", timeout=2)
        print("[INFO] Chrome ja esta rodando na porta 9222.")
        return
    except:
        pass

    print("[INFO] Iniciando Chrome com porta de debug...")
    subprocess.Popen([
        caminho_chrome,
        "--remote-debugging-port=9222",
        f"--user-data-dir={pasta_perfil}",
        "https://operacao.ccee.org.br/"
    ])
    time.sleep(4)
    print("[INFO] Faca o login na CCEE e selecione o relatorio.")
    input("[ENTER] Pressione Enter quando estiver pronto para o robo comecar...")


def conectar_chrome_aberto():
    """Conecta ao Chrome ja aberto na porta 9222."""
    print("Conectando ao Chrome (Porta 9222)...")
    options = Options()
    options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
    try:
        driver = webdriver.Chrome(options=options)
        return driver
    except Exception as e:
        print(f"[ERRO] Chrome nao encontrado. Execute o script novamente.\nErro: {e}")
        sys.exit()


def configurar_pasta_download():
    """Garante que a pasta de download existe."""
    if not os.path.exists(PASTA_DOWNLOAD):
        print(f"[AVISO] Pasta nao existe, criando: {PASTA_DOWNLOAD}")
        try:
            os.makedirs(PASTA_DOWNLOAD)
        except Exception as e:
            print(f"[ERRO] Nao foi possivel criar a pasta: {e}")
            sys.exit()
    return PASTA_DOWNLOAD


# =============================================================================
# VALIDACOES
# =============================================================================

def validar_url_ccee(driver):
    """Verifica se esta no site correto da CCEE."""
    url_atual = driver.current_url.lower()
    for url_valida in URL_ESPERADA_CCEE:
        if url_valida.lower() in url_atual:
            return True
    print(f"[ALERTA] URL nao e da CCEE: {url_atual}")
    return False


def validar_conteudo_carregado(driver):
    """Verifica se ha conteudo valido na tela antes de tentar exportar PDF."""
    print("  -> Validando conteudo da pagina...")

    mensagens_erro = [
        "//div[contains(text(), 'Nenhum resultado')]",
        "//div[contains(text(), 'Sem dados')]",
        "//div[contains(text(), 'erro')]",
        "//*[contains(@class, 'ErrorMessage')]"
    ]

    driver.switch_to.default_content()
    for xpath_erro in mensagens_erro:
        try:
            if driver.find_elements(By.XPATH, xpath_erro):
                print("  [AVISO] Mensagem de erro/sem dados detectada!")
                return False
        except:
            pass

    elementos_encontrados = 0

    driver.switch_to.default_content()
    for xpath_validacao in ELEMENTOS_VALIDACAO:
        try:
            if driver.find_elements(By.XPATH, xpath_validacao):
                elementos_encontrados += 1
        except:
            pass

    iframes = driver.find_elements(By.TAG_NAME, "iframe")
    for frame in iframes:
        try:
            driver.switch_to.default_content()
            driver.switch_to.frame(frame)
            for xpath_validacao in ELEMENTOS_VALIDACAO:
                try:
                    if driver.find_elements(By.XPATH, xpath_validacao):
                        elementos_encontrados += 1
                        break
                except:
                    pass
        except:
            continue

    driver.switch_to.default_content()

    if elementos_encontrados > 0:
        print(f"  [OK] Conteudo validado ({elementos_encontrados} elementos encontrados).")
        return True
    else:
        print("  [ERRO] Nenhum conteudo de relatorio detectado!")
        return False


# =============================================================================
# NAVEGACAO E ELEMENTOS
# =============================================================================

def buscar_elemento_inteligente(driver, xpath, tempo=5):
    """Procura elemento na pagina e dentro de iframes."""
    wait = WebDriverWait(driver, tempo)

    driver.switch_to.default_content()
    try:
        return wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
    except:
        pass

    iframes = driver.find_elements(By.TAG_NAME, "iframe")
    for frame in iframes:
        try:
            driver.switch_to.default_content()
            driver.switch_to.frame(frame)
            return wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
        except:
            continue

    driver.switch_to.default_content()
    raise Exception(f"Elemento nao encontrado: {xpath}")


def buscar_submenu_pdf(driver, tempo_max=10):
    """Procura o botao PDF no submenu em todos os contextos possiveis."""
    xpaths_pdf = [
        "//td[contains(@class, 'MenuItemTextCell') and contains(text(), 'PDF')]",
        "//td[contains(text(), 'Pagina Atual como PDF')]",
        "//a[contains(@class, 'MenuItem') and contains(., 'PDF')]",
        "//*[contains(@id, 'PDF') or contains(@id, 'pdf')]//td[contains(@class, 'MenuItemTextCell')]",
        "//td[@class='MenuItemTextCell' and normalize-space(text())='PDF']"
    ]

    inicio = time.time()
    while time.time() - inicio < tempo_max:
        driver.switch_to.default_content()
        for xpath in xpaths_pdf:
            try:
                elementos = driver.find_elements(By.XPATH, xpath)
                for el in elementos:
                    if el.is_displayed():
                        return el
            except:
                continue

        iframes = driver.find_elements(By.TAG_NAME, "iframe")
        for frame in iframes:
            try:
                driver.switch_to.default_content()
                driver.switch_to.frame(frame)
                for xpath in xpaths_pdf:
                    try:
                        elementos = driver.find_elements(By.XPATH, xpath)
                        for el in elementos:
                            if el.is_displayed():
                                return el
                    except:
                        continue
            except:
                continue

        time.sleep(0.3)

    driver.switch_to.default_content()
    return None


def clicar_js(driver, elemento):
    driver.execute_script("arguments[0].click();", elemento)


def disparar_submenu_imprimir(driver, el_imprimir):
    """Forca a abertura do submenu de impressao via multiplas estrategias."""
    print("    -> Disparando evento onmouseover via JS...")
    try:
        driver.execute_script("""
            var element = arguments[0];
            var event = new MouseEvent('mouseover', {
                'view': window,
                'bubbles': true,
                'cancelable': true
            });
            element.dispatchEvent(event);
        """, el_imprimir)
        time.sleep(1)
    except Exception as e:
        print(f"    [AVISO] Erro ao disparar mouseover: {e}")

    print("    -> Executando funcao saw.dashboard.DisplayLayouts...")
    try:
        onmouseover_attr = el_imprimir.get_attribute("onmouseover")
        if onmouseover_attr and "saw.dashboard.DisplayLayouts" in onmouseover_attr:
            driver.execute_script(f"""
                var element = arguments[0];
                var event = new MouseEvent('mouseover', {{
                    'view': window,
                    'bubbles': true,
                    'cancelable': true
                }});
                element.onmouseover.call(element, event);
            """, el_imprimir)
            time.sleep(1.5)
            print("    [OK] Funcao executada.")
        else:
            print("    [AVISO] Atributo onmouseover nao encontrado.")
    except Exception as e:
        print(f"    [AVISO] Erro ao executar onmouseover: {e}")

    print("    -> Movendo mouse fisico sobre o elemento...")
    try:
        actions = ActionChains(driver)
        actions.move_to_element(el_imprimir).pause(1).perform()
        time.sleep(1)
    except Exception as e:
        print(f"    [AVISO] Erro no ActionChains: {e}")


# =============================================================================
# ESPERAS INTELIGENTES
# =============================================================================

def aguardar_loader_ccee(driver, timeout=20):
    """
    Aguarda o indicador de carregamento do Oracle BI desaparecer.
    Substitui o time.sleep(5) generico apos clicar em Aplicar.
    """
    xpaths_loader = [
        "//*[contains(@class, 'LoadingIndicator') and not(contains(@style, 'display: none'))]",
        "//*[contains(@class, 'BusyIndicator')]",
        "//div[contains(@id, 'loading')]",
    ]
    print("  -> Aguardando dados carregarem...", end="", flush=True)
    try:
        for xpath in xpaths_loader:
            try:
                WebDriverWait(driver, 3).until(
                    EC.presence_of_element_located((By.XPATH, xpath))
                )
                break
            except:
                continue

        for xpath in xpaths_loader:
            try:
                WebDriverWait(driver, timeout).until(
                    EC.invisibility_of_element_located((By.XPATH, xpath))
                )
                break
            except:
                continue

        print(" OK.")
    except:
        print(" (loader nao detectado, continuando)")


def aguardar_pdf_na_aba(driver, timeout=40):
    """
    Aguarda o PDF do Oracle BI carregar completamente na aba do Chrome (saw.dll).

    O Oracle BI injeta o conteudo do relatorio na pagina via JavaScript depois
    que o HTML inicial carrega. Por isso document.readyState == complete nao
    e suficiente -- precisamos esperar o HTML parar de crescer.

    Logica: monitora o tamanho do innerHTML a cada 1s. Quando fica igual
    por 3 segundos consecutivos, o conteudo esta estabilizado e o PDF esta
    pronto para ser impresso com Ctrl+P.
    """
    print("  -> Aguardando PDF carregar na aba...", end="", flush=True)

    # 1. Espera readyState = complete (carregamento basico do HTML)
    try:
        WebDriverWait(driver, 15).until(
            lambda d: d.execute_script("return document.readyState") == "complete"
        )
    except:
        pass

    # 2. Espera o conteudo parar de crescer (renderizacao JS do Oracle BI)
    tamanho_anterior = 0
    estavel_count = 0
    inicio = time.time()
    while time.time() - inicio < timeout:
        try:
            tamanho_atual = driver.execute_script(
                "return document.documentElement ? document.documentElement.innerHTML.length : 0"
            )
            if tamanho_atual > 500 and tamanho_atual == tamanho_anterior:
                estavel_count += 1
                if estavel_count >= 3:
                    print(" OK (conteudo estabilizado).")
                    return True
            else:
                estavel_count = 0
            tamanho_anterior = tamanho_atual
        except:
            pass
        time.sleep(1)

    print(" (timeout atingido, continuando mesmo assim)")
    return False


# =============================================================================
# SALVAR PDF (PyAutoGUI -- unico metodo compativel com Oracle BI da CCEE)
# O sistema Oracle BI da CCEE usa tokens de sessao vinculados ao processo
# do Chrome. Qualquer requisicao externa (requests, urllib) recebe 403.
# O unico caminho e controlar o dialogo de impressao do proprio Chrome.
# =============================================================================

def focar_janela_chrome(driver):
    """Garante que o Chrome esta em foco antes de usar PyAutoGUI."""
    try:
        driver.maximize_window()
        time.sleep(0.5)
        driver.execute_script("window.focus();")
        time.sleep(0.5)
        size = driver.get_window_size()
        posicao = driver.get_window_position()
        centro_x = posicao['x'] + (size['width'] // 2)
        centro_y = posicao['y'] + (size['height'] // 2)
        pyautogui.moveTo(centro_x, centro_y, duration=0.3)
        time.sleep(0.3)
    except Exception as e:
        print(f"  [AVISO] Erro ao focar janela: {e}")


def salvar_pdf_com_pyautogui(driver, nome_arquivo):
    """
    Salva o PDF via Ctrl+P usando PyAutoGUI.
    A aba ja deve estar na pagina do visualizador PDF (saw.dll) com o
    conteudo totalmente carregado antes de chamar esta funcao.
    """
    caminho_completo = os.path.join(PASTA_DOWNLOAD, nome_arquivo)
    print(f"  Salvando em: {caminho_completo}")

    # Garante foco no Chrome ANTES de qualquer comando de teclado
    focar_janela_chrome(driver)

    # Abre dialogo de impressao
    print("  -> Abrindo dialogo de impressao (Ctrl+P)...")
    pyautogui.hotkey('ctrl', 'p')
    time.sleep(4)

    # Pressiona Enter para confirmar (abre "Salvar como")
    print("  -> Confirmando impressao...")
    pyautogui.press('enter')
    time.sleep(4)

    # Digita o caminho completo do arquivo
    print("  -> Digitando caminho do arquivo...")
    pyautogui.write(caminho_completo, interval=0.05)
    time.sleep(1.5)

    # Confirma salvamento
    print("  -> Confirmando salvamento...")
    pyautogui.press('enter')
    time.sleep(4)

    # Aguarda o arquivo aparecer no disco
    for _ in range(10):
        if os.path.exists(caminho_completo):
            print("  [OK] PDF salvo com sucesso!")
            return True
        time.sleep(1)

    print("  [AVISO] Arquivo nao encontrado apos salvar.")
    return False


# =============================================================================
# RECUPERACAO DE ERROS
# =============================================================================

def resetar_estado(driver, janela_principal):
    """Fecha abas extras e volta para a janela principal apos uma falha."""
    print("  [RESET] Limpando estado para proxima empresa...")
    try:
        for handle in driver.window_handles:
            if handle != janela_principal:
                driver.switch_to.window(handle)
                driver.close()
        driver.switch_to.window(janela_principal)
        driver.switch_to.default_content()
    except Exception as e:
        print(f"  [RESET] Aviso: {e}")


# =============================================================================
# LOOP PRINCIPAL
# =============================================================================

def robo_extrator():
    driver = conectar_chrome_aberto()
    actions = ActionChains(driver)

    configurar_pasta_download()

    print(f"--- Iniciando extracao para {len(empresas_alvo)} empresas ---")
    print(f"--- PDFs serao salvos em: {PASTA_DOWNLOAD} ---")
    print("\n[!] NAO CLIQUE NO COMPUTADOR DURANTE A EXECUCAO!")
    print("[!] O PyAutoGUI precisa controlar mouse e teclado.\n")
    time.sleep(3)

    janela_principal = driver.current_window_handle
    sucesso = 0
    falhas = 0

    for empresa in empresas_alvo:
        try:
            print(f"\n> Processando: {empresa}")

            # Validacao: esta no site correto?
            if not validar_url_ccee(driver):
                resposta = input("  [SEGURANCA] Nao esta no site da CCEE. Continuar? (s/n): ")
                if resposta.lower() != 's':
                    print("  [ABORTADO] Processo cancelado pelo usuario.")
                    break

            # [PASSO 1] SELECIONAR A EMPRESA NA LISTA
            xpath_empresa = f"//div[contains(@class, 'listBoxPanelOptionBasic') and @title='{empresa}']"
            try:
                el_empresa = buscar_elemento_inteligente(driver, xpath_empresa, tempo=5)
                driver.execute_script("arguments[0].scrollIntoView({block: 'center', inline: 'nearest'});", el_empresa)
                time.sleep(0.5)
                actions.move_to_element(el_empresa).click().perform()
                print("  Selecionado na lista.")
            except:
                print(f"  [ERRO] Nao achei a empresa '{empresa}'.")
                falhas += 1
                continue

            # [PASSO 2] CLICAR EM "APLICAR"
            try:
                time.sleep(1)
                try:
                    btn_aplicar = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID, "gobtn")))
                    actions.move_to_element(btn_aplicar).click().perform()
                except:
                    btn_aplicar = buscar_elemento_inteligente(driver, "//input[@value='Aplicar']")
                    clicar_js(driver, btn_aplicar)
            except:
                print("  [ERRO] Botao aplicar falhou.")
                falhas += 1
                continue

            # Aguarda Oracle BI carregar os dados (substitui time.sleep(5) fixo)
            aguardar_loader_ccee(driver, timeout=20)

            # Valida se ha conteudo real na tela
            if not validar_conteudo_carregado(driver):
                print("  [ERRO] Conteudo nao carregou ou esta vazio. Pulando empresa.")
                falhas += 1
                continue

            # [PASSO 3] CLICAR NA ENGRENAGEM
            print("  Abrindo opcoes...", end="")
            xpath_engrenagem = "//img[contains(@id, 'dashboardpageoptions') or contains(@src, 'popupmenu')]"
            try:
                btn_opcoes = buscar_elemento_inteligente(driver, xpath_engrenagem)
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn_opcoes)
                time.sleep(0.5)
                clicar_js(driver, btn_opcoes)
                print(" OK.")
            except:
                print(" FALHA (Engrenagem).")
                falhas += 1
                continue

            time.sleep(2)

            # [PASSO 4] LOCALIZAR BOTAO IMPRIMIR (sem clicar ainda)
            print("  Buscando botao 'Imprimir'...")
            try:
                el_imprimir = buscar_elemento_inteligente(driver, "//*[@id='idPagePrint']", tempo=5)
                print("  [OK] Botao 'Imprimir' localizado.")
            except Exception as e:
                print(f"  [ERRO] Botao 'Imprimir' nao encontrado: {e}")
                falhas += 1
                continue

            # [PASSO 5] ABRIR SUBMENU DE IMPRESSAO
            print("  Disparando submenu PDF...")
            disparar_submenu_imprimir(driver, el_imprimir)
            time.sleep(2)

            # [PASSO 6] CLICAR NA OPCAO "PDF" DO SUBMENU
            print("  Procurando opcao 'PDF' no submenu...")
            el_pdf = buscar_submenu_pdf(driver, tempo_max=8)

            if el_pdf is None:
                print("  [ERRO] Submenu PDF nao encontrado.")
                try:
                    screenshot_path = os.path.join(PASTA_DOWNLOAD, f"erro_submenu_{empresa.replace('/', '-')}.png")
                    driver.save_screenshot(screenshot_path)
                    print(f"  [DEBUG] Screenshot: {screenshot_path}")
                except:
                    pass
                falhas += 1
                continue

            try:
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", el_pdf)
                time.sleep(0.3)
                try:
                    actions.move_to_element(el_pdf).click().perform()
                except:
                    clicar_js(driver, el_pdf)
                print("  [SUCESSO] Clicou em 'PDF'.")
            except Exception as e:
                print(f"  [ERRO] Falha ao clicar no PDF: {e}")
                falhas += 1
                continue

            # [PASSO 7] DETECTAR A NOVA ABA E AGUARDAR O PDF CARREGAR COMPLETAMENTE
            # O Oracle BI abre o relatorio numa nova aba (saw.dll).
            # Precisamos esperar o conteudo estabilizar ANTES do Ctrl+P,
            # senao o PDF sai em branco ou com dados incompletos.
            print("  Aguardando nova aba do PDF abrir...")
            nova_aba = None
            for _ in range(20):  # polling a cada 0.5s, ate 10s
                time.sleep(0.5)
                if len(driver.window_handles) > 1:
                    nova_aba = driver.window_handles[-1]
                    break

            if nova_aba is None:
                print("  [ERRO] Nova aba com PDF nao detectada.")
                falhas += 1
                continue

            driver.switch_to.window(nova_aba)
            print("  [OK] Aba do PDF detectada.")

            # Espera o Oracle BI terminar de renderizar os dados na aba
            aguardar_pdf_na_aba(driver, timeout=40)

            # [PASSO 8] SALVAR O PDF VIA Ctrl+P (PyAutoGUI)
            nome_arquivo = (
                f"{empresa.replace(' ', '_')}_{NOME_RELATORIO}_{MES}_{ANO}.pdf"
                .replace("/", "-")
                .replace("\\", "-")
            )

            if salvar_pdf_com_pyautogui(driver, nome_arquivo):
                sucesso += 1
            else:
                falhas += 1

            # Fecha a aba do PDF e volta para a janela principal
            try:
                if driver.current_window_handle != janela_principal:
                    driver.close()
                    print("  [OK] Aba do PDF fechada.")
                driver.switch_to.window(janela_principal)
                print("  [OK] Foco retornado para janela principal.")
            except Exception as e:
                print(f"  [AVISO] Erro ao fechar aba: {e}")
                try:
                    driver.switch_to.window(janela_principal)
                except:
                    pass

            driver.switch_to.default_content()
            print(f"  [OK] {empresa} processado!")

        except Exception as e:
            print(f"  [FALHA GERAL] {empresa}: {e}")
            falhas += 1
            resetar_estado(driver, janela_principal)

    print("\n" + "=" * 60)
    print("--- EXTRACAO FINALIZADA ---")
    print(f"Total de empresas: {len(empresas_alvo)}")
    print(f"Sucesso: {sucesso}")
    print(f"Falhas:  {falhas}")
    print(f"PDFs salvos em: {PASTA_DOWNLOAD}")
    print("=" * 60)


if __name__ == "__main__":
    pyautogui.FAILSAFE = True
    pyautogui.PAUSE = 0.5
    iniciar_chrome_debug()
    robo_extrator()