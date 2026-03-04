import time
import sys
import os
import glob
import requests
import subprocess

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from config import empresas_alvo, MES, ANO, NOME_RELATORIO, PASTA_DOWNLOAD, URL_ESPERADA_CCEE, ELEMENTOS_VALIDACAO

def iniciar_chrome_debug():
    """Abre o Chrome com porta de debug, sem precisar do .bat"""
    caminho_chrome = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
    pasta_perfil = r"C:\selenium\chrome_perfil"  # mesmo caminho do seu .bat

    # Verifica se Chrome já está aberto na porta 9222
    import urllib.request
    try:
        urllib.request.urlopen("http://127.0.0.1:9222/json", timeout=2)
        print("[INFO] Chrome já está rodando na porta 9222.")
        return  # já está aberto, não abre de novo
    except:
        pass  # não está aberto, vamos abrir

    print("[INFO] Iniciando Chrome com porta de debug...")
    subprocess.Popen([
        caminho_chrome,
        "--remote-debugging-port=9222",
        f"--user-data-dir={pasta_perfil}",
        "https://operacao.ccee.org.br/"
    ])
    time.sleep(4)  # tempo para o Chrome abrir
    print("[INFO] Faça o login na CCEE e selecione o relatório.")
    input("[ENTER] Pressione Enter quando estiver pronto para o robô começar...")

def conectar_chrome_aberto():
    options = Options()
    options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
    try:
        driver = webdriver.Chrome(options=options)

        # INJEÇÃO CDP: Chrome baixa o PDF direto na pasta, sem abrir visualizador
        driver.execute_cdp_cmd("Page.setDownloadBehavior", {
            "behavior": "allow",
            "downloadPath": PASTA_DOWNLOAD
        })

        return driver
    except Exception as e:
        print(f"[ERRO] Chrome não encontrado. Execute o script novamente.\nErro: {e}")
        sys.exit()

def validar_url_ccee(driver):
    """Verifica se está no site correto da CCEE"""
    url_atual = driver.current_url.lower()
    
    for url_valida in URL_ESPERADA_CCEE:
        if url_valida.lower() in url_atual:
            return True
    
    print(f"[ALERTA] URL não é da CCEE: {url_atual}")
    return False

def validar_conteudo_carregado(driver):
    """
    Verifica se há conteúdo válido na tela antes de tentar exportar PDF
    """
    print("  → Validando conteúdo da página...")
    
    # 1. Verifica se não há mensagem de erro
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
    
    # 2. Verifica elementos esperados do relatório
    elementos_encontrados = 0
    
    # Tenta no root
    driver.switch_to.default_content()
    for xpath_validacao in ELEMENTOS_VALIDACAO:
        try:
            if driver.find_elements(By.XPATH, xpath_validacao):
                elementos_encontrados += 1
        except:
            pass
    
    # Tenta nos iframes
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
        print(f"  [OK] Conteúdo validado ({elementos_encontrados} elementos encontrados).")
        return True
    else:
        print("  [ERRO] Nenhum conteúdo de relatório detectado!")
        return False

def buscar_elemento_inteligente(driver, xpath, tempo=5):
    """Procura elemento na página e dentro de iframes (Varredura Completa)"""
    wait = WebDriverWait(driver, tempo)
    
    # 1. Tenta no conteúdo padrão (Root)
    driver.switch_to.default_content()
    try:
        return wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
    except:
        pass 

    # 2. Tenta nos Iframes
    iframes = driver.find_elements(By.TAG_NAME, "iframe")
    for frame in iframes:
        try:
            driver.switch_to.default_content()
            driver.switch_to.frame(frame)
            return wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
        except:
            continue

    # Se falhar, volta pro default para não quebrar o fluxo
    driver.switch_to.default_content()
    raise Exception(f"Elemento não encontrado: {xpath}")

def buscar_submenu_pdf(driver, tempo_max=10):
    """
    Procura o submenu PDF em todos os contextos possíveis
    (root, iframes, e elementos dinâmicos)
    """
    # Possíveis XPaths para o botão PDF
    xpaths_pdf = [
        "//td[contains(@class, 'MenuItemTextCell') and contains(text(), 'PDF')]",
        "//td[contains(text(), 'Página Atual como PDF')]",
        "//a[contains(@class, 'MenuItem') and contains(., 'PDF')]",
        "//*[contains(@id, 'PDF') or contains(@id, 'pdf')]//td[contains(@class, 'MenuItemTextCell')]",
        "//td[@class='MenuItemTextCell' and normalize-space(text())='PDF']"
    ]
    
    inicio = time.time()
    while time.time() - inicio < tempo_max:
        # Tenta no contexto atual (root)
        driver.switch_to.default_content()
        for xpath in xpaths_pdf:
            try:
                elementos = driver.find_elements(By.XPATH, xpath)
                for el in elementos:
                    if el.is_displayed():
                        return el
            except:
                continue
        
        # Tenta em iframes
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
    """
    Força a criação do submenu usando múltiplas estratégias
    """
    # ESTRATÉGIA 1: Disparar evento mouseover via JavaScript
    print("    → Disparando evento onmouseover via JS...")
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
    
    # ESTRATÉGIA 2: Executar o código do onmouseover diretamente
    print("    → Executando função saw.dashboard.DisplayLayouts...")
    try:
        # Extrai os parâmetros do atributo onmouseover
        onmouseover_attr = el_imprimir.get_attribute("onmouseover")
        
        if onmouseover_attr and "saw.dashboard.DisplayLayouts" in onmouseover_attr:
            # Executa a função diretamente via JS
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
            print("    [OK] Função executada.")
        else:
            print("    [AVISO] Atributo onmouseover não encontrado.")
    except Exception as e:
        print(f"    [AVISO] Erro ao executar onmouseover: {e}")
    
    # ESTRATÉGIA 3: Mover mouse e pausar (deixa JS do site executar)
    print("    → Movendo mouse físico sobre o elemento...")
    try:
        actions = ActionChains(driver)
        actions.move_to_element(el_imprimir).pause(1).perform()
        time.sleep(1)
    except Exception as e:
        print(f"    [AVISO] Erro no ActionChains: {e}")

def configurar_pasta_download():
    """
    Garante que a pasta de download existe e retorna o caminho completo
    """
    if not os.path.exists(PASTA_DOWNLOAD):
        print(f"[AVISO] Pasta não existe, criando: {PASTA_DOWNLOAD}")
        try:
            os.makedirs(PASTA_DOWNLOAD)
        except Exception as e:
            print(f"[ERRO] Não foi possível criar a pasta: {e}")
            sys.exit()
    
    return PASTA_DOWNLOAD

def resetar_estado(driver, janela_principal):
    """
    Após uma falha, fecha abas extras, volta para a janela principal
    e recarrega para limpar filtros/menus travados.
    """
    print("  [RESET] Limpando estado para próxima empresa...")
    try:
        # Fecha todas as abas extras
        for handle in driver.window_handles:
            if handle != janela_principal:
                driver.switch_to.window(handle)
                driver.close()

        # Volta para a janela principal
        driver.switch_to.window(janela_principal)
        driver.switch_to.default_content()

    except Exception as e:
        print(f"  [RESET] Aviso: {e}")

def aguardar_loader_ccee(driver, timeout=20):
    """
    Aguarda o indicador de carregamento do Oracle BI desaparecer.
    Substitui o time.sleep(5) genérico.
    """
    # Tenta no root e nos iframes
    xpaths_loader = [
        "//*[contains(@class, 'LoadingIndicator') and not(contains(@style, 'display: none'))]",
        "//*[contains(@class, 'BusyIndicator')]",
        "//div[contains(@id, 'loading')]",
    ]
    print("  → Aguardando dados carregarem...", end="")
    try:
        # Espera qualquer loader aparecer (máx 3s) — pode não aparecer em telas rápidas
        for xpath in xpaths_loader:
            try:
                WebDriverWait(driver, 3).until(
                    EC.presence_of_element_located((By.XPATH, xpath))
                )
                break
            except:
                continue

        # Depois espera sumir
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
        print(" (loader não detectado, continuando)")

def baixar_pdf_da_aba(driver, caminho_completo, janela_principal, timeout=30):
    """
    Baixa o PDF diretamente da URL da aba do Chrome (saw.dll),
    reusando os cookies da sessão autenticada do Selenium.
    Obtém o arquivo PDF original, sem reimprimir nem capturar tela.
    """
    import requests

    print("  → Aguardando URL do PDF ficar disponível na aba...")

    # Aguarda a URL conter o token de download (indica que o PDF está pronto)
    url_pdf = None
    inicio = time.time()
    while time.time() - inicio < timeout:
        url_atual = driver.current_url
        if "downloadExportedFile" in url_atual or ("saw.dll" in url_atual and "Format=pdf" in url_atual):
            url_pdf = url_atual
            break
        time.sleep(0.5)

    if not url_pdf:
        print("  [ERRO] URL de download do PDF não detectada na aba.")
        return False

    print(f"  → Baixando PDF via requests (URL autenticada)...")

    # Extrai cookies do Selenium para autenticar a requisição
    selenium_cookies = driver.get_cookies()
    session = requests.Session()
    for cookie in selenium_cookies:
        session.cookies.set(cookie['name'], cookie['value'], domain=cookie.get('domain', ''))

    # Copia headers do navegador para evitar bloqueio por User-Agent
    session.headers.update({
        "User-Agent": driver.execute_script("return navigator.userAgent;"),
        "Referer": driver.current_url,
    })

    try:
        resposta = session.get(url_pdf, timeout=60, stream=True)
        resposta.raise_for_status()

        # Verifica se a resposta é realmente um PDF
        content_type = resposta.headers.get("Content-Type", "")
        if "pdf" not in content_type.lower() and len(resposta.content) < 1000:
            print(f"  [ERRO] Resposta inesperada (Content-Type: {content_type})")
            return False

        with open(caminho_completo, "wb") as f:
            for chunk in resposta.iter_content(chunk_size=8192):
                f.write(chunk)

        tamanho_kb = os.path.getsize(caminho_completo) / 1024
        print(f"  [OK] PDF salvo: {os.path.basename(caminho_completo)} ({tamanho_kb:.1f} KB)")
        return True

    except Exception as e:
        print(f"  [ERRO] Falha ao baixar PDF: {e}")
        return False

def robo_extrator():
    driver = conectar_chrome_aberto()
    actions = ActionChains(driver)
    
    # Configura pasta de download
    configurar_pasta_download()
    
    print(f"--- Iniciando extração para {len(empresas_alvo)} empresas ---")
    print(f"--- PDFs serão salvos em: {PASTA_DOWNLOAD} ---")
    time.sleep(3)

    janela_principal = driver.current_window_handle
    
    # Contadores de resultados
    sucesso = 0
    falhas = 0

    for empresa in empresas_alvo:
        try:
            print(f"\n> Processando: {empresa}")

            # VALIDAÇÃO INICIAL: Verifica se está no site correto
            if not validar_url_ccee(driver):
                resposta = input("  [SEGURANÇA] Não está no site da CCEE. Continuar? (s/n): ")
                if resposta.lower() != 's':
                    print("  [ABORTADO] Processo cancelado pelo usuário.")
                    break

            # [PASSO 1] SELECIONAR A EMPRESA
            xpath_empresa = f"//div[contains(@class, 'listBoxPanelOptionBasic') and @title='{empresa}']"
            try:
                el_empresa = buscar_elemento_inteligente(driver, xpath_empresa, tempo=5)
                driver.execute_script("arguments[0].scrollIntoView({block: 'center', inline: 'nearest'});", el_empresa)
                time.sleep(0.5)
                actions.move_to_element(el_empresa).click().perform()
                print(f"  Selecionado na lista.")
            except Exception as e:
                print(f"  [ERRO] Não achei a empresa '{empresa}'.")
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
                print(f"  [ERRO] Botão aplicar falhou.")
                falhas += 1
                continue

            aguardar_loader_ccee(driver, timeout=20)
            
            # VALIDAÇÃO: Verifica se há conteúdo carregado
            if not validar_conteudo_carregado(driver):
                print("  [ERRO] Conteúdo não carregou ou está vazio. Pulando empresa.")
                falhas += 1
                continue

            # [PASSO 3] CLICAR NA ENGRENAGEM
            print("  Abrindo opções...", end="")
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

            # [PASSO 4] LOCALIZAR BOTÃO IMPRIMIR (SEM CLICAR AINDA)
            print("  Buscando botão 'Imprimir'...")
            xpath_imprimir_id = "//*[@id='idPagePrint']" 
            
            try:
                el_imprimir = buscar_elemento_inteligente(driver, xpath_imprimir_id, tempo=5)
                print("  [OK] Botão 'Imprimir' localizado.")
            except Exception as e:
                print(f"  [ERRO] Botão 'Imprimir' não encontrado: {e}")
                falhas += 1
                continue

            # [PASSO 5] DISPARAR SUBMENU
            print("  Disparando submenu PDF...")
            disparar_submenu_imprimir(driver, el_imprimir)
            
            # Aguarda um pouco mais para o submenu renderizar
            time.sleep(2)

            # [PASSO 6] BUSCAR E CLICAR NO PDF
            print("  Procurando opção 'PDF' no submenu...")
            el_pdf = buscar_submenu_pdf(driver, tempo_max=8)
            
            if el_pdf is None:
                print("  [ERRO] Submenu PDF não encontrado após múltiplas tentativas.")
                print("  [DEBUG] Salvando screenshot...")
                try:
                    screenshot_path = os.path.join(PASTA_DOWNLOAD, f"erro_submenu_{empresa.replace('/', '-')}.png")
                    driver.save_screenshot(screenshot_path)
                    print(f"  [DEBUG] Screenshot salvo em: {screenshot_path}")
                except:
                    pass
                falhas += 1
                continue
            
            # Clica no PDF
            try:
                # Garante que o elemento está visível
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", el_pdf)
                time.sleep(0.3)
                
                # Tenta clicar (ActionChains primeiro, JS como backup)
                try:
                    actions.move_to_element(el_pdf).click().perform()
                except:
                    clicar_js(driver, el_pdf)
                
                print("  [SUCESSO] Clicou em 'PDF'.")
            except Exception as e:
                print(f"  [ERRO] Falha ao clicar no PDF: {e}")
                falhas += 1
                continue

            # [PASSO 7] DETECTA A NOVA ABA COM O PDF E MUDA O FOCO
            print("  Aguardando nova aba do PDF abrir...")
            nova_aba = None
            for _ in range(20):  # aguarda até 10s (20x 0.5s)
                time.sleep(0.5)
                janelas = driver.window_handles
                if len(janelas) > 1:
                    nova_aba = janelas[-1]
                    break

            if nova_aba is None:
                print("  [ERRO] Nova aba com PDF não detectada.")
                falhas += 1
                continue

            driver.switch_to.window(nova_aba)
            print("  [OK] Aba do PDF detectada.")

            # [PASSO 8] BAIXA O PDF ORIGINAL DIRETAMENTE DA URL DA ABA
            nome_arquivo = f"{empresa.replace(' ', '_')}_{NOME_RELATORIO}_{MES}_{ANO}.pdf".replace("/", "-").replace("\\", "-")
            caminho_completo = os.path.join(PASTA_DOWNLOAD, nome_arquivo)

            if baixar_pdf_da_aba(driver, caminho_completo, janela_principal):
                sucesso += 1
            else:
                falhas += 1

            # Fecha a aba do PDF e volta para janela principal
            try:
                driver.close()
                driver.switch_to.window(janela_principal)
                print("  [OK] Aba do PDF fechada.")
            except Exception as e:
                print(f"  [AVISO] Erro ao fechar aba: {e}")
                try:
                    driver.switch_to.window(janela_principal)
                except:
                    pass

            # Reseta contexto para o próximo loop
            driver.switch_to.default_content()
            print(f"  ✓ {empresa} processado!")

        except Exception as e:
            print(f"  [FALHA GERAL] {empresa}: {e}")
            falhas += 1
            resetar_estado(driver, janela_principal)

    # Relatório final
    print("\n" + "="*60)
    print(f"--- EXTRAÇÃO FINALIZADA ---")
    print(f"Total de empresas: {len(empresas_alvo)}")
    print(f"✓ Sucesso: {sucesso}")
    print(f"✗ Falhas: {falhas}")
    print(f"PDFs salvos em: {PASTA_DOWNLOAD}")
    print("="*60)

if __name__ == "__main__":
    iniciar_chrome_debug()
    robo_extrator()