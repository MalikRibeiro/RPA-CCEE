from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from config import empresas_alvo, MES, ANO, NOME_RELATORIO, PASTA_DOWNLOAD, URL_ESPERADA_CCEE, ELEMENTOS_VALIDACAO
import time
import sys
import pyautogui
import os

def conectar_chrome_aberto():
    print("Conectando ao Chrome (Porta 9222)...")
    options = Options()
    options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
    try:
        driver = webdriver.Chrome(options=options)
        return driver
    except Exception as e:
        print(f"[ERRO] Chrome não encontrado.\n1. Execute o .bat\n2. Logue na CCEE.\nErro: {e}")
        sys.exit()

def focar_janela_chrome(driver):
    """
    Força o foco no Chrome antes de usar PyAutoGUI
    """
    try:
        # Maximiza a janela do Chrome para garantir que está visível
        driver.maximize_window()
        time.sleep(0.5)
        
        # Executa JavaScript para dar foco na janela
        driver.execute_script("window.focus();")
        time.sleep(0.5)
        
        # Move o mouse para o centro da janela do Chrome
        # Isso ajuda a garantir que os comandos do PyAutoGUI vão para o Chrome
        try:
            # Pega as dimensões da janela
            size = driver.get_window_size()
            posicao = driver.get_window_position()
            
            # Calcula o centro da janela
            centro_x = posicao['x'] + (size['width'] // 2)
            centro_y = posicao['y'] + (size['height'] // 2)
            
            # Move o mouse para o centro
            pyautogui.moveTo(centro_x, centro_y, duration=0.3)
            time.sleep(0.3)
        except:
            pass
            
    except Exception as e:
        print(f"  [AVISO] Erro ao focar janela: {e}")

def aguardar_pdf_carregar(driver, timeout=15):
    """
    Aguarda o PDF carregar completamente na nova aba antes de tentar imprimir
    """
    print("  → Aguardando PDF renderizar...")
    
    inicio = time.time()
    while time.time() - inicio < timeout:
        try:
            # Verifica se a página não está mais carregando
            estado = driver.execute_script("return document.readyState")
            
            # Verifica se há um visualizador de PDF na página
            tem_pdf = driver.execute_script("""
                // Verifica se existe um embed ou object de PDF
                return document.querySelector('embed[type="application/pdf"]') !== null ||
                       document.querySelector('object[type="application/pdf"]') !== null ||
                       document.querySelector('iframe[src*="pdf"]') !== null ||
                       // Visualizador do Chrome
                       document.querySelector('#viewer') !== null ||
                       // Verifica se a URL contém indicadores de PDF
                       window.location.href.includes('pdf');
            """)
            
            if estado == "complete" and tem_pdf:
                print("  [OK] PDF carregado!")
                time.sleep(2)  # Tempo adicional para renderização visual
                return True
                
        except Exception as e:
            pass
        
        time.sleep(0.5)
    
    print("  [AVISO] Timeout ao aguardar PDF. Tentando continuar...")
    return False

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

def salvar_pdf_com_caminho(driver, nome_arquivo):
    """
    Salva o PDF usando PyAutoGUI com caminho completo
    IMPORTANTE: Garante que o Chrome tem foco antes de executar
    """
    caminho_completo = os.path.join(PASTA_DOWNLOAD, nome_arquivo)
    
    print(f"  Salvando em: {caminho_completo}")
    
    # CRÍTICO: Foca no Chrome antes de usar PyAutoGUI
    focar_janela_chrome(driver)
    
    # Abre diálogo de impressão
    print("  → Abrindo diálogo de impressão (Ctrl+P)...")
    pyautogui.hotkey('ctrl', 'p')
    time.sleep(4)  # Aumentado para dar tempo do diálogo abrir
    
    # Confirma impressão (abre salvar como)
    print("  → Confirmando impressão...")
    pyautogui.press('enter')
    time.sleep(4)  # Aumentado para dar tempo do diálogo "Salvar como" abrir
    
    # Digita o caminho COMPLETO (garante que vai para a pasta correta)
    print("  → Digitando caminho do arquivo...")
    pyautogui.write(caminho_completo, interval=0.05)
    time.sleep(1.5)
    
    # Confirma salvamento
    print("  → Confirmando salvamento...")
    pyautogui.press('enter')
    time.sleep(4)
    
    # Verifica se o arquivo foi criado
    tentativas = 0
    while tentativas < 10:
        if os.path.exists(caminho_completo):
            print(f"  [OK] PDF salvo com sucesso!")
            return True
        time.sleep(1)
        tentativas += 1
    
    print(f"  [AVISO] Arquivo não encontrado após salvar.")
    return False

def robo_extrator():
    driver = conectar_chrome_aberto()
    actions = ActionChains(driver)
    
    # Configura pasta de download
    configurar_pasta_download()
    
    print(f"--- Iniciando extração para {len(empresas_alvo)} empresas ---")
    print(f"--- PDFs serão salvos em: {PASTA_DOWNLOAD} ---")
    print("\n⚠️  IMPORTANTE: NÃO CLIQUE NO COMPUTADOR DURANTE A EXECUÇÃO!")
    print("⚠️  O PyAutoGUI precisa controlar mouse e teclado.\n")
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

            print("  Aguardando carregamento dos dados...")
            time.sleep(5)
            
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

            # [PASSO 7] AGUARDAR NOVA JANELA E PDF CARREGAR
            print("  Aguardando nova janela...")
            time.sleep(5)
            nova_janela_abriu = False
            janelas = driver.window_handles
            
            if len(janelas) > 1:
                driver.switch_to.window(janelas[-1])
                nova_janela_abriu = True
                print("  [OK] Nova janela detectada.")
                
                # CRÍTICO: Aguarda o PDF carregar ANTES de tentar imprimir
                aguardar_pdf_carregar(driver, timeout=20)
                
            else:
                print("  [AVISO] Janela não abriu. Tentando salvar na aba atual.")
                # Mesmo sem nova janela, aguarda um pouco
                
                time.sleep(3)

            # [PASSO 8] SALVAR PDF COM VALIDAÇÃO
            nome_arquivo = f"{empresa.replace(' ', '_')}_{NOME_RELATORIO}_{MES}_{ANO}.pdf".replace("/", "-").replace("\\", "-")
            
            if salvar_pdf_com_caminho(driver, nome_arquivo):
                sucesso += 1
            else:
                falhas += 1
            
            # Fecha janela se abriu nova aba
            if nova_janela_abriu:
                try:
                    time.sleep(2)
                    # PROTEÇÃO: Verifica se a janela atual NÃO é a principal antes de fechar
                    janela_atual = driver.current_window_handle
                    if janela_atual != janela_principal:
                        driver.close() # Só fecha se for realmente uma aba extra
                        print("  [OK] Aba do PDF fechada.")
                    else:
                        print("  [AVISO] O robô tentou fechar a janela principal, mas foi impedido.")
                    
                    # Garante o retorno para a janela principal
                    driver.switch_to.window(janela_principal)
                    print("  [OK] Foco retornado para janela principal.")
                    
                except Exception as e:
                    print(f"  [AVISO] Erro ao manipular fechamento de janela: {e}")
                    # Tenta forçar a volta para a principal em caso de erro
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
            try:
                # Tenta voltar para a janela principal
                driver.switch_to.window(janela_principal)
                driver.switch_to.default_content()
            except:
                pass

    # Relatório final
    print("\n" + "="*60)
    print(f"--- EXTRAÇÃO FINALIZADA ---")
    print(f"Total de empresas: {len(empresas_alvo)}")
    print(f"✓ Sucesso: {sucesso}")
    print(f"✗ Falhas: {falhas}")
    print(f"PDFs salvos em: {PASTA_DOWNLOAD}")
    print("="*60)

if __name__ == "__main__":
    pyautogui.FAILSAFE = True 
    pyautogui.PAUSE = 0.5  # Pausa de 0.5s entre comandos do PyAutoGUI
    robo_extrator()