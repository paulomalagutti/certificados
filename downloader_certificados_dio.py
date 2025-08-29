import time
import os
import re
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def setup_driver(pasta_downloads):
    """Configura o driver do Chrome com opções para evitar detecção e define a pasta de download."""
    options = Options()
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--start-maximized")
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
    
    prefs = {
        "download.default_directory": pasta_downloads,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True
    }
    options.add_experimental_option("prefs", prefs)
    
    driver = webdriver.Chrome(options=options)
    return driver

def sanitizar_nome_arquivo(texto):
    """Limpa um texto para que ele possa ser usado como nome de arquivo."""
    if not texto:
        return "certificado_sem_nome"
    
    texto_limpo = re.sub(r'[\\/*?:"<>|]', "", texto)
    texto_limpo = re.sub(r'\s+', ' ', texto_limpo).strip()
    
    return texto_limpo

def baixar_certificados(driver, pasta_certificados):
    """Navega até a página de certificados e baixa todos eles."""
    print("\nNavegando para a página de certificados...")
    driver.get("https://web.dio.me/certificates")
    
    if not os.path.exists(pasta_certificados):
        os.makedirs(pasta_certificados)
        print(f"📁 Pasta '{pasta_certificados}' criada com sucesso.")
        
    # Aumentando o tempo de espera para 40 segundos por segurança
    wait = WebDriverWait(driver, 40)
    
    print("Carregando todos os certificados na página (scroll)...")
    last_height = driver.execute_script("return document.body.scrollHeight")
    while True:
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(3)
        new_height = driver.execute_script("return document.body.scrollHeight")
        if new_height == last_height:
            break
        last_height = new_height
            
    try:
        # --- SELETOR PRINCIPAL CORRIGIDO ---
        # Agora procuramos por um item de lista (li) que contenha um h5 E um botão de baixar.
        # Isso é muito mais robusto e corresponde à estrutura das suas imagens.
        seletor_container = "//li[.//h5 and .//button[contains(., 'BAIXAR CERTIFICADO')]]"
        
        print("Aguardando a lista de certificados carregar...")
        wait.until(EC.presence_of_all_elements_located((By.XPATH, seletor_container)))
        certificados = driver.find_elements(By.XPATH, seletor_container)
        
        if not certificados:
            print("❌ Nenhum certificado encontrado na página.")
            return

        print(f"✅ Encontrados {len(certificados)} certificados. Iniciando downloads...")
        
        for i, cert in enumerate(certificados):
            try:
                print(f"\n--- Processando certificado {i+1}/{len(certificados)} ---")

                # A lógica interna para achar o nome e o botão já estava correta.
                titulo_elemento = cert.find_element(By.CSS_SELECTOR, "h5")
                nome_certificado = sanitizar_nome_arquivo(titulo_elemento.text)
                print(f"  > Título: {nome_certificado}")

                caminho_final_arquivo = os.path.join(pasta_certificados, f"{nome_certificado}.pdf")

                if os.path.exists(caminho_final_arquivo):
                    print("  > 🟡 Aviso: Certificado já existe. Pulando.")
                    continue

                arquivos_antes = set(os.listdir(pasta_certificados))

                botao_baixar = cert.find_element(By.XPATH, ".//button[contains(., 'BAIXAR CERTIFICADO')]")
                driver.execute_script("arguments[0].click();", botao_baixar)
                
                print("  > ⏳ Clicou para baixar. Aguardando a conclusão do download...")

                novo_arquivo_path = None
                for _ in range(30): # Espera no máximo 30 segundos
                    time.sleep(1)
                    arquivos_depois = set(os.listdir(pasta_certificados))
                    novos_arquivos = arquivos_depois - arquivos_antes
                    
                    if novos_arquivos:
                        nome_novo_arquivo = novos_arquivos.pop()
                        if not nome_novo_arquivo.endswith(('.crdownload', '.tmp')):
                            novo_arquivo_path = os.path.join(pasta_certificados, nome_novo_arquivo)
                            break
                
                if novo_arquivo_path:
                    os.rename(novo_arquivo_path, caminho_final_arquivo)
                    print(f"  > ✅ Sucesso: Salvo e renomeado para '{os.path.basename(caminho_final_arquivo)}'")
                else:
                    print("  > ❌ Erro: Download não foi concluído a tempo ou não foi detectado.")

            except Exception as e:
                print(f"  > 💥 Erro ao processar um certificado: {e}")
                continue

    except Exception as e:
        print(f"❌ Erro fatal ao buscar a lista de certificados: {e}")

def main():
    print("🚀 === INICIANDO DOWNLOADER DE CERTIFICADOS DA DIO ===")
    
    pasta_destino = r"C:\Users\paulo\OneDrive\Documentos\CERTIFICADOS"
    
    driver = setup_driver(pasta_destino)
    
    try:
        driver.get("https://web.dio.me/sign-in")
        print("\n--- AÇÃO NECESSÁRIA ---")
        print("Uma janela do navegador foi aberta. Por favor, realize o login manualmente.")
        
        input("Após fazer o login no navegador, volte aqui e pressione ENTER para continuar...")
        
        print("\nLogin confirmado pelo usuário. Continuando o processo...")
        
        baixar_certificados(driver, pasta_destino)
        
        print("\n🎉 Processo finalizado com sucesso!")
    
    except Exception as e:
        print(f"\n❌ Ocorreu um erro inesperado no processo: {e}")
        import traceback
        traceback.print_exc()
    
    finally:
        print("\nO navegador será fechado em 15 segundos...")
        time.sleep(15)
        driver.quit()

if __name__ == "__main__":
    main()