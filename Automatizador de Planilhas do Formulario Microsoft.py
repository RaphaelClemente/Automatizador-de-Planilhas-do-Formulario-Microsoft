from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.chrome.options import Options # Importação necessária
from selenium.webdriver.chrome.service import Service as ChromeService # Melhor prática
from webdriver_manager.chrome import ChromeDriverManager # Melhor prática
import time
import os


load_dotenv()

# Credenciais de login
EMAIL = os.getenv("EMAIL_MS")
SENHA = os.getenv("SENHA_MS")

# Lista de URLs para verificar
urls = {
    "Teste":[
    ("Nome da Planilha","https://example.com")  # Exemplo de URL],
    ],
    "Outros":[]
}



# Tempo de espera antes de atualizar (em segundos)
TEMPO_ESPERA = 25



# =========================================================================
# --- Configuração do navegador (MODIFICADO PARA MODO HEADLESS) ---
# =========================================================================
print("Configurando o navegador para rodar em modo headless...")
options = Options()
options.add_argument("--headless")
options.add_argument("--window-size=1920,1080") # Define um tamanho de janela para evitar problemas com layout responsivo
options.add_argument("--no-sandbox") # Necessário para rodar como root/admin em alguns ambientes
options.add_argument("--disable-dev-shm-usage") # Resolve alguns problemas de crash em containers/servidores
options.add_argument("--log-level=3") # Deixa o log do navegador mais limpo no terminal

# Usando Service e WebDriverManager para gerenciar o chromedriver automaticamente
service = ChromeService(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=options)
print("Navegador iniciado em segundo plano.")
# =========================================================================


def login_microsoft(driver):
    """Faz login automaticamente na conta da Microsoft."""
    driver.get("https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=4765445b-32c6-49b0-83e6-1d93765276ca&redirect_uri=https%3A%2F%2Fwww.office.com%2Flandingv2&response_type=code%20id_token&scope=openid%20profile%20https%3A%2F%2Fwww.office.com%2Fv2%2FOfficeHome.All&response_mode=form_post&nonce=638756513104447175.ZGMwNWRiNjYtZWU1YS00ZTMzLWJkNzktNGE4MWNiYzFkMWRmOWU1NzY2Y2UtZGZlZi00MTYyLWFiZGQtYmZhNmI0NDhiZmQy&ui_locales=pt-BR&mkt=pt-BR&client-request-id=8f05b0fa-350c-4c62-8ab2-01b9069efec6&state=_BtTrhnaNhLqlC7zOvdttcN7pLU_X_ZR49AdFgW8wi_YHmgQC0cxYm9_P62aJ4tfksGP_ytkvRkCXnzQRdtaA1LZjL2SKTwOZPWKS8Fp0_AxVZqNTYo2Gv9jJhjxtpimyXfpLX45IqkRw_64IvMK98Tacbmx1M71odF0sH1UXnW9fOT4pbyNEo1ya970RLRDmGVMgRmfcI7cv9w-MSwKmOSbtVWu8JH-yfPQ8bxdrgOklKnrm72BYvQjs7g_F0wit4bwS_x-0Fnf4t3Rv-K9MAUK3SB6vUS2UaGlXhihl4g&x-client-SKU=ID_NET8_0&x-client-ver=7.5.1.0")

    try:
        wait = WebDriverWait(driver, 10)

        # Aguarda e preenche o e-mail
        email_input = wait.until(EC.presence_of_element_located((By.NAME, "loginfmt")))
        email_input.send_keys(EMAIL)
        email_input.send_keys(Keys.RETURN)
        time.sleep(2)

        # Aguarda e preenche a senha
        senha_input = wait.until(EC.presence_of_element_located((By.NAME, "passwd")))
        senha_input.send_keys(SENHA)
        senha_input.send_keys(Keys.RETURN)
        time.sleep(2)

        # Verifica se existe o botão "Sim" para manter logado
        try:
            sim_button = wait.until(EC.element_to_be_clickable((By.ID, "idSIButton9")))
            sim_button.click()
            time.sleep(2)
        except:
            pass  # Se não aparecer, segue o fluxo normalmente

        print("Login realizado com sucesso!")
        return True # Adicionado para melhor controle de fluxo

    except Exception as e:
        print(f"Erro no login: {e}")
        driver.save_screenshot("erro_login.png")
        print("Screenshot 'erro_login.png' salvo para depuração.")
        return False # Adicionado para melhor controle de fluxo

# Sua função `verificar_erros` foi integrada no final da função de login para maior clareza.
# Sua função `esperar_sincronizacao` está perfeita, não precisa de mudanças.
def esperar_sincronizacao(driver, timeout=60):
    """
    Espera até que a planilha esteja sincronizada com o Forms,
    verificando dentro de iframes se necessário.
    """
    try:
        # Primeiro tenta encontrar a mensagem diretamente na página
        try:
            WebDriverWait(driver, 5).until( # Tempo curto para a verificação principal
                EC.presence_of_element_located(
                    (By.XPATH, "//span[contains(@class, 'messageTitle') and contains(., 'Sincronizado com Forms')]")
                )
            )
            print("  Sincronização detectada na página principal!")
            return True
        except:
            pass

        # Se não encontrar, procura por iframes e verifica dentro deles
        frames = driver.find_elements(By.TAG_NAME, "iframe")
        if frames:
            print(f"  Encontrados {len(frames)} iframes. Verificando dentro deles...")

        for index, frame in enumerate(frames):
            try:
                driver.switch_to.frame(frame)
                try:
                    WebDriverWait(driver, 5).until( # Tempo curto para a verificação no iframe
                        EC.presence_of_element_located(
                            (By.XPATH, "//span[contains(@class, 'messageTitle') and contains(., 'Sincronizado com Forms')]")
                        )
                    )
                    print(f"  Sincronização detectada dentro do iframe #{index}!")
                    driver.switch_to.default_content()  # Volta para o contexto principal
                    return True
                except:
                    driver.switch_to.default_content()  # Volta se não encontrar
                    continue
            except Exception as e:
                #print(f"  Erro ao acessar iframe #{index}: {e}")
                driver.switch_to.default_content()

        return False

    except Exception as e:
        print(f"  Erro geral ao verificar sincronização: {e}")
        return False

# --- Estrutura principal com try/finally para garantir que o driver feche ---
try:
    if login_microsoft(driver):
        # Acessa cada categoria de URLs
        for categoria, lista_urls in urls.items():
            print(f"\nCategoria: {categoria}")
            
            for nome_planilha, url in lista_urls:
                print(f"  Acessando: {nome_planilha}")
                driver.get(url)
                time.sleep(5)  # Espera inicial para carregar
                
                if esperar_sincronizacao(driver):
                    print("  Status: OK")
                else:
                    print(f"  Status: AVISO - Mensagem de sincronização não encontrada.")
                    # Fallback com tempo fixo pode ser adicionado aqui se necessário
                    # time.sleep(15)

    print("\nTodas as planilhas foram processadas!")

except Exception as e:
    print(f"\n[ERRO GERAL] O script falhou inesperadamente: {e}")
    driver.save_screenshot("erro_geral.png")
    print("Screenshot 'erro_geral.png' salvo para depuração.")

finally:
    # Fecha o navegador
    print("\nFechando o navegador...")
    driver.quit()