import tkinter as tk
from tkinter import ttk, messagebox
from playwright.sync_api import sync_playwright
import os
import re
import urllib.request
import random
import base64
import pydub
import shutil
from pathlib import Path
from speech_recognition import Recognizer, AudioFile

# Para OCR do captcha matemático do TRT15
try:
    import pytesseract
    from PIL import Image
    # Define o caminho do Tesseract diretamente (evita problemas com PATH)
    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
    OCR_DISPONIVEL = True
except ImportError:
    OCR_DISPONIVEL = False
    print("AVISO: pytesseract ou Pillow não instalado. Instale com:")
    print("  pip install pytesseract Pillow")
    print("  E instale o Tesseract OCR: https://github.com/tesseract-ocr/tesseract")


# =========================
# FUNÇÃO PARA MOVER DOWNLOAD
# =========================

def mover_e_renomear_download(download, nome_arquivo_final, pasta_destino):
    """
    Move o arquivo baixado do Playwright para a pasta correta com nome correto.
    Também cria uma cópia na pasta Downloads do Windows como backup.
    
    Args:
        download: Objeto do download do Playwright
        nome_arquivo_final: Nome final do arquivo (ex: certidao_TRT15_xxx.pdf)
        pasta_destino: Pasta onde salvar o arquivo
    
    Returns:
        Caminho do arquivo salvo ou None se falhar
    """
    try:
        # Caminho absoluto da pasta de destino
        pasta_destino_abs = os.path.abspath(pasta_destino)
        os.makedirs(pasta_destino_abs, exist_ok=True)
        
        # Caminho final do arquivo (com caminho absoluto)
        caminho_final = os.path.join(pasta_destino_abs, nome_arquivo_final)
        
        # Move o arquivo da pasta temporária do Playwright para o destino
        download.save_as(caminho_final)
        print(f"✓ Arquivo salvo em: {caminho_final}")
        
        # Cria backup na pasta Downloads do Windows
        downloads_dir = str(Path.home() / "Downloads")
        caminho_backup = os.path.join(downloads_dir, nome_arquivo_final)
        
        if os.path.exists(caminho_final):
            shutil.copy2(caminho_final, caminho_backup)
            print(f"✓ Backup criado em: {caminho_backup}")
        
        return caminho_final
        
    except Exception as e:
        print(f"✗ Erro ao mover arquivo: {e}")
        return None


# =========================
# CLASSE PARA RESOLVER reCAPTCHA VIA ÁUDIO
# =========================

class RecaptchaSolver:
    """
    Resolve reCAPTCHA v2 automaticamente usando o desafio de áudio.
    Fluxo:
      1. Clica no checkbox do reCAPTCHA
      2. Se não resolveu direto, clica no botão de áudio
      3. Baixa o MP3 do desafio
      4. Converte para WAV e transcreve com Google Speech Recognition
      5. Preenche a resposta e verifica
      6. Repete até 5 vezes se necessário
    """

    def __init__(self, page):
        self.page = page
        self.main_frame = None
        self.recaptcha_frame = None

    def _delay(self, min_sec=1, max_sec=3):
        """Espera aleatória para simular comportamento humano."""
        self.page.wait_for_timeout(random.randint(min_sec, max_sec) * 1000)

    def _encontrar_frames(self):
        """Localiza os iframes do reCAPTCHA na página."""
        self.page.wait_for_selector("iframe[title='reCAPTCHA']", timeout=15000)
        nome_frame_checkbox = self.page.locator(
            "//iframe[@title='reCAPTCHA']"
        ).get_attribute("name")
        self.recaptcha_frame = self.page.frame(name=nome_frame_checkbox)

    def _clicar_checkbox(self):
        """Clica no checkbox do reCAPTCHA."""
        self.recaptcha_frame.click("//div[@class='recaptcha-checkbox-border']")
        self._delay(2, 4)

    def _ja_resolvido(self):
        """Verifica se o reCAPTCHA já foi resolvido (checkmark verde)."""
        try:
            anchor = self.recaptcha_frame.locator("//span[@id='recaptcha-anchor']")
            return anchor.get_attribute("aria-checked") != "false"
        except Exception:
            return False

    def _abrir_desafio_audio(self):
        """Clica no botão de áudio no desafio do reCAPTCHA."""
        self.page.wait_for_selector(
            "iframe[src*='google.com/recaptcha/api2/bframe']", timeout=10000
        )
        nome_frame_desafio = self.page.locator(
            "//iframe[contains(@src,'google.com/recaptcha/api2/bframe')]"
        ).get_attribute("name")
        self.main_frame = self.page.frame(name=nome_frame_desafio)
        self.main_frame.click("id=recaptcha-audio-button")
        self._delay(2, 3)

    def _resolver_audio(self):
        """Baixa o áudio, transcreve e preenche a resposta."""
        link_audio = self.main_frame.locator(
            "//a[@class='rc-audiochallenge-tdownload-link']"
        ).get_attribute("href")

        caminho_mp3 = os.path.join(os.getcwd(), "recaptcha_audio.mp3")
        caminho_wav = os.path.join(os.getcwd(), "recaptcha_audio.wav")

        urllib.request.urlretrieve(link_audio, caminho_mp3)
        print("  Áudio baixado com sucesso.")

        som = pydub.AudioSegment.from_mp3(caminho_mp3)
        som.export(caminho_wav, format="wav")
        print("  Áudio convertido para WAV.")

        recognizer = Recognizer()
        with AudioFile(caminho_wav) as source:
            audio_data = recognizer.record(source)

        texto = recognizer.recognize_google(audio_data, language="en-US")
        print(f"  Texto transcrito: {texto}")

        self.main_frame.fill("id=audio-response", texto)
        self._delay(1, 2)

        self.main_frame.click("id=recaptcha-verify-button")
        self._delay(3, 5)

        self._limpar_arquivos(caminho_mp3, caminho_wav)
        return texto

    def _limpar_arquivos(self, *caminhos):
        """Remove arquivos temporários de áudio."""
        for caminho in caminhos:
            try:
                if os.path.exists(caminho):
                    os.remove(caminho)
            except Exception:
                pass

    def resolver(self):
        """
        Método principal: tenta resolver o reCAPTCHA automaticamente.
        Retorna True se resolveu com sucesso, False caso contrário.
        """
        print("=" * 50)
        print("INICIANDO RESOLUÇÃO AUTOMÁTICA DO reCAPTCHA")
        print("=" * 50)

        try:
            print("[1/3] Localizando reCAPTCHA...")
            self._encontrar_frames()

            print("[2/3] Clicando no checkbox...")
            self._clicar_checkbox()

            if self._ja_resolvido():
                print("reCAPTCHA resolvido apenas com o clique!")
                return True

            print("[3/3] Abrindo desafio de áudio...")
            self._abrir_desafio_audio()

            tentativas = 0
            max_tentativas = 5

            while tentativas < max_tentativas:
                tentativas += 1
                print(f"\n--- Tentativa {tentativas}/{max_tentativas} ---")

                try:
                    self._resolver_audio()

                    if self._ja_resolvido():
                        print("\nreCAPTCHA RESOLVIDO COM SUCESSO!")
                        return True
                    else:
                        print("  Resposta incorreta, tentando novamente...")
                        try:
                            self.main_frame.click("id=recaptcha-reload-button")
                            self._delay(2, 4)
                        except Exception:
                            pass

                except Exception as e:
                    print(f"  Erro na tentativa {tentativas}: {e}")
                    try:
                        self.main_frame.click("id=recaptcha-reload-button")
                        self._delay(2, 4)
                    except Exception:
                        pass

            print("\nNão foi possível resolver o reCAPTCHA após todas as tentativas.")
            return False

        except Exception as e:
            print(f"Erro geral ao resolver reCAPTCHA: {e}")
            return False


# =========================
# FUNÇÃO PARA SALVAR PDF VIA CDP
# =========================

def salvar_pdf_via_cdp(page, context, caminho_arquivo):
    """
    Salva a página atual como PDF usando o Chrome DevTools Protocol (CDP).
    Funciona com headless=False, sem abrir diálogo de impressão.
    """
    print("Gerando PDF via CDP (sem diálogo de impressão)...")

    cdp_session = context.new_cdp_session(page)

    resultado = cdp_session.send("Page.printToPDF", {
        "landscape": False,
        "displayHeaderFooter": False,
        "printBackground": True,
        "preferCSSPageSize": True,
        "paperWidth": 8.27,
        "paperHeight": 11.69,
        "marginTop": 0.4,
        "marginBottom": 0.4,
        "marginLeft": 0.4,
        "marginRight": 0.4,
    })

    pdf_bytes = base64.b64decode(resultado["data"])

    with open(caminho_arquivo, "wb") as f:
        f.write(pdf_bytes)

    cdp_session.detach()

    print(f"PDF salvo com sucesso em: {caminho_arquivo}")
    return True


# =========================
# FUNÇÃO TRT2
# =========================

def emitir_trt2(documento):

    pasta_download = "certidoes_baixadas"
    os.makedirs(pasta_download, exist_ok=True)

    with sync_playwright() as p:

        browser = p.chromium.launch(headless=False)
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                       "AppleWebKit/537.36 (KHTML, like Gecko) "
                       "Chrome/120.0.0.0 Safari/537.36"
        )
        page = context.new_page()

        print("Acessando TRT2...")
        page.goto("https://pje.trt2.jus.br/certidoes/trabalhista/emissao", timeout=60000)

        page.wait_for_selector("mat-radio-group", timeout=60000)
        print("Formulário carregado.")

        if len(documento) > 11:
            print("Selecionando Raiz do CNPJ...")
            page.get_by_text("Raiz do CNPJ", exact=True).click()
        else:
            print("Selecionando CPF...")
            page.get_by_text("CPF", exact=True).click()

        page.wait_for_timeout(1000)

        page.locator("input[type='text']").first.fill(documento)
        print("Documento preenchido.")

        # ================= CAPTCHA AUTOMÁTICO =================

        print("\nIniciando resolução automática do reCAPTCHA...")

        solver = RecaptchaSolver(page)
        captcha_resolvido = solver.resolver()

        if not captcha_resolvido:
            print("\n⚠ Resolução automática falhou.")
            print("Por favor, resolva o captcha manualmente na janela do navegador.")
            page.wait_for_timeout(30000)

        # ================= EMITIR =================

        print("\nAguardando botão liberar...")

        page.wait_for_function("""
        () => {
            const btn = [...document.querySelectorAll("button")]
                .find(b => b.innerText.toLowerCase().includes("emitir"));
            return btn && !btn.disabled;
        }
        """, timeout=120000)

        print("Emitindo certidão...")
        page.locator("button:has-text('Emitir')").click()

        # ================= ESPERA BOTÃO IMPRIMIR =================

        print("Aguardando botão IMPRIMIR aparecer e habilitar...")

        page.wait_for_function("""
        () => {
            const btn = [...document.querySelectorAll("button")]
                .find(b => b.innerText.toLowerCase().includes("imprimir"));
            return btn && !btn.disabled;
        }
        """, timeout=120000)

        print("Botão IMPRIMIR habilitado.")

        page.wait_for_timeout(3000)

        # ================= SALVAR PDF AUTOMATICAMENTE =================

        nome_limpo = documento.replace(".", "").replace("/", "").replace("-", "")
        nome_arquivo = f"certidao_TRT2_{nome_limpo}.pdf"
        caminho_final = os.path.join(pasta_download, nome_arquivo)

        salvar_pdf_via_cdp(page, context, caminho_final)

        print("=" * 50)
        print("CERTIDÃO TRT2 SALVA COM SUCESSO!")
        print(f"Arquivo: {caminho_final}")
        print(f"Pasta: {os.path.abspath(pasta_download)}")
        print("=" * 50)

        browser.close()


# =========================
# FUNÇÃO TRT9
# =========================

def emitir_trt9(documento):

    pasta_download = "certidoes_baixadas"
    os.makedirs(pasta_download, exist_ok=True)

    with sync_playwright() as p:

        browser = p.chromium.launch(headless=False)
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                       "AppleWebKit/537.36 (KHTML, like Gecko) "
                       "Chrome/120.0.0.0 Safari/537.36"
        )
        page = context.new_page()

        print("Acessando TRT9...")
        page.goto("https://pje.trt9.jus.br/certidoes/trabalhista/emissao", timeout=60000)

        page.wait_for_selector("mat-radio-group", timeout=60000)
        print("Formulário carregado.")

        if len(documento) > 11:
            print("Selecionando Raiz do CNPJ...")
            page.get_by_text("Raiz do CNPJ", exact=True).click()
        else:
            print("Selecionando CPF...")
            page.get_by_text("CPF", exact=True).click()

        page.wait_for_timeout(1000)

        page.locator("input[type='text']").first.fill(documento)
        print("Documento preenchido.")

        # ================= CAPTCHA AUTOMÁTICO =================

        print("\nIniciando resolução automática do reCAPTCHA...")

        solver = RecaptchaSolver(page)
        captcha_resolvido = solver.resolver()

        if not captcha_resolvido:
            print("\n⚠ Resolução automática falhou.")
            print("Por favor, resolva o captcha manualmente na janela do navegador.")
            page.wait_for_timeout(30000)

        # ================= EMITIR =================

        print("\nAguardando botão liberar...")

        page.wait_for_function("""
        () => {
            const btn = [...document.querySelectorAll("button")]
                .find(b => b.innerText.toLowerCase().includes("emitir"));
            return btn && !btn.disabled;
        }
        """, timeout=120000)

        print("Emitindo certidão...")
        page.locator("button:has-text('Emitir')").click()

        # ================= ESPERA BOTÃO IMPRIMIR =================

        print("Aguardando botão IMPRIMIR aparecer e habilitar...")

        page.wait_for_function("""
        () => {
            const btn = [...document.querySelectorAll("button")]
                .find(b => b.innerText.toLowerCase().includes("imprimir"));
            return btn && !btn.disabled;
        }
        """, timeout=120000)

        print("Botão IMPRIMIR habilitado.")

        page.wait_for_timeout(3000)

        # ================= SALVAR PDF AUTOMATICAMENTE =================

        nome_limpo = documento.replace(".", "").replace("/", "").replace("-", "")
        nome_arquivo = f"certidao_TRT9_{nome_limpo}.pdf"
        caminho_final = os.path.join(pasta_download, nome_arquivo)

        salvar_pdf_via_cdp(page, context, caminho_final)

        print("=" * 50)
        print("CERTIDÃO TRT9 SALVA COM SUCESSO!")
        print(f"Arquivo: {caminho_final}")
        print(f"Pasta: {os.path.abspath(pasta_download)}")
        print("=" * 50)

        browser.close()


# =========================
# FUNÇÃO TRT1
# =========================

def emitir_trt1(documento):

    pasta_download = "certidoes_baixadas"
    os.makedirs(pasta_download, exist_ok=True)

    with sync_playwright() as p:

        browser = p.chromium.launch(headless=False)
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                       "AppleWebKit/537.36 (KHTML, like Gecko) "
                       "Chrome/120.0.0.0 Safari/537.36"
        )
        page = context.new_page()

        print("Acessando TRT1...")
        page.goto("https://pje.trt1.jus.br/certidoes/trabalhista/emissao", timeout=60000)

        page.wait_for_selector("mat-radio-group", timeout=60000)
        print("Formulário carregado.")

        if len(documento) > 11:
            print("Selecionando Raiz do CNPJ...")
            page.get_by_text("Raiz do CNPJ", exact=True).click()
        else:
            print("Selecionando CPF...")
            page.get_by_text("CPF", exact=True).click()

        page.wait_for_timeout(1000)

        page.locator("input[type='text']").first.fill(documento)
        print("Documento preenchido.")

        # ================= CAPTCHA AUTOMÁTICO =================

        print("\nIniciando resolução automática do reCAPTCHA...")

        solver = RecaptchaSolver(page)
        captcha_resolvido = solver.resolver()

        if not captcha_resolvido:
            print("\n⚠ Resolução automática falhou.")
            print("Por favor, resolva o captcha manualmente na janela do navegador.")
            page.wait_for_timeout(30000)

        # ================= EMITIR =================

        print("\nAguardando botão liberar...")

        page.wait_for_function("""
        () => {
            const btn = [...document.querySelectorAll("button")]
                .find(b => b.innerText.toLowerCase().includes("emitir"));
            return btn && !btn.disabled;
        }
        """, timeout=120000)

        print("Emitindo certidão...")
        page.locator("button:has-text('Emitir')").click()

        # ================= ESPERA BOTÃO IMPRIMIR =================

        print("Aguardando botão IMPRIMIR aparecer e habilitar...")

        page.wait_for_function("""
        () => {
            const btn = [...document.querySelectorAll("button")]
                .find(b => b.innerText.toLowerCase().includes("imprimir"));
            return btn && !btn.disabled;
        }
        """, timeout=120000)

        print("Botão IMPRIMIR habilitado.")

        page.wait_for_timeout(3000)

        # ================= SALVAR PDF AUTOMATICAMENTE =================

        nome_limpo = documento.replace(".", "").replace("/", "").replace("-", "")
        nome_arquivo = f"certidao_TRT1_{nome_limpo}.pdf"
        caminho_final = os.path.join(pasta_download, nome_arquivo)

        salvar_pdf_via_cdp(page, context, caminho_final)

        print("=" * 50)
        print("CERTIDÃO TRT1 SALVA COM SUCESSO!")
        print(f"Arquivo: {caminho_final}")
        print(f"Pasta: {os.path.abspath(pasta_download)}")
        print("=" * 50)

        browser.close()


# =========================
# FUNÇÃO TRF1
# =========================
def emitir_trf1(documento, email):
    """
    Solicita a certidão unificada do TRF1 (não baixa PDF).
    Preenche: tipo cível, todos os órgãos, CPF/CNPJ, e-mail e confirma.
    """

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                       "AppleWebKit/537.36 (KHTML, like Gecko) "
                       "Chrome/120.0.0.0 Safari/537.36"
        )
        page = context.new_page()

        print("Acessando TRF1 unificado...")
        page.goto("https://certidao-unificada.cjf.jus.br/#/solicitacao-certidao", timeout=60000)
        page.wait_for_selector("form", timeout=60000)
        print("Formulário carregado.")

        # Seleciona tipo de certidão Cível (PrimeNG dropdown)
        # clicar no botão de trigger garante abertura da lista
        trigger = page.locator("div.p-dropdown-trigger")
        trigger.scroll_into_view_if_needed()
        trigger.wait_for(state="visible", timeout=10000)
        trigger.click()
        # aguarda menu aparecer e escolhe Cível
        opt = page.get_by_role("option", name="Cível")
        opt.wait_for(state="visible", timeout=10000)
        opt.click()

        # órgãos já vêm todos selecionados por padrão, não precisa mexer

        # Preenche CPF ou CNPJ
        # os inputs de rádio do PrimeNG podem estar escondidos; clicar no label funciona melhor
        if len(documento) > 11:
            print("Selecionando CNPJ via label...")
            try:
                page.get_by_label("CNPJ").click()
            except Exception:
                # fallback para clique via JS caso o label também não seja clicável
                page.evaluate("document.getElementById('cnpj').click()")
        else:
            print("Selecionando CPF via label...")
            # .first garante que pegamos o label do rádio, não o label do campo de texto
            try:
                page.get_by_label("CPF").first.click()
            except Exception:
                page.evaluate("document.getElementById('cpf').click()")
        page.locator("input[name='cpfCnpj']").fill(documento)

        # Preenche e-mail e confirmação
        page.fill("#email", email)
        page.fill("#emailConfirmacao", email)

        # Envia solicitação
        page.get_by_role("button", name="Solicitar certidão").click()
        print("Solicitação enviada. A certidão deverá chegar no e-mail informado.")
        page.wait_for_timeout(5000)
        browser.close()


# =========================
# FUNÇÃO TRT15
# =========================

def emitir_trt15(documento):
    """
    Emite certidão do TRT15 (Campinas).
    Fluxo:
      1. Acessa diretamente a URL do CEAT (evita problemas com iframe)
      2. Preenche CPF/CNPJ usando o ID exato do campo
      3. Resolve o captcha matemático (soma simples)
      4. Clica em "Emitir Certidão"
      5. Clica em "Imprimir Certidão" (que baixa o PDF automaticamente)
    """

    pasta_download = "certidoes_baixadas"
    os.makedirs(pasta_download, exist_ok=True)

    # IDs exatos dos campos (obtidos via inspeção da página)
    ID_CPF = "certidaoActionForm:j_id23:doctoPesquisa"
    ID_CAPTCHA = "certidaoActionForm:j_id51:verifyCaptcha"
    ID_EMITIR = "certidaoActionForm:certidaoActionEmitir"

    with sync_playwright() as p:

        browser = p.chromium.launch(headless=False)
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                       "AppleWebKit/537.36 (KHTML, like Gecko) "
                       "Chrome/120.0.0.0 Safari/537.36",
            accept_downloads=True,
        )
        page = context.new_page()

        nome_limpo = documento.replace(".", "").replace("/", "").replace("-", "")

        # ================= ACESSAR CEAT DIRETAMENTE =================
        # Acessa a URL do CEAT diretamente, evitando problemas com iframe

        print("Acessando TRT15 (CEAT)...")
        page.goto(
            "https://ceat.trt15.jus.br/ceat/certidaoAction.seam",
            timeout=60000,
        )
        page.wait_for_timeout(3000)
        print("Página carregada.")

        # ================= PREENCHER CPF/CNPJ =================

        print("Preenchendo CPF/CNPJ...")
        campo_cpf = page.locator(f"[id='{ID_CPF}']")
        campo_cpf.wait_for(state="visible", timeout=15000)
        campo_cpf.fill(documento)
        print(f"Documento preenchido: {documento}")

        page.wait_for_timeout(1000)

        # ================= RESOLVER CAPTCHA MATEMÁTICO =================

        print("\nResolvendo captcha matemático (imagem)...")

        captcha_resolvido = False

        try:
            # O captcha é uma IMAGEM com a conta matemática.
            # Precisamos tirar screenshot da imagem e usar OCR para ler.
            img_captcha = page.locator("img[src*='captcha']")
            img_captcha.wait_for(state="visible", timeout=10000)
            print("Imagem do captcha encontrada.")

            # Salva screenshot da imagem do captcha
            caminho_captcha = os.path.join(os.getcwd(), "captcha_trt15.png")
            img_captcha.screenshot(path=caminho_captcha)
            print("Screenshot do captcha salvo.")

            if OCR_DISPONIVEL:
                # Lê a imagem com OCR
                imagem = Image.open(caminho_captcha)

                # Pré-processamento para melhorar OCR:
                # Converte para escala de cinza e aumenta o tamanho
                imagem = imagem.convert("L")  # Escala de cinza
                largura, altura = imagem.size
                imagem = imagem.resize((largura * 3, altura * 3), Image.LANCZOS)

                # OCR com configuração para números e operadores
                texto_captcha = pytesseract.image_to_string(
                    imagem,
                    config="--psm 7 -c tessedit_char_whitelist=0123456789+-="
                )
                texto_captcha = texto_captcha.strip()
                print(f"Texto lido pelo OCR: '{texto_captcha}'")

                # Procura padrão "X + Y" no texto
                match = re.search(r'(\d+)\s*[\+\+]\s*(\d+)', texto_captcha)

                if match:
                    num1 = int(match.group(1))
                    num2 = int(match.group(2))
                    resultado = num1 + num2
                    print(f"Conta encontrada: {num1} + {num2} = {resultado}")

                    campo_captcha = page.locator(f"[id='{ID_CAPTCHA}']")
                    campo_captcha.wait_for(state="visible", timeout=10000)
                    campo_captcha.fill(str(resultado))
                    print(f"Resposta do captcha preenchida: {resultado}")
                    captcha_resolvido = True
                else:
                    # Tenta avaliar a expressão diretamente
                    try:
                        # Limpa o texto e tenta extrair números
                        numeros = re.findall(r'\d+', texto_captcha)
                        if len(numeros) >= 2:
                            num1 = int(numeros[0])
                            num2 = int(numeros[1])
                            resultado = num1 + num2
                            print(f"Números extraídos: {num1} + {num2} = {resultado}")

                            campo_captcha = page.locator(f"[id='{ID_CAPTCHA}']")
                            campo_captcha.wait_for(state="visible", timeout=10000)
                            campo_captcha.fill(str(resultado))
                            print(f"Resposta do captcha preenchida: {resultado}")
                            captcha_resolvido = True
                    except Exception:
                        pass

                # Limpa arquivo temporário
                try:
                    os.remove(caminho_captcha)
                except Exception:
                    pass

            else:
                print("OCR não disponível (pytesseract não instalado).")

        except Exception as e:
            print(f"Erro ao resolver captcha por OCR: {e}")

        if not captcha_resolvido:
            print("\nNão foi possível resolver o captcha automaticamente.")
            print("Preencha o captcha manualmente na janela do navegador.")
            print("Você tem 30 segundos...")
            page.wait_for_timeout(30000)

        page.wait_for_timeout(1000)

        # ================= EMITIR CERTIDÃO =================

        print("\nClicando em 'Emitir Certidão'...")
        botao_emitir = page.locator(f"[id='{ID_EMITIR}']")
        botao_emitir.wait_for(state="visible", timeout=10000)
        botao_emitir.click()
        print("Botão 'Emitir Certidão' clicado!")

        # ================= AGUARDAR E BAIXAR PDF =================

        print("\nAguardando certidão ser gerada...")
        page.wait_for_timeout(5000)

        print("Procurando botão 'Imprimir Certidão'...")

        # Tenta encontrar o botão de imprimir
        botao_imprimir = None
        seletores_imprimir = [
            "input[type='submit'][value*='Imprimir']",
            "input[type='button'][value*='Imprimir']",
            "button:has-text('Imprimir')",
            "a:has-text('Imprimir')",
            "input[value*='Imprimir']",
        ]

        # Espera até 30 segundos pelo botão de imprimir aparecer
        for tentativa in range(6):
            for seletor in seletores_imprimir:
                try:
                    botao = page.locator(seletor).first
                    if botao.is_visible(timeout=2000):
                        botao_imprimir = botao
                        print(f"Botão 'Imprimir Certidão' encontrado com seletor: {seletor}")
                        break
                except Exception:
                    continue

            if botao_imprimir:
                break

            print(f"  Aguardando botão imprimir... ({(tentativa + 1) * 5}s)")
            page.wait_for_timeout(5000)

        if botao_imprimir:
            # Captura o download ao clicar no botão
            nome_arquivo = f"certidao_TRT15_{nome_limpo}.pdf"

            try:
                with page.expect_download(timeout=30000) as download_info:
                    botao_imprimir.click()
                    print("Botão 'Imprimir Certidão' clicado! Aguardando download...")

                download = download_info.value
                
                # Move o arquivo para a pasta correta com nome correto
                caminho_final = mover_e_renomear_download(download, nome_arquivo, pasta_download)
                
                if caminho_final:
                    print("=" * 50)
                    print("CERTIDÃO TRT15 SALVA COM SUCESSO!")
                    print(f"Arquivo: {caminho_final}")
                    print(f"Pasta: {os.path.abspath(pasta_download)}")
                    print("=" * 50)
                else:
                    print("Erro ao mover o arquivo.")

            except Exception as e:
                print(f"Download direto não capturado: {e}")
                print("Tentando salvar via CDP...")

                # Fallback: salva via CDP caso o download não funcione
                try:
                    caminho_final = os.path.join(os.path.abspath(pasta_download), nome_arquivo)
                    salvar_pdf_via_cdp(page, context, caminho_final)
                    print("=" * 50)
                    print("CERTIDÃO TRT15 SALVA COM SUCESSO (via CDP)!")
                    print(f"Arquivo: {caminho_final}")
                    print(f"Pasta: {os.path.abspath(pasta_download)}")
                    print("=" * 50)
                except Exception as e2:
                    print(f"Erro no fallback CDP: {e2}")
                    print("Verifique a pasta de Downloads do navegador.")
        else:
            print("ERRO: Botão 'Imprimir Certidão' não encontrado!")
            print("Tente clicar manualmente na janela do navegador.")
            page.wait_for_timeout(30000)

        browser.close()


# =========================
# INTERFACE
# =========================

def executar():

    tribunal = combo_tribunal.get()
    trt = combo_trt.get()
    documento = entry_documento.get().strip()

    if not documento:
        messagebox.showerror("Erro", "Informe CPF ou CNPJ.")
        return

    if tribunal == "TRT":
        if trt == "TRT1":
            emitir_trt1(documento)
        elif trt == "TRT2":
            emitir_trt2(documento)
        elif trt == "TRT9":
            emitir_trt9(documento)
        elif trt == "TRT15":
            emitir_trt15(documento)
        else:
            messagebox.showinfo("Aviso", "Ainda não implementado para este tribunal.")
    elif tribunal == "TRF":
        # preciso de e-mail
        email = entry_email.get().strip()
        if not email:
            messagebox.showerror("Erro", "Informe um e-mail para receber a certidão.")
            return
        emitir_trf1(documento, email)
    else:
        messagebox.showinfo("Aviso", "Ainda não implementado para este tribunal.")


def atualizar_trt(event):

    escolha = combo_tribunal.get()
    if escolha == "TRT":
        combo_trt["values"] = [f"TRT{i}" for i in range(1, 25)]
        combo_trt.current(1)  # TRT2 por padrão
        combo_trt.config(state="readonly")
        # email não é necessário para TRT
        label_email.pack_forget()
        entry_email.pack_forget()
    elif escolha == "TRF":
        combo_trt["values"] = ["TRF1"]
        combo_trt.current(0)
        combo_trt.config(state="readonly")
        # mostrar campo de e-mail para o TRF
        label_email.pack(pady=5)
        entry_email.pack()
    else:
        combo_trt.set("")
        combo_trt.config(state="disabled")
        label_email.pack_forget()
        entry_email.pack_forget()


# =========================
# UI
# =========================

root = tk.Tk()
root.title("Emissor de Certidão")
root.geometry("400x260")

tk.Label(root, text="Tribunal:").pack(pady=5)
combo_tribunal = ttk.Combobox(root, values=["TRT", "TJ", "TRF"], state="readonly")
combo_tribunal.pack()
combo_tribunal.bind("<<ComboboxSelected>>", atualizar_trt)

tk.Label(root, text="Selecione o TRT:").pack(pady=5)
combo_trt = ttk.Combobox(root, state="disabled")
combo_trt.pack()

tk.Label(root, text="CPF ou CNPJ:").pack(pady=5)
entry_documento = tk.Entry(root, width=30)
entry_documento.pack()

# campo de e-mail que só aparece quando o tribunal for TRF
label_email = tk.Label(root, text="E-mail para recebimento:")
entry_email = tk.Entry(root, width=30)
# inicialmente escondido


# botão principal
tk.Button(root, text="Emitir Certidão", command=executar).pack(pady=20)

root.mainloop()
