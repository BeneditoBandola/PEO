import os
import smtplib
from datetime import datetime
from zoneinfo import ZoneInfo
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from playwright.sync_api import sync_playwright

# ==========================================
# CONFIGURAÇÕES E VARIÁVEIS DE AMBIENTE
# ==========================================
USUARIO_PDV = os.environ.get("USUARIO_PDV")
SENHA_PDV = os.environ.get("SENHA_PDV")

# E-mails destinatários fixos e por filial
EMAILS_MEUS = [
    "beneditobandola@gmail.com",
    "benedito.bandola@minassal.com.br"
]

EMAILS_PROMOTORES = {
    "MINASSAL LTDA - POCOS DE CALDAS": ["pamelaalmeida5@hotmail.com"],
    "MINASSAL LTDA - SAO JOAO DA BOA VISTA": ["crbruno27123@gmail.com"],
    "MINASSAL LTDA - SAO JOSE DO RIO PRETO": ["saruetesjrp79@gmail.com"],
    "MINASSAL LTDA - JUIZ DE FORA": ["fernandaferreira_jf@yahoo.com.br", "madallareis66@gmail.com"]
}

INICIO_P8 = "2026-07-20"
CAMINHO_CSV_FINAL = "historico_p8_p9.csv"


# ==========================================
# 1. DOWNLOAD DOS DADOS DO PDV PET
# ==========================================
def baixar_dados_pdvpet():
    print(f"\n--- INICIANDO DOWNLOAD DO HISTÓRICO (P8 e P9) ---")
    
    if not USUARIO_PDV or not SENHA_PDV:
        print("❌ ERRO CRÍTICO: Variáveis USUARIO_PDV e SENHA_PDV não foram encontradas nos Secrets do GitHub!")
        return False

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=True,
            args=["--no-sandbox", "--disable-setuid-sandbox", "--disable-dev-shm-usage"]
        )
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            viewport={"width": 1280, "height": 800}
        )
        page = context.new_page()

        try:
            print("🔗 Acessando o site PDV Pet...")
            page.goto("https://www.pdvpet.com.br/", timeout=60000, wait_until="networkidle")

            print("🔑 Preenchendo dados de login...")
            page.fill('input[type="text"], input[name*="user"], input[name*="cpf"], input[name*="login"]', USUARIO_PDV)
            page.fill('input[type="password"]', SENHA_PDV)

            try:
                page.click('button[type="submit"], input[type="submit"], button:has-text("Entrar")', timeout=5000)
            except Exception:
                page.keyboard.press("Enter")

            print("⏳ Aguardando confirmação do login...")
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(3000)

            print("📋 Navegando até a aba de Questionários...")
            page.wait_for_selector('text=/Questionários|QUESTIONÁRIOS/i', timeout=60000)
            page.click('text=/Questionários|QUESTIONÁRIOS/i')
            
            page.wait_for_selector('#DataDe', timeout=60000)

            fuso_br = ZoneInfo("America/Sao_Paulo")
            data_hoje = datetime.now(fuso_br).strftime("%Y-%m-%d")

            print(f"📅 Preenchendo as datas: {INICIO_P8} até {data_hoje}...")
            page.fill('#DataDe', INICIO_P8)
            page.fill('#DataAte', data_hoje)
            page.click('button[type="submit"]:has-text("Buscar")')
            page.wait_for_timeout(8000)

            print("⏳ Baixando o relatório CSV...")
            with page.expect_download(timeout=60000) as download_info:
                page.click('button.btn-outline-success:has-text("Exportar")')
                try:
                    page.wait_for_selector('a:has-text("Abrir")', timeout=5000)
                    page.click('a:has-text("Abrir")')
                except Exception:
                    pass

            download_info.value.save_as(CAMINHO_CSV_FINAL)
            print("✅ CSV Baixado com Sucesso!")
            return True

        except Exception as e:
            print(f"❌ OCORREU UM ERRO DURANTE A NAVEGAÇÃO/DOWNLOAD: {e}")
            return False
        finally:
            browser.close()


# ==========================================
# 2. DISPARO DOS E-MAILS VIA SMTP
# ==========================================
def enviar_email(filial, caminho_pdf):
    usuario_smtp = os.environ.get("EMAIL_REMETENTE")
    senha_smtp = os.environ.get("SENHA_EMAIL")

    if not usuario_smtp or not senha_smtp:
        print(f"❌ ERRO CRÍTICO: Variáveis EMAIL_REMETENTE ou SENHA_EMAIL não configuradas nos Secrets do GitHub!")
        return

    senha_smtp = senha_smtp.replace(" ", "")

    destinatarios = list(set(EMAILS_MEUS + EMAILS_PROMOTORES.get(filial, [])))

    msg = MIMEMultipart()
    msg['From'] = usuario_smtp
    msg['To'] = ", ".join(destinatarios)
    msg['Subject'] = f"Relatório de Oportunidades e Auditoria - {filial}"

    corpo = (
        f"Olá,\n\n"
        f"Segue em anexo o relatório diário de oportunidades e auditoria de preços referente à filial {filial}.\n\n"
        f"Atenciosamente,\n"
        f"Automação PDV Pet\n"
    )
    msg.attach(MIMEText(corpo, 'plain'))

    if os.path.exists(caminho_pdf):
        with open(caminho_pdf, "rb") as f:
            part = MIMEApplication(f.read(), Name=os.path.basename(caminho_pdf))
            part['Content-Disposition'] = f'attachment; filename="{os.path.basename(caminho_pdf)}"'
            msg.attach(part)
    else:
        print(f"⚠️ PDF não encontrado em {caminho_pdf}, enviando e-mail sem anexo.")

    try:
        server = smtplib.SMTP_SSL('smtp.gmail.com', 465, timeout=30)
        server.login(usuario_smtp, senha_smtp)
        server.sendmail(usuario_smtp, destinatarios, msg.as_string())
        server.quit()
        print(f"📧 E-mail enviado com sucesso para a filial '{filial}': {destinatarios}")
    except Exception as e:
        print(f"❌ Erro ao enviar e-mail para {filial}: {e}")


# ==========================================
# EXECUÇÃO PRINCIPAL
# ==========================================
if __name__ == "__main__":
    sucesso_download = baixar_dados_pdvpet()
    
    if sucesso_download and os.path.exists(CAMINHO_CSV_FINAL):
        print("\n--- PROCESSANDO DADOS E GERANDO RELATÓRIOS ---")
        
        # Percorre cada filial para gerar e enviar seus relatórios
        for filial in EMAILS_PROMOTORES.keys():
            # Substitua a variável abaixo caso os seus PDFs sigam um padrão de nome diferente
            caminho_pdf = f"relatorio_{filial}.pdf" 
            
            print(f"📤 Iniciando envio para a filial: {filial}")
            enviar_email(filial, caminho_pdf)
            
    else:
        print("\n❌ Execução interrompida devido à falha no download do CSV.")
