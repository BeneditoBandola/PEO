import os
import smtplib
import pandas as pd
from datetime import datetime
from zoneinfo import ZoneInfo
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from playwright.sync_api import sync_playwright
import streamlit as st

from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

# Configuração da Página
st.set_page_config(page_title="Automação PDV Pet", page_icon="🐾", layout="centered")

st.title("🐾 Painel de Controle - PDV Pet")
st.write("Execute a extração e escolha para quem deseja enviar os relatórios e quais filtros aplicar.")

# ==========================================
# CONFIGURAÇÕES E VARIÁVEIS DE AMBIENTE
# ==========================================
USUARIO_PDV = os.environ.get("USUARIO_PDV")
SENHA_PDV = os.environ.get("SENHA_PDV")
EMAIL_REMETENTE = os.environ.get("EMAIL_REMETENTE")
SENHA_EMAIL = os.environ.get("SENHA_EMAIL")

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

PERIODO_ATUAL = "P10"
INICIO_P9  = "2026-08-10"
FIM_P9     = "2026-09-06"
INICIO_P10 = "2026-09-07"
FIM_P10    = "2026-10-04"

CONFIGURACOES_PERIODOS = {
    "P10": {
        "metas": {
            'MINASSAL LTDA - POCOS DE CALDAS': {'Small Bags': 107, 'Ponto Extra de Sachês/Petiscos': 28, 'Combos Virtuais Sachês': 10, 'Combos Virtuais Petiscos': 10, 'Sheba Cremoso - Leve 2 Pague 1': 10},
            'MINASSAL LTDA - SAO JOAO DA BOA VISTA': {'Small Bags': 334, 'Ponto Extra de Sachês/Petiscos': 53, 'Combos Virtuais Sachês': 10, 'Combos Virtuais Petiscos': 10, 'Sheba Cremoso - Leve 2 Pague 1': 10},
            'MINASSAL LTDA - SAO JOSE DO RIO PRETO': {'Small Bags': 226, 'Ponto Extra de Sachês/Petiscos': 46, 'Combos Virtuais Sachês': 10, 'Combos Virtuais Petiscos': 10, 'Sheba Cremoso - Leve 2 Pague 1': 10},
            'MINASSAL LTDA - JUIZ DE FORA': {'Small Bags': 83, 'Ponto Extra de Sachês/Petiscos': 19, 'Combos Virtuais Sachês': 10, 'Combos Virtuais Petiscos': 10, 'Sheba Cremoso - Leve 2 Pague 1': 10}
        }
    }
}

METAS = CONFIGURACOES_PERIODOS[PERIODO_ATUAL]["metas"]

PRECOS_MAXIMOS = {
    'KiteKat Adulto Mix de Carnes - Small Bags 0,9Kg': 11.90,
    'Whiskas Sabor CARNE - Small Bags 0,5Kg': 14.90,
    'Whiskas Sabor CARNE CASTRADOS - Small Bags 0,5Kg': 14.90,
    'Whiskas Sabor PEIXE - Small Bags 0,5Kg': 14.90,
    'Whiskas Sabor FILHOTE CARNE - Small Bags 0,5Kg': 14.90,
    'Champ Adulto Carne e Cereal - Small Bags 0,9Kg': 11.90,
    'Champ Filhotes Carne e Cereal - Small Bags 0,9Kg': 11.90,
    'Pedigree CARNE, FRANGO E CEREAIS - Small Bags 0,9Kg': 19.90,
    'Pedigree FILHOTES - Small Bags 0,9Kg': 19.90,
    'Pedigree RG CARNE E VEGETAIS - Small Bags 0,9Kg': 19.90,
    'Pedigree RP CARNE E VEGETAIS - Small Bags 0,9Kg': 19.90,
    'Pedigree NE AO LEITE - Small Bags 0,9Kg': 19.90,
    'Pedigree NE Carne - Small Bags 0,9Kg': 15.90,
    'Whiskas Sabor CARNE - Small Bags 0,9Kg': 20.90,
    'Whiskas Sabor CARNE CASTRADOS - Small Bags 0,9Kg': 20.90,
    'Whiskas Sabor FRANGO - Small Bags 0,9Kg': 20.90,
    'Whiskas Sabor PEIXE - Small Bags 0,9Kg': 20.90,
    'Whiskas Sabor PEIXE CASTRADOS - Small Bags 0,9Kg': 20.90,
    'Whiskas Sabor FILHOTE CARNE - Small Bags 0,9Kg': 20.90,
    'Pedigree CARNE, FRANGO E CEREAIS - Small Bags 2,7Kg': 54.90,
    'Pedigree FILHOTES - Small Bags 2,7Kg': 54.90,
    'Pedigree RG CARNE E VEGETAIS - Small Bags 2,7Kg': 54.90,
    'Pedigree RP CARNE E VEGETAIS - Small Bags 2,7Kg': 54.90,
}

PASTA_PROJETO = os.path.dirname(os.path.abspath(__file__))
CAMINHO_CSV_FINAL = os.path.join(PASTA_PROJETO, "historico_p9_p10.csv")

# ==========================================
# PAINEL LATERAL DE FILTROS E OPÇÕES
# ==========================================
st.sidebar.header("Filtros de Envio e Execução")

filial_selecionada = st.sidebar.selectbox(
    "Filial para Processar:", 
    ["Todas"] + list(METAS.keys())
)

enviar_apenas_para_mim = st.sidebar.checkbox("Enviar apenas para o meu e-mail (Teste)", value=True)

incluir_oportunidades = st.sidebar.checkbox("Incluir Seção de Oportunidades", value=True)
incluir_precos = st.sidebar.checkbox("Incluir Seção de Preços Máximos", value=True)

# ==========================================
# 1. DOWNLOAD DOS DADOS DO PDV PET
# ==========================================
def baixar_dados_pdvpet():
    if not USUARIO_PDV or not SENHA_PDV:
        st.error("❌ ERRO CRÍTICO: Variáveis USUARIO_PDV e SENHA_PDV não encontradas nos Secrets!")
        return False

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=True,
            args=[
                "--no-sandbox", 
                "--disable-setuid-sandbox", 
                "--disable-dev-shm-usage",
                "--disable-gpu",
                "--no-zygote",
                "--single-process"
            ]
        )
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            viewport={"width": 1280, "height": 800}
        )
        page = context.new_page()

        try:
            st.info("🔗 Acessando o site PDV Pet...")
            page.goto("https://www.pdvpet.com.br/", timeout=60000, wait_until="networkidle")

            st.info("🔑 Preenchendo dados de login...")
            page.fill('input[type="text"], input[name*="user"], input[name*="cpf"], input[name*="login"]', USUARIO_PDV)
            page.fill('input[type="password"]', SENHA_PDV)

            try:
                page.click('button[type="submit"], input[type="submit"], button:has-text("Entrar")', timeout=5000)
            except Exception:
                page.keyboard.press("Enter")

            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(3000)

            st.info("📋 Navegando até a aba de Questionários...")
            page.wait_for_selector('text=/Questionários|QUESTIONÁRIOS/i', timeout=60000)
            page.click('text=/Questionários|QUESTIONÁRIOS/i')
            
            page.wait_for_selector('#DataDe', timeout=60000)

            fuso_br = ZoneInfo("America/Sao_Paulo")
            data_hoje = datetime.now(fuso_br).strftime("%Y-%m-%d")

            page.fill('#DataDe', INICIO_P9)
            page.fill('#DataAte', data_hoje)
            page.click('button[type="submit"]:has-text("Buscar")')
            page.wait_for_timeout(8000)

            st.info("⏳ Baixando o relatório CSV...")
            with page.expect_download(timeout=60000) as download_info:
                page.click('button.btn-outline-success:has-text("Exportar")')
                try:
                    page.wait_for_selector('a:has-text("Abrir")', timeout=5000)
                    page.click('a:has-text("Abrir")')
                except Exception:
                    pass

            download_info.value.save_as(CAMINHO_CSV_FINAL)
            st.success("✅ CSV Baixado com Sucesso!")
            return True

        except Exception as e:
            st.error(f"❌ OCORREU UM ERRO DURANTE A NAVEGAÇÃO: {e}")
            return False
        finally:
            browser.close()

# ==========================================
# 2. GERADOR DE PDF
# ==========================================
def gerar_pdf_filial(filial, meta_geral_batida, df_oportunidades_filial, df_precos_filial, usar_oportunidades, usar_precos):
    nome_arquivo = f"Relatorio_Oportunidades_{filial.replace(' ', '_').replace('-', '')}.pdf"
    caminho_pdf = os.path.join(PASTA_PROJETO, nome_arquivo)
    
    doc = SimpleDocTemplate(caminho_pdf, pagesize=letter, rightMargin=25, leftMargin=25, topMargin=25, bottomMargin=25)
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontSize=15, textColor=colors.HexColor('#1f6feb'), spaceAfter=4)
    sub_style = ParagraphStyle('SubStyle', parent=styles['Normal'], fontSize=9.5, textColor=colors.HexColor('#57606a'), spaceAfter=12)
    sec_style = ParagraphStyle('SecStyle', parent=styles['Heading2'], fontSize=11, textColor=colors.HexColor('#24292f'), spaceBefore=10, spaceAfter=6)
    alert_style = ParagraphStyle('AlertStyle', parent=styles['Normal'], fontSize=11, textColor=colors.HexColor('#2ea043'), fontName='Helvetica-Bold')
    
    style_corrigido = ParagraphStyle('StyleCorrigido', parent=styles['Normal'], fontSize=7.5, leading=9.5, textColor=colors.HexColor('#047857'))
    style_pendente = ParagraphStyle('StylePendente', parent=styles['Normal'], fontSize=7.5, leading=9.5, textColor=colors.HexColor('#dc2626'), fontName='Helvetica-Bold')
    style_prod = ParagraphStyle('StyleProd', parent=styles['Normal'], fontSize=8, leading=10, textColor=colors.HexColor('#1f2937'))
    style_pdv = ParagraphStyle('StylePdv', parent=styles['Normal'], fontSize=8, leading=10, textColor=colors.HexColor('#1f2937'))

    elements = []
    elements.append(Paragraph(f"<b>Relatório de Oportunidades e Auditoria</b>", title_style))
    elements.append(Paragraph(f"<b>Filial:</b> {filial} | <b>Período:</b> {PERIODO_ATUAL} | <b>Gerado em:</b> {datetime.now().strftime('%d/%m/%Y %H:%M')}", sub_style))

    # SEÇÃO 1: OPORTUNIDADES (Se marcada no filtro)
    if usar_oportunidades:
        elements.append(Paragraph("1. Oportunidades de Leitura (Somente Categoria PENDENTES de Meta)", sec_style))
        if meta_geral_batida:
            elements.append(Paragraph("🏆 TODAS AS METAS BATIDAS! Nenhuma oportunidade pendente.", alert_style))
        else:
            if df_oportunidades_filial.empty:
                elements.append(Paragraph("✅ Nenhuma oportunidade pendente para as metas em aberto.", styles['Normal']))
            else:
                dados_tabela = [["PDV / Cidade", "Item / Opção Embalagem"]]
                for _, row in df_oportunidades_filial.iterrows():
                    style_p = ParagraphStyle('PdvS', parent=styles['Normal'], fontSize=8.5, textColor=colors.HexColor('#24292f'))
                    style_i = ParagraphStyle('ItemS', parent=styles['Normal'], fontSize=8.5, textColor=colors.HexColor('#0056b3'), fontName='Helvetica-Bold')
                    dados_tabela.append([Paragraph(str(row['Pdv_Com_Cidade']), style_p), Paragraph(str(row['Item_Nome']), style_i)])

                t = Table(dados_tabela, colWidths=[250, 310])
                t.setStyle(TableStyle([
                    ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#f6f8fa')),
                    ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#d0d7de')),
                ]))
                elements.append(t)
        elements.append(Spacer(1, 12))

    # SEÇÃO 2: PREÇOS (Se marcada no filtro)
    if usar_precos:
        elements.append(Paragraph("2. Auditoria de Preço Máximo (Small Bags)", sec_style))
        if df_precos_filial.empty:
            elements.append(Paragraph("✅ Nenhum preço acima do teto foi detectado.", styles['Normal']))
        else:
            dados_preco = [["PDV / Cidade", "Produto / Embalagem", "Lido (R$)", "Teto (R$)", "Situação"]]
            for _, row in df_precos_filial.iterrows():
                st_val = str(row['Status'])
                if st_val == 'Cancelado':
                    p_status = Paragraph("✅ Corrigido", style_corrigido)
                else:
                    p_status = Paragraph("⚠️ PENDENTE", style_pendente)
                
                dados_preco.append([
                    Paragraph(str(row['Pdv_Com_Cidade']), style_pdv),
                    Paragraph(str(row['Item_Nome']), style_prod),
                    f"R$ {row['Preco_Lido']:.2f}",
                    f"R$ {row['Preco_Maximo']:.2f}",
                    p_status
                ])

            t_preco = Table(dados_preco, colWidths=[150, 160, 55, 55, 140])
            t_preco.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#ffebe9')),
                ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#ffc1c0')),
            ]))
            elements.append(t_preco)

    doc.build(elements)
    return caminho_pdf

# ==========================================
# 3. DISPARO DE E-MAILS
# ==========================================
def enviar_email(filial, caminho_pdf, tem_preco_pendente, apenas_meu_email):
    if not EMAIL_REMETENTE or not SENHA_EMAIL:
        return

    senha_limpa = SENHA_EMAIL.replace(" ", "")
    
    if apenas_meu_email:
        destinatarios = EMAILS_MEUS
    else:
        destinatarios = list(set(EMAILS_MEUS + EMAILS_PROMOTORES.get(filial, [])))

    msg = MIMEMultipart()
    msg['From'] = EMAIL_REMETENTE
    msg['To'] = ", ".join(destinatarios)
    
    if tem_preco_pendente:
        msg['Subject'] = f"⚠️ [TESTE/FILTRO] Relatório - {filial}"
    else:
        msg['Subject'] = f"Relatório de Oportunidades - {filial}"

    corpo = f"Segue em anexo o relatório referente à filial {filial}.\n\nAtenciosamente,\nAutomação PDV Pet"
    msg.attach(MIMEText(corpo, 'plain'))

    if os.path.exists(caminho_pdf):
        with open(caminho_pdf, "rb") as f:
            part = MIMEApplication(f.read(), Name=os.path.basename(caminho_pdf))
            part['Content-Disposition'] = f'attachment; filename="{os.path.basename(caminho_pdf)}"'
            msg.attach(part)

    try:
        server = smtplib.SMTP_SSL('smtp.gmail.com', 465, timeout=30)
        server.login(EMAIL_REMETENTE, senha_limpa)
        server.sendmail(EMAIL_REMETENTE, destinatarios, msg.as_string())
        server.quit()
        st.success(f"📧 E-mail enviado com sucesso para: {destinatarios}")
    except Exception as e:
        st.error(f"❌ Erro ao enviar e-mail: {e}")

# ==========================================
# BOTÃO DE EXECUÇÃO NA TELA DO STREAMLIT
# ==========================================
if st.button("Executar Automação e Enviar Relatórios"):
    with st.spinner("Processando dados e gerando relatórios..."):
        if not os.path.exists(CAMINHO_CSV_FINAL):
            sucesso = baixar_dados_pdvpet()
            if not sucesso:
                st.stop()

        df = pd.read_csv(CAMINHO_CSV_FINAL, sep=';', encoding='latin1')
        df['Data_Parsed'] = pd.to_datetime(df['Data'].astype(str).str.split(' ').str[0], format='%d/%m/%Y', errors='coerce')
        df['Preco_Num'] = pd.to_numeric(df['PrecoKg'].astype(str).str.replace('R$', '', regex=False).str.strip().str.replace(',', '.'), errors='coerce')
        
        df['Cidade_Clean'] = df['Cidade'].fillna('')
        df['Uf_Clean'] = df['Uf'].fillna('')
        df['Pdv_Com_Cidade'] = df.apply(
            lambda r: f"{r['Pdv']} - {r['Cidade_Clean']}/{r['Uf_Clean']}" if r['Cidade_Clean'] != '' else str(r['Pdv']),
            axis=1
        )

        df['OpcaoEmbalagem_Clean'] = df['OpcaoEmbalagem'].fillna('')
        df['Embalagem_Clean'] = df['Embalagem'].fillna('')
        df['Item_Nome'] = df.apply(
            lambda r: f"{r['Item']} - {r['OpcaoEmbalagem_Clean']}" if r['OpcaoEmbalagem_Clean'] != '' else str(r['Item']), 
            axis=1
        )
        df['Chave_Preco'] = df.apply(
            lambda r: f"{r['OpcaoEmbalagem_Clean']} - {r['Embalagem_Clean']}".strip(" -"),
            axis=1
        )

        df_p9  = df[(df['Data_Parsed'] >= INICIO_P9) & (df['Data_Parsed'] <= FIM_P9)]
        df_p10 = df[(df['Data_Parsed'] >= INICIO_P10) & (df['Data_Parsed'] <= FIM_P10)]

        p9_pares  = df_p9[['Distribuidor', 'Cidade_Clean', 'Pdv_Com_Cidade', 'Item', 'Item_Nome']].drop_duplicates()
        p10_pares = df_p10[['Distribuidor', 'Cidade_Clean', 'Pdv_Com_Cidade', 'Item', 'Item_Nome']].drop_duplicates()

        df_oportunidades = pd.merge(p9_pares, p10_pares, on=['Distribuidor', 'Cidade_Clean', 'Pdv_Com_Cidade', 'Item', 'Item_Nome'], how='left', indicator=True)
        df_oportunidades = df_oportunidades[df_oportunidades['_merge'] == 'left_only'].drop(columns=['_merge'])

        alertas_preco = []
        for _, row in df_p10.iterrows():
            chave = row['Chave_Preco']
            preco = row['Preco_Num']
            if chave in PRECOS_MAXIMOS and pd.notnull(preco) and preco > 0:
                teto = PRECOS_MAXIMOS[chave]
                if preco > teto:
                    alertas_preco.append({
                        'Distribuidor': row['Distribuidor'],
                        'Cidade_Clean': row['Cidade_Clean'],
                        'Pdv_Com_Cidade': row['Pdv_Com_Cidade'],
                        'Item_Nome': chave,
                        'Preco_Lido': preco,
                        'Preco_Maximo': teto,
                        'Status': row['Status']
                    })
        df_alertas_preco = pd.DataFrame(alertas_preco)

        # Definir quais filiais rodar com base no filtro da barra lateral
        filiais_para_rodar = METAS.keys() if filial_selecionada == "Todas" else [filial_selecionada]

        for distribuidor in filiais_para_rodar:
            metas_filial = METAS[distribuidor]
            sub_a = df_p10[(df_p10['Distribuidor'] == distribuidor) & (df_p10['Status'] == 'Aprovado')]['Item'].value_counts().to_dict()
            
            categorias_pendentes = []
            meta_geral_batida = True

            for item_meta, valor_meta in metas_filial.items():
                if sub_a.get(item_meta, 0) < valor_meta:
                    meta_geral_batida = False
                    categorias_pendentes.append(item_meta)

            df_op_filial = df_oportunidades[(df_oportunidades['Distribuidor'] == distribuidor) & (df_oportunidades['Item'].isin(categorias_pendentes))].copy()
            df_pr_filial = df_alertas_preco[df_alertas_preco['Distribuidor'] == distribuidor].copy() if not df_alertas_preco.empty else pd.DataFrame()

            tem_pendente = not df_pr_filial.empty and any(df_pr_filial['Status'] != 'Cancelado')

            caminho_pdf_gerado = gerar_pdf_filial(distribuidor, meta_geral_batida, df_op_filial, df_pr_filial, incluir_oportunidades, incluir_precos)
            enviar_email(distribuidor, caminho_pdf_gerado, tem_pendente, enviar_apenas_para_mim)

    st.success("🎉 Processo finalizado com sucesso!")
