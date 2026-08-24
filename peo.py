import os
import smtplib
import pandas as pd
from datetime import datetime
from zoneinfo import ZoneInfo
from playwright.sync_api import sync_playwright

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

# --- CONFIGURAÇÕES DE DESTINATÁRIOS ---
EMAILS_MEUS = ['beneditobandola@gmail.com', 'benedito.bandola@minassal.com.br']

EMAILS_PROMOTORES = {
    'MINASSAL LTDA - POCOS DE CALDAS': ['pamelaalmeida5@hotmail.com'],
    'MINASSAL LTDA - SAO JOAO DA BOA VISTA': ['crbruno27123@gmail.com'], # Rodrigo removido
    'MINASSAL LTDA - SAO JOSE DO RIO PRETO': ['saruetesjrp79@gmail.com'],
    'MINASSAL LTDA - JUIZ DE FORA': ['fernandaferreira_jf@yahoo.com.br', 'madallareis66@gmail.com']
}

# --- CONFIGURAÇÕES DE PERÍODOS E METAS ---
PERIODO_ATUAL = "P9"
INICIO_P8 = "2026-07-13"
FIM_P8    = "2026-08-09"
INICIO_P9 = "2026-08-10"
FIM_P9    = "2026-09-06"

CONFIGURACOES_PERIODOS = {
    "P9": {
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
USUARIO_PDV = os.environ.get("USUARIO_PDV", "13731800640")
SENHA_PDV = os.environ.get("SENHA_PDV", "Fizzvini1234*")
CAMINHO_CSV_FINAL = os.path.join(PASTA_PROJETO, "relatorio_questionarios.csv")

def baixar_dados_pdvpet():
    print(f"\n--- INICIANDO DOWNLOAD DO HISTÓRICO (P8 e P9) ---")
    with sync_playwright() as p:
        # headless=True para execução em segundo plano / servidor na nuvem
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        
        print("🔗 Acessando o site PDV Pet...")
        page.goto("https://www.pdvpet.com.br/")
        
        print("🔑 Preenchendo dados de login...")
        page.fill('input[type="text"], input[name*="user"]', USUARIO_PDV)
        page.fill('input[type="password"]', SENHA_PDV)
        page.keyboard.press("Enter")
        page.wait_for_load_state("networkidle")
        
        print("📋 Navegando até a aba de Questionários...")
        page.click("text=Questionários")
        page.wait_for_selector('#DataDe')

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
            except:
                pass

        download_info.value.save_as(CAMINHO_CSV_FINAL)
        print("✅ CSV Baixado com Sucesso!")
        browser.close()

def enviar_email(filial, caminho_pdf):
    usuario_smtp = os.environ.get("EMAIL_REMETENTE")
    senha_smtp = os.environ.get("SENHA_EMAIL")

    if not usuario_smtp or not senha_smtp:
        print(f"⚠️ Variáveis EMAIL_REMETENTE/SENHA_EMAIL não encontradas. Pulando envio por e-mail para {filial}.")
        return

    destinatarios = EMAILS_MEUS + EMAILS_PROMOTORES.get(filial, [])
    
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

    with open(caminho_pdf, "rb") as f:
        part = MIMEApplication(f.read(), Name=os.path.basename(caminho_pdf))
        part['Content-Disposition'] = f'attachment; filename="{os.path.basename(caminho_pdf)}"'
        msg.attach(part)

    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(usuario_smtp, senha_smtp)
        server.sendmail(usuario_smtp, destinatarios, msg.as_string())
        server.quit()
        print(f"📧 E-mail enviado com sucesso para: {destinatarios}")
    except Exception as e:
        print(f"❌ Erro ao enviar e-mail da filial {filial}: {e}")

def formatar_texto_por_tipo(item_nome, pdv_info, styles):
    item_str = str(item_nome)
    pdv_str = str(pdv_info)
    
    if "Small Bags" in item_str:
        cor = "#0056b3"
    elif "Ponto Extra" in item_str:
        cor = "#6f42c1"
    elif "Combo" in item_str:
        cor = "#d97706"
    elif "Sheba" in item_str:
        cor = "#059669"
    else:
        cor = "#24292f"

    style_pdv = ParagraphStyle('PdvStyle', parent=styles['Normal'], fontSize=9, textColor=colors.HexColor('#24292f'))
    style_item = ParagraphStyle('ItemStyle', parent=styles['Normal'], fontSize=9, textColor=colors.HexColor(cor), fontName='Helvetica-Bold')

    return Paragraph(pdv_str, style_pdv), Paragraph(item_str, style_item)

def gerar_pdf_filial(filial, meta_geral_batida, df_oportunidades_filial, df_precos_filial):
    nome_arquivo = f"Relatorio_Oportunidades_{filial.replace(' ', '_').replace('-', '')}.pdf"
    caminho_pdf = os.path.join(PASTA_PROJETO, nome_arquivo)
    
    doc = SimpleDocTemplate(caminho_pdf, pagesize=letter, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontSize=16, textColor=colors.HexColor('#1f6feb'), spaceAfter=6)
    sub_style = ParagraphStyle('SubStyle', parent=styles['Normal'], fontSize=10, textColor=colors.HexColor('#57606a'), spaceAfter=15)
    sec_style = ParagraphStyle('SecStyle', parent=styles['Heading2'], fontSize=12, textColor=colors.HexColor('#24292f'), spaceBefore=12, spaceAfter=8)
    alert_style = ParagraphStyle('AlertStyle', parent=styles['Normal'], fontSize=12, textColor=colors.HexColor('#2ea043'), fontName='Helvetica-Bold')
    red_alert_style = ParagraphStyle('RedAlertStyle', parent=styles['Normal'], fontSize=9, textColor=colors.HexColor('#dc2626'))

    elements = []

    # Cabeçalho
    elements.append(Paragraph(f"<b>Relatório de Oportunidades e Auditoria</b>", title_style))
    elements.append(Paragraph(f"<b>Filial:</b> {filial} | <b>Período:</b> P9 | <b>Gerado em:</b> {datetime.now().strftime('%d/%m/%Y %H:%M')}", sub_style))

    # SEÇÃO 1: OPORTUNIDADES
    elements.append(Paragraph("1. Oportunidades de Leitura (Somente Categoria PENDENTES de Meta)", sec_style))
    
    if meta_geral_batida:
        elements.append(Paragraph("🏆 TODAS AS METAS BATIDAS! Nenhuma oportunidade pendente.", alert_style))
    else:
        if df_oportunidades_filial.empty:
            elements.append(Paragraph("✅ Nenhuma oportunidade pendente para as metas em aberto.", styles['Normal']))
        else:
            dados_tabela = [["PDV / Cidade", "Item / Opção Embalagem"]]
            for _, row in df_oportunidades_filial.iterrows():
                p_pdv, p_item = formatar_texto_por_tipo(row['Item_Nome'], row['Pdv_Com_Cidade'], styles)
                dados_tabela.append([p_pdv, p_item])

            t = Table(dados_tabela, colWidths=[240, 290])
            t.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#f6f8fa')),
                ('TEXTCOLOR', (0,0), (-1,0), colors.HexColor('#24292f')),
                ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                ('BOTTOMPADDING', (0,0), (-1,0), 6),
                ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#d0d7de')),
            ]))
            elements.append(t)

    elements.append(Spacer(1, 15))

    # SEÇÃO 2: DIVERGÊNCIAS DE PREÇO
    elements.append(Paragraph("2. Divergências de Preço (Preços Acima do Máximo)", sec_style))

    if df_precos_filial.empty:
        elements.append(Paragraph("✅ Nenhum preço acima do teto foi detectado.", styles['Normal']))
    else:
        dados_preco = [["PDV / Cidade", "Item / Opção Embalagem", "Lido (R$)", "Teto (R$)", "Dif (R$)"]]
        for _, row in df_precos_filial.iterrows():
            dados_preco.append([
                Paragraph(str(row['Pdv_Com_Cidade']), styles['Normal']),
                Paragraph(f"<b>{row['Item_Nome']}</b>", red_alert_style),
                f"R$ {row['Preco_Lido']:.2f}",
                f"R$ {row['Preco_Maximo']:.2f}",
                f"+ R$ {row['Diferenca']:.2f}"
            ])

        t_preco = Table(dados_preco, colWidths=[160, 170, 65, 65, 70])
        t_preco.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#ffebe9')),
            ('TEXTCOLOR', (0,0), (-1,0), colors.HexColor('#cf222e')),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('BOTTOMPADDING', (0,0), (-1,0), 6),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#ffc1c0')),
        ]))
        elements.append(t_preco)

    doc.build(elements)
    print(f"📄 PDF Gerado: {nome_arquivo}")
    
    # Envio do e-mail
    enviar_email(filial, caminho_pdf)

def processar_e_gerar_relatorios():
    if not os.path.exists(CAMINHO_CSV_FINAL):
        baixar_dados_pdvpet()

    df = pd.read_csv(CAMINHO_CSV_FINAL, sep=';', encoding='latin1')
    
    df['Data_Parsed'] = pd.to_datetime(df['Data'].astype(str).str.split(' ').str[0], format='%d/%m/%Y', errors='coerce')
    df['Preco_Num'] = pd.to_numeric(df['PrecoKg'].astype(str).str.replace(',', '.'), errors='coerce')
    
    df['Cidade_Clean'] = df['Cidade'].fillna('')
    df['Uf_Clean'] = df['Uf'].fillna('')
    df['Pdv_Com_Cidade'] = df.apply(
        lambda r: f"{r['Pdv']} - {r['Cidade_Clean']}/{r['Uf_Clean']}" if r['Cidade_Clean'] != '' else str(r['Pdv']),
        axis=1
    )

    df['OpcaoEmbalagem_Clean'] = df['OpcaoEmbalagem'].fillna('')
    df['Item_Nome'] = df.apply(
        lambda r: f"{r['Item']} - {r['OpcaoEmbalagem_Clean']}" if r['OpcaoEmbalagem_Clean'] != '' else str(r['Item']), 
        axis=1
    )

    df_p8 = df[(df['Data_Parsed'] >= INICIO_P8) & (df['Data_Parsed'] <= FIM_P8)]
    df_p9 = df[(df['Data_Parsed'] >= INICIO_P9) & (df['Data_Parsed'] <= FIM_P9)]

    p8_pares = df_p8[['Distribuidor', 'Cidade_Clean', 'Pdv_Com_Cidade', 'Item', 'Item_Nome']].drop_duplicates()
    p9_pares = df_p9[['Distribuidor', 'Cidade_Clean', 'Pdv_Com_Cidade', 'Item', 'Item_Nome']].drop_duplicates()

    df_oportunidades = pd.merge(p8_pares, p9_pares, on=['Distribuidor', 'Cidade_Clean', 'Pdv_Com_Cidade', 'Item', 'Item_Nome'], how='left', indicator=True)
    df_oportunidades = df_oportunidades[df_oportunidades['_merge'] == 'left_only'].drop(columns=['_merge'])

    alertas_preco = []
    for _, row in df_p9.iterrows():
        item_nome = row['Item_Nome']
        opcao_emb = row['OpcaoEmbalagem_Clean']
        preco = row['Preco_Num']
        
        chave_preco = opcao_emb if opcao_emb in PRECOS_MAXIMOS else item_nome

        if chave_preco in PRECOS_MAXIMOS and pd.notnull(preco):
            teto = PRECOS_MAXIMOS[chave_preco]
            if preco > teto:
                alertas_preco.append({
                    'Distribuidor': row['Distribuidor'],
                    'Cidade_Clean': row['Cidade_Clean'],
                    'Pdv_Com_Cidade': row['Pdv_Com_Cidade'],
                    'Item_Nome': item_nome,
                    'Preco_Lido': preco,
                    'Preco_Maximo': teto,
                    'Diferenca': round(preco - teto, 2)
                })
    df_alertas_preco = pd.DataFrame(alertas_preco)

    for distribuidor, metas_filial in METAS.items():
        sub_a = df_p9[(df_p9['Distribuidor'] == distribuidor) & (df_p9['Status'] == 'Aprovado')]['Item'].value_counts().to_dict()
        
        categorias_pendentes = []
        meta_geral_batida = True

        for item_meta, valor_meta in metas_filial.items():
            realizado = sub_a.get(item_meta, 0)
            if realizado < valor_meta:
                meta_geral_batida = False
                categorias_pendentes.append(item_meta)

        df_op_filial = df_oportunidades[
            (df_oportunidades['Distribuidor'] == distribuidor) & 
            (df_oportunidades['Item'].isin(categorias_pendentes))
        ].copy()

        df_pr_filial = df_alertas_preco[df_alertas_preco['Distribuidor'] == distribuidor].copy() if not df_alertas_preco.empty else pd.DataFrame()

        # Exclusão de Ribeirão Preto para São João da Boa Vista
        if distribuidor == 'MINASSAL LTDA - SAO JOAO DA BOA VISTA':
            if not df_op_filial.empty:
                df_op_filial = df_op_filial[~df_op_filial['Cidade_Clean'].str.upper().str.contains("RIBEIRAO PRETO|RIBEIRÃO PRETO", na=False)]
            if not df_pr_filial.empty:
                df_pr_filial = df_pr_filial[~df_pr_filial['Cidade_Clean'].str.upper().str.contains("RIBEIRAO PRETO|RIBEIRÃO PRETO", na=False)]

        gerar_pdf_filial(distribuidor, meta_geral_batida, df_op_filial, df_pr_filial)

if __name__ == "__main__":
    try:
        processar_e_gerar_relatorios()
        print("\n--------------------------------------------------")
        print("✅ PROCESSO CONCLUÍDO COM SUCESSO!")
    except Exception as e:
        print("\n--------------------------------------------------")
        print(f"❌ OCORREU UM ERRO DURANTE A EXECUÇÃO:")
        print(e)
        print("--------------------------------------------------")
