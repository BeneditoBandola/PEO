import os
import time
import subprocess
import smtplib
import pandas as pd
import streamlit as st
from datetime import datetime
from zoneinfo import ZoneInfo
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from playwright.sync_api import sync_playwright
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

# --- INSTALAÇÃO SEGURA DO PLAYWRIGHT ---
@st.cache_resource
def instalar_navegador():
    # Instala apenas o Chromium (necessário para a nuvem) uma única vez
    os.system("playwright install chromium")

instalar_navegador()
# ---------------------------------------

# ==========================================
# CONFIGURAÇÕES E VARIÁVEIS DE AMBIENTE
# ==========================================
USUARIO_PDV = os.environ.get("USUARIO_PDV")
SENHA_PDV = os.environ.get("SENHA_PDV")
# ... (continue com o resto do seu código a partir daqui)
