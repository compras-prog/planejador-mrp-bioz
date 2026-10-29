"""Identidade visual Bioz Green para o app Streamlit."""
from __future__ import annotations

VERDE_ESCURO = "#0B592F"
VERDE_CLARO = "#8CA70A"
AMARELO = "#EEB627"
LARANJA = "#DF6F3B"
CREME = "#F0E7D7"
MARROM = "#936440"
CINZA = "#3D3D3D"
LINHA = "#D9D9D9"

# Ordem oficial das séries de gráfico
SERIES = [VERDE_ESCURO, VERDE_CLARO, AMARELO, LARANJA, MARROM, CREME]

COR_FAIXA = {
    "Crítico (< 5 dias)": LARANJA,
    "Atenção (5 a 10 dias)": AMARELO,
    "Moderado (10 a 30 dias)": VERDE_CLARO,
    "Adequado (30 a 60 dias)": VERDE_ESCURO,
    "Excesso (> 60 dias)": MARROM,
    "Sem meta": LINHA,
}

COR_STATUS_MRP = {
    "ATRASADO": "background-color:#F6D5C6;color:#7A2E0E;font-weight:600",
    "PEDIR ESTA SEMANA": "background-color:#FBE9B5;color:#6B4B00;font-weight:600",
    "PROGRAMAR": "background-color:#E2EBB8;color:#3D4A00",
    "COBERTO": "background-color:#DDEBE2;color:#0B592F",
}

CSS = f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Nunito:wght@400;600&family=Poppins:wght@500;600;700&display=swap');
html, body, [class*="css"] {{ font-family: 'Nunito', Calibri, Arial, sans-serif; }}
h1, h2, h3, h4 {{ font-family: 'Poppins', Calibri, Arial, sans-serif !important; color: {VERDE_ESCURO}; }}
.bz-topo {{ background:{VERDE_ESCURO}; height:10px; border-radius:4px 4px 0 0; margin-top:-2rem; }}
.bz-fio {{ background:{VERDE_CLARO}; height:4px; margin-bottom:14px; border-radius:0 0 4px 4px; }}
.bz-sub {{ color:{VERDE_CLARO}; font-family:'Poppins'; font-weight:600; margin:-10px 0 6px 0; }}
div[data-testid="stMetric"] {{ background:{CREME}; border-radius:10px; padding:10px 14px; border:1px solid {LINHA}; }}
div[data-testid="stMetricValue"] {{ color:{VERDE_ESCURO}; font-family:'Poppins'; font-size:1.75rem; }}
button[data-baseweb="tab"] {{ font-family:'Poppins'; font-weight:600; }}
.bz-alerta {{ border-left:6px solid {LARANJA}; background:#FBEDE5; padding:10px 14px; border-radius:6px; margin:6px 0; }}
.bz-ok {{ border-left:6px solid {VERDE_CLARO}; background:#F1F5DC; padding:10px 14px; border-radius:6px; margin:6px 0; }}
.bz-rodape {{ color:#595959; font-size:0.85rem; margin-top:30px; border-top:1px solid {LINHA}; padding-top:8px; }}
</style>
"""

STREAMLIT_CONFIG = f"""[theme]
primaryColor = "{VERDE_ESCURO}"
backgroundColor = "#FFFFFF"
secondaryBackgroundColor = "{CREME}"
textColor = "{CINZA}"
font = "sans serif"

[server]
maxUploadSize = 100
enableXsrfProtection = true

[browser]
gatherUsageStats = false
"""


def fmt_int(v) -> str:
    try:
        return f"{float(v):,.0f}".replace(",", ".")
    except (TypeError, ValueError):
        return "—"


def fmt_num(v, casas=2) -> str:
    try:
        s = f"{float(v):,.{casas}f}"
    except (TypeError, ValueError):
        return "—"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def fmt_brl(v) -> str:
    s = fmt_num(v, 2)
    return "R$ " + s if s != "—" else s
