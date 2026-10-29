"""
Conector opcional com o Google Drive (arquivos .xlsx ou Google Sheets) via conta de serviço.

Configuração (uma vez):
  1. No Google Cloud, ative a "Google Drive API" e crie uma Conta de Serviço (JSON).
  2. Compartilhe a pasta/arquivo no Drive com o e-mail da conta de serviço (permissão de leitor).
  3. Cole o JSON em st.secrets["gcp_service_account"] (Streamlit Cloud: Settings > Secrets)
     ou monte como variável GCP_SERVICE_ACCOUNT_JSON (Cloud Run).
"""
from __future__ import annotations

import io
import json
import os
import re
from typing import Dict, List, Optional, Tuple

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
GSHEET_MIME = "application/vnd.google-apps.spreadsheet"
ESCOPO = ["https://www.googleapis.com/auth/drive.readonly"]


def extrair_id(texto: str) -> str:
    """Aceita o ID puro ou a URL completa do Drive/Sheets."""
    texto = (texto or "").strip()
    m = re.search(r"/d/([a-zA-Z0-9_-]{20,})", texto) or re.search(r"[?&]id=([a-zA-Z0-9_-]{20,})", texto)
    return m.group(1) if m else texto


def _credenciais(info: Optional[dict] = None):
    from google.oauth2 import service_account  # import tardio: dependência opcional

    if info is None:
        bruto = os.environ.get("GCP_SERVICE_ACCOUNT_JSON")
        if not bruto:
            raise RuntimeError("Credenciais do Google não configuradas (st.secrets['gcp_service_account'] "
                               "ou variável GCP_SERVICE_ACCOUNT_JSON).")
        info = json.loads(bruto)
    return service_account.Credentials.from_service_account_info(dict(info), scopes=ESCOPO)


def baixar_planilha(file_id: str, info: Optional[dict] = None) -> Tuple[str, bytes]:
    """Baixa um .xlsx (ou exporta um Google Sheets como .xlsx). Retorna (nome, bytes)."""
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaIoBaseDownload

    svc = build("drive", "v3", credentials=_credenciais(info), cache_discovery=False)
    meta = svc.files().get(fileId=file_id, fields="id,name,mimeType", supportsAllDrives=True).execute()
    nome = meta["name"] if meta["name"].lower().endswith(".xlsx") else meta["name"] + ".xlsx"
    if meta["mimeType"] == GSHEET_MIME:
        req = svc.files().export_media(fileId=file_id, mimeType=XLSX_MIME)
    else:
        req = svc.files().get_media(fileId=file_id, supportsAllDrives=True)
    buf = io.BytesIO()
    dl = MediaIoBaseDownload(buf, req)
    feito = False
    while not feito:
        _, feito = dl.next_chunk()
    return nome, buf.getvalue()


def listar_planilhas(pasta_id: str, info: Optional[dict] = None) -> List[Dict[str, str]]:
    """Lista .xlsx e Google Sheets de uma pasta do Drive (para escolher na tela)."""
    from googleapiclient.discovery import build

    svc = build("drive", "v3", credentials=_credenciais(info), cache_discovery=False)
    q = (f"'{pasta_id}' in parents and trashed = false and "
         f"(mimeType = '{XLSX_MIME}' or mimeType = '{GSHEET_MIME}')")
    r = svc.files().list(q=q, fields="files(id,name,modifiedTime)", orderBy="modifiedTime desc",
                         pageSize=50, supportsAllDrives=True, includeItemsFromAllDrives=True).execute()
    return r.get("files", [])
