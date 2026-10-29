"""
Leitura e normalização das planilhas de entrada.

As abas são encontradas pela ASSINATURA DE COLUNAS (não pelo nome), então funciona com:
  • "Dados ponto Reposição.xlsx"   (Consulta1 / Reposicao + Dados SKU)
  • "Demanda_Plenno_12_2026.xlsx"  (Estruturas_Completas, Estoque Plenno_*, Pedido de Compra *, Dados_Fornecedores, Meta de vendas)
  • o modelo padronizado gerado por modelo.py
Pode-se enviar um arquivo único com tudo ou vários arquivos.
"""
from __future__ import annotations

import io
import re
import warnings
from dataclasses import dataclass, field
from datetime import date, datetime
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd

from engine import classificar_pedidos, norm

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

MESES = ["JANEIRO", "FEVEREIRO", "MARCO", "ABRIL", "MAIO", "JUNHO", "JULHO", "AGOSTO",
         "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"]

CAPACIDADE_PADRAO = {"Plenno": 50, "LiveMax": 50, "Bioz": 60}   # da aba "Simulador" — CONFIRMAR


# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #
def achar_col(df: pd.DataFrame, *aliases: str) -> Optional[str]:
    """Acha coluna por nome normalizado: igualdade primeiro, depois 'começa com'."""
    cols = {norm(c): c for c in df.columns}
    for a in aliases:
        if norm(a) in cols:
            return cols[norm(a)]
    for a in aliases:
        na = norm(a)
        for k, c in cols.items():
            if k.startswith(na):
                return c
    return None


def to_num(s: pd.Series) -> pd.Series:
    """Converte para número aceitando texto pt-BR ('1.234,5', 'R$ 10,00', '8,00%')."""
    def f(v):
        if v is None or (isinstance(v, float) and np.isnan(v)):
            return np.nan
        if isinstance(v, (int, float, np.integer, np.floating)):
            return float(v)
        t = str(v).strip().replace("R$", "").replace("%", "").strip()
        if t in ("", "-"):
            return np.nan
        if "," in t:
            t = t.replace(".", "").replace(",", ".")
        try:
            return float(t)
        except ValueError:
            return np.nan
    return s.map(f).astype(float)


def to_date(s: pd.Series) -> pd.Series:
    def f(v):
        if v is None or (isinstance(v, float) and np.isnan(v)):
            return pd.NaT
        if isinstance(v, (pd.Timestamp, datetime, date)):
            return pd.Timestamp(v)
        if isinstance(v, (int, float, np.integer, np.floating)):
            return pd.Timestamp("1899-12-30") + pd.to_timedelta(float(v), unit="D")
        t = str(v).strip()
        if not t or t == "-":
            return pd.NaT
        return pd.to_datetime(t, dayfirst=True, errors="coerce")
    return pd.to_datetime(s.map(f), errors="coerce")


def _txt(s: pd.Series) -> pd.Series:
    return s.astype("object").map(lambda v: "" if v is None or (isinstance(v, float) and np.isnan(v)) else str(v).strip())


# --------------------------------------------------------------------------- #
# Conjunto de planilhas
# --------------------------------------------------------------------------- #
class Planilhas:
    """Vários arquivos .xlsx tratados como um único conjunto de abas."""

    def __init__(self, arquivos: Dict[str, bytes]):
        self.xls: Dict[str, pd.ExcelFile] = {
            nome: pd.ExcelFile(io.BytesIO(b), engine="openpyxl") for nome, b in arquivos.items()
        }
        self._cab: Dict[Tuple[str, str], List[str]] = {}
        self.ocultas: List[Tuple[str, str]] = []
        for nome, x in self.xls.items():
            try:
                for ws in x.book.worksheets:
                    if getattr(ws, "sheet_state", "visible") != "visible":
                        self.ocultas.append((nome, ws.title))
            except Exception:
                pass

    def abas(self) -> List[Tuple[str, str]]:
        """Somente abas visíveis: abas ocultas costumam ser rascunho/legado e já causaram leitura errada."""
        return [(a, s) for a, x in self.xls.items() for s in x.sheet_names if (a, s) not in self.ocultas]

    def cabecalho(self, arq: str, aba: str) -> List[str]:
        k = (arq, aba)
        if k not in self._cab:
            try:
                df = self.xls[arq].parse(aba, nrows=0)
                self._cab[k] = [norm(c) for c in df.columns]
            except Exception:
                self._cab[k] = []
        return self._cab[k]

    def ler(self, ref: Tuple[str, str]) -> pd.DataFrame:
        df = self.xls[ref[0]].parse(ref[1])
        # linhas de exemplo do modelo ("EXEMPLO — substituir") nunca entram nos cálculos
        marca = df.apply(lambda col: col.map(lambda v: isinstance(v, str) and v.strip().upper().startswith("EXEMPLO")))
        return df[~marca.any(axis=1)].reset_index(drop=True) if len(df) else df

    def aba_por_nome(self, *nomes: str) -> Optional[Tuple[str, str]]:
        alvo = {norm(n).replace("_", " ") for n in nomes}
        for a, sh in self.abas():
            if norm(sh).replace("_", " ") in alvo:
                return (a, sh)
        return None


def _tem(cab: List[str], *aliases: str) -> bool:
    """True se alguma das colunas normalizadas casa (igual ou começa com) com algum alias."""
    for a in aliases:
        na = norm(a)
        if any(c == na or c.startswith(na) for c in cab):
            return True
    return False


# Assinaturas por papel -----------------------------------------------------
ASSINATURAS = {
    "skus": lambda c: _tem(c, "sku", "codigo (sku)") and _tem(c, "estoque unit", "estoque_un", "estoque unitario")
                      and _tem(c, "meta mes", "meta_mensal"),
    "dados_sku": lambda c: _tem(c, "codigo (sku)", "sku") and _tem(c, "cx pallet", "cx_pallet", "caixas por pallet")
                           and _tem(c, "cx", "un_por_cx") and not _tem(c, "estoque"),
    "bom": lambda c: (_tem(c, "data.sku kit/fabricado") and _tem(c, "data.sku componente"))
                     or (_tem(c, "kit_sku") and _tem(c, "comp_sku") and _tem(c, "qtd")),
    "estoque_insumos": lambda c: (_tem(c, "sku componente") and _tem(c, "estoque") and not _tem(c, "meta"))
                                 or (_tem(c, "comp_sku") and _tem(c, "estoque") and not _tem(c, "kit_sku")),
    "pedidos": lambda c: (_tem(c, "ordem de compra") and _tem(c, "fornecedor") and _tem(c, "item"))
                         or (_tem(c, "oc") and _tem(c, "fornecedor") and _tem(c, "item") and _tem(c, "qtd")),
    "cadastro": lambda c: (_tem(c, "preco de custo") and _tem(c, "fornecedor") and _tem(c, "codigo (sku)"))
                          or (_tem(c, "comp_sku") and _tem(c, "preco") and _tem(c, "fornecedor")),
    "metas": lambda c: _tem(c, "produto") and sum(_tem(c, m) for m in MESES) >= 3,
    "lead_time": lambda c: _tem(c, "fornecedor") and _tem(c, "lead time")
                           and not _tem(c, "codigo", "categoria", "nome") and len(c) <= 8,
    "capacidade": lambda c: _tem(c, "fabrica") and _tem(c, "capacidade"),
}

PREFERENCIA_NOME = {
    "skus": ["skus", "consulta1", "reposicao"],
    "metas": ["metas", "meta de vendas"],
}


def descobrir(pl: Planilhas) -> Dict[str, List[Tuple[str, str]]]:
    """Papel -> lista de (arquivo, aba) candidatas, da mais provável para a menos."""
    achados: Dict[str, List[Tuple[str, str]]] = {k: [] for k in ASSINATURAS}
    for arq, aba in pl.abas():
        cab = pl.cabecalho(arq, aba)
        if not cab:
            continue
        for papel, sig in ASSINATURAS.items():
            try:
                if sig(cab):
                    achados[papel].append((arq, aba))
            except Exception:
                pass

    def chave(papel):
        pref = PREFERENCIA_NOME.get(papel, [])

        def k(ref):
            n = norm(ref[1])
            for i, p in enumerate(pref):
                if n == p:
                    return (0, i, "")
            # abas com data no nome: a mais recente (ordem alfabética invertida) primeiro
            return (1, 0, "".join(chr(255 - ord(ch)) if ord(ch) < 256 else ch for ch in n))
        return k

    for papel in achados:
        achados[papel].sort(key=chave(papel))
    return achados


# --------------------------------------------------------------------------- #
# Normalizadores
# --------------------------------------------------------------------------- #
def normalizar_skus(df: pd.DataFrame, dados_sku: Optional[pd.DataFrame] = None) -> pd.DataFrame:
    c = {
        "sku": achar_col(df, "sku", "codigo (sku)", "codigo sku"),
        "descricao": achar_col(df, "descricao", "produto"),
        "fabrica": achar_col(df, "fabrica"),
        "unidade": achar_col(df, "unidade"),
        "un_por_cx": achar_col(df, "un_por_cx", "unid por caixa"),
        "cx_por_pallet": achar_col(df, "cx_por_pallet", "pallet cx", "cx pallet"),
        "estoque_un": achar_col(df, "estoque_un", "estoque unit", "estoque unitario"),
        "em_producao": achar_col(df, "em_producao", "em producao"),
        "meta_mensal": achar_col(df, "meta_mensal", "meta mes"),
        "reservado_cx": achar_col(df, "reservado_cx", "reservado"),
        "prioridade": achar_col(df, "prioridade"),
    }
    df = df[df[c["sku"]].notna()].copy()
    out = pd.DataFrame({"sku": _txt(df[c["sku"]])})
    out["descricao"] = _txt(df[c["descricao"]]) if c["descricao"] else out["sku"]
    out["fabrica"] = _txt(df[c["fabrica"]]) if c["fabrica"] else "(sem fábrica)"
    out["estoque_un"] = to_num(df[c["estoque_un"]]).fillna(0)
    out["em_producao"] = to_num(df[c["em_producao"]]).fillna(0) if c["em_producao"] else 0.0
    out["meta_mensal"] = to_num(df[c["meta_mensal"]]).fillna(0)
    out["reservado_cx"] = to_num(df[c["reservado_cx"]]).fillna(0) if c["reservado_cx"] else 0.0
    out["prioridade"] = to_num(df[c["prioridade"]]) if c["prioridade"] else np.nan

    # unidades por caixa: coluna explícita > Dados SKU > texto "CX12"
    un_cx = pd.Series(np.nan, index=out.index)
    if c["un_por_cx"]:
        un_cx = to_num(df[c["un_por_cx"]])
    if dados_sku is not None:
        ds_sku = achar_col(dados_sku, "sku", "codigo (sku)")
        ds_cx = achar_col(dados_sku, "un_por_cx", "cx")
        ds_cxp = achar_col(dados_sku, "cx_por_pallet", "cx pallet", "caixas por pallet")
        m_cx = pd.Series(to_num(dados_sku[ds_cx]).values, index=_txt(dados_sku[ds_sku]).values)
        m_cx = m_cx[~m_cx.index.duplicated()]
        un_cx = un_cx.fillna(out["sku"].map(m_cx))
        m_cxp = pd.Series(to_num(dados_sku[ds_cxp]).values, index=_txt(dados_sku[ds_sku]).values)
        m_cxp = m_cxp[~m_cxp.index.duplicated()]
    else:
        m_cxp = None
    if c["unidade"]:
        parsed = _txt(df[c["unidade"]]).str.extract(r"(\d+)")[0].astype(float)
        un_cx = un_cx.fillna(parsed)
    out["un_por_cx"] = un_cx

    cxp = pd.Series(np.nan, index=out.index)
    if c["cx_por_pallet"]:
        cxp = to_num(df[c["cx_por_pallet"]])
    if m_cxp is not None:
        # Dados SKU é a fonte oficial; a coluna "Pallet CX" da consulta tem zeros indevidos
        cxp = out["sku"].map(m_cxp).fillna(cxp.where(cxp > 0))
    out["cx_por_pallet"] = cxp.fillna(0)

    out = out.drop_duplicates("sku", keep="first").reset_index(drop=True)
    return out[["sku", "descricao", "fabrica", "un_por_cx", "cx_por_pallet", "estoque_un",
                "em_producao", "meta_mensal", "reservado_cx", "prioridade"]]


def normalizar_bom(df: pd.DataFrame) -> pd.DataFrame:
    c = {
        "kit_sku": achar_col(df, "data.sku kit/fabricado", "kit_sku"),
        "comp_id": achar_col(df, "data.id componente", "comp_id"),
        "comp_sku": achar_col(df, "data.sku componente", "comp_sku"),
        "comp_desc": achar_col(df, "data.descricao componente", "comp_desc"),
        "qtd": achar_col(df, "data.quantidade componente", "qtd"),
        "un": achar_col(df, "data.unidade", "un"),
    }
    out = pd.DataFrame({
        "kit_sku": _txt(df[c["kit_sku"]]),
        "comp_id": to_num(df[c["comp_id"]]) if c["comp_id"] else np.nan,
        "comp_sku": _txt(df[c["comp_sku"]]),
        "comp_desc": _txt(df[c["comp_desc"]]) if c["comp_desc"] else "",
        "qtd": to_num(df[c["qtd"]]),
        "un": _txt(df[c["un"]]) if c["un"] else "",
    })
    out = out[(out["kit_sku"] != "") & (out["comp_sku"] != "") & out["qtd"].notna()]
    return out.reset_index(drop=True)


def inferir_local(nome_aba: str) -> str:
    """Deduz onde está o estoque pelo nome da aba/arquivo (ex.: 'Estoque Plenno_02_10' → Plenno)."""
    n = norm(nome_aba).replace(" ", "")
    if "plenno" in n:
        return "Plenno"
    if "livemax" in n or "live" in n:
        return "LiveMax"
    return ""


def normalizar_estoque_insumos(df: pd.DataFrame, local: str = "") -> Tuple[pd.DataFrame, int]:
    """Retorna (comp_sku, estoque, local). A coluna 'local' da aba vence; senão usa o parâmetro."""
    c_sku = achar_col(df, "sku componente", "comp_sku")
    c_est = achar_col(df, "estoque")
    c_cod = achar_col(df, "cod. plenno", "cod_plenno")
    c_loc = achar_col(df, "local", "deposito", "fabrica")
    c_dt = achar_col(df, "data_saldo", "data do saldo")
    out = pd.DataFrame({
        "comp_sku": _txt(df[c_sku]),
        "estoque": to_num(df[c_est]),
        "cod_plenno": _txt(df[c_cod]) if c_cod else "",
        "local": _txt(df[c_loc]) if c_loc else local,
        "data_saldo": to_date(df[c_dt]) if c_dt else pd.NaT,
    })
    out["local"] = out["local"].replace("", local)
    out = out[(out["comp_sku"] != "") & out["estoque"].notna()]
    antes = len(out)
    # mesma linha repetida (mesmo SKU + mesmo código no terceiro + mesmo saldo) = duplicidade da planilha
    out = out.drop_duplicates(["comp_sku", "cod_plenno", "estoque", "local"]).reset_index(drop=True)
    return out[["comp_sku", "estoque", "local", "data_saldo"]], antes - len(out)


def normalizar_pedidos(df: pd.DataFrame) -> pd.DataFrame:
    c = {
        "oc": achar_col(df, "ordem de compra", "oc"),
        "data": achar_col(df, "data"),
        "prevista": achar_col(df, "data prevista", "prevista"),
        "fornecedor": achar_col(df, "fornecedor"),
        "item": achar_col(df, "item"),
        "qtd": achar_col(df, "quantidade", "qtd"),
        "cod_id": achar_col(df, "cod.id", "cod_id"),
    }
    out = pd.DataFrame({
        "oc": _txt(df[c["oc"]]),
        "cod_id": df[c["cod_id"]] if c["cod_id"] else np.nan,
        "data": to_date(df[c["data"]]) if c["data"] else pd.NaT,
        "prevista": to_date(df[c["prevista"]]) if c["prevista"] else pd.NaT,
        "fornecedor": _txt(df[c["fornecedor"]]),
        "item": _txt(df[c["item"]]),
        "qtd": to_num(df[c["qtd"]]).fillna(0),
    })
    return out[out["item"] != ""].reset_index(drop=True)


def normalizar_pedidos_criados(df: pd.DataFrame) -> pd.DataFrame:
    """Aba 'Pedidos_Criados' do modelo: só linhas com insumo e quantidade > 0 (as demais são linhas em branco formatadas)."""
    c_ins = achar_col(df, "comp_sku", "insumo", "sku")
    c_qtd = achar_col(df, "qtd", "quantidade")
    if not c_ins or not c_qtd:
        return pd.DataFrame(columns=["oc", "comp_sku", "qtd", "data_entrega", "entrega_em", "fornecedor"])
    c_oc, c_dt = achar_col(df, "oc", "pedido"), achar_col(df, "data_entrega", "data de entrega")
    c_loc, c_for = achar_col(df, "entrega_em", "entrega em"), achar_col(df, "fornecedor")
    out = pd.DataFrame({
        "oc": _txt(df[c_oc]) if c_oc else "", "comp_sku": _txt(df[c_ins]), "qtd": to_num(df[c_qtd]),
        "data_entrega": to_date(df[c_dt]) if c_dt else pd.NaT,
        "entrega_em": _txt(df[c_loc]) if c_loc else "", "fornecedor": _txt(df[c_for]) if c_for else "",
    })
    return out[(out["comp_sku"] != "") & (out["qtd"].fillna(0) > 0)].reset_index(drop=True)


def normalizar_cadastro(df: pd.DataFrame) -> pd.DataFrame:
    c = {
        "comp_sku": achar_col(df, "codigo (sku)", "comp_sku"),
        "descricao": achar_col(df, "descricao"),
        "unidade": achar_col(df, "unidade"),
        "preco": achar_col(df, "preco de custo", "preco"),
        "fornecedor": achar_col(df, "fornecedor"),
    }
    out = pd.DataFrame({
        "comp_sku": _txt(df[c["comp_sku"]]),
        "descricao": _txt(df[c["descricao"]]) if c["descricao"] else "",
        "unidade": _txt(df[c["unidade"]]) if c["unidade"] else "",
        "preco": to_num(df[c["preco"]]),
        "fornecedor": _txt(df[c["fornecedor"]]).replace("", np.nan),
    })
    return out[out["comp_sku"] != ""].drop_duplicates("comp_sku").reset_index(drop=True)


def normalizar_metas(df: pd.DataFrame) -> pd.DataFrame:
    """Meta de vendas por mês: índice = descrição normalizada, colunas = meses em MAIÚSCULAS."""
    c_prod = achar_col(df, "produto")
    cols = {}
    for col in df.columns:
        n = norm(col).upper().replace("Ç", "C")
        if n in MESES:
            cols[col] = n
    out = df[[c_prod] + list(cols)].rename(columns=cols)
    out["_k"] = out[c_prod].map(norm)
    out = out.drop(columns=[c_prod]).drop_duplicates("_k").set_index("_k")
    return out.apply(to_num)


def normalizar_lead_time(df: pd.DataFrame) -> pd.DataFrame:
    c_f = achar_col(df, "fornecedor")
    c_lt = achar_col(df, "lead time")
    c_lm = achar_col(df, "lote minimo", "lote_minimo")
    c_m = achar_col(df, "multiplo", "multiplo de compra")
    return pd.DataFrame({
        "fornecedor": _txt(df[c_f]),
        "lead_time_dias": to_num(df[c_lt]),
        "lote_minimo": to_num(df[c_lm]) if c_lm else np.nan,
        "multiplo": to_num(df[c_m]) if c_m else np.nan,
    }).query("fornecedor != ''").reset_index(drop=True)


def normalizar_capacidade(df: pd.DataFrame) -> pd.DataFrame:
    c_f = achar_col(df, "fabrica")
    c_c = achar_col(df, "capacidade")
    return pd.DataFrame({
        "fabrica": _txt(df[c_f]),
        "capacidade_posicoes": to_num(df[c_c]),
    }).query("fabrica != ''").reset_index(drop=True)


# --------------------------------------------------------------------------- #
# Pacote de dados
# --------------------------------------------------------------------------- #
@dataclass
class Dados:
    skus: pd.DataFrame
    metas: Optional[pd.DataFrame]
    bom: Optional[pd.DataFrame]
    estoque: Optional[pd.DataFrame]
    pedidos: Optional[pd.DataFrame]
    cadastro: Optional[pd.DataFrame]
    lead_time: pd.DataFrame
    capacidade: pd.DataFrame
    dups_estoque: int = 0
    pedidos_criados: Optional[pd.DataFrame] = None
    abas_usadas: Dict[str, str] = field(default_factory=dict)
    avisos: List[Tuple[str, str]] = field(default_factory=list)      # (nível, mensagem)


def _data_no_nome_da_aba(nome: str) -> "pd.Timestamp":
    """Lê 'dd_mm' do nome da aba (ex.: 'Estoque Plenno_02_10' = 02/10 do ano corrente). NaT se não houver."""
    m = re.search(r"(?<!\d)(\d{2})_(\d{2})(?!\d)", nome or "")
    if m:
        try:
            return pd.Timestamp(year=datetime.now().year, month=int(m.group(2)), day=int(m.group(1)))
        except ValueError:
            pass
    return pd.NaT


def carregar(pl: Planilhas, escolhas: Optional[Dict[str, Tuple[str, str]]] = None) -> Dados:
    """Lê as abas (auto-detectadas ou escolhidas) e devolve tudo normalizado."""
    cand = descobrir(pl)
    escolhas = escolhas or {}
    ref = {p: escolhas.get(p) or (cand[p][0] if cand[p] else None) for p in cand}
    avisos: List[Tuple[str, str]] = []
    usadas = {p: f"{r[0]} › {r[1]}" for p, r in ref.items() if r}

    if ref["skus"] is None:
        raise ValueError(
            "Não encontrei a aba de SKUs (precisa ter colunas de SKU, Estoque unit e Meta mês). "
            "Envie a planilha de reposição ou use o modelo padronizado."
        )
    dados_sku = pl.ler(ref["dados_sku"]) if ref["dados_sku"] else None
    skus = normalizar_skus(pl.ler(ref["skus"]), dados_sku)
    if dados_sku is None:
        avisos.append(("aviso", "Aba 'Dados SKU' (caixas por pallet) não encontrada — usei a coluna de pallet da própria aba de SKUs."))

    metas = normalizar_metas(pl.ler(ref["metas"])) if ref["metas"] else None
    bom = normalizar_bom(pl.ler(ref["bom"])) if ref["bom"] else None
    estoque, dups = None, 0
    if ref["estoque_insumos"]:
        if "estoque_insumos" in escolhas or len(cand["estoque_insumos"]) <= 1:
            fontes = [ref["estoque_insumos"]]
        else:                                    # várias abas: a mais recente de cada local (Plenno, Live Max…)
            vistos, fontes = set(), []
            for r in cand["estoque_insumos"]:
                loc = inferir_local(r[1]) or inferir_local(r[0]) or r[1]
                if loc not in vistos:
                    vistos.add(loc)
                    fontes.append(r)
        partes = []
        for r in fontes:
            e, dp = normalizar_estoque_insumos(pl.ler(r), inferir_local(r[1]) or inferir_local(r[0]))
            e["data_saldo"] = e["data_saldo"].fillna(_data_no_nome_da_aba(r[1]))      # 'Estoque Plenno_02_10' -> 02/10
            partes.append(e)
            dups += dp
        estoque = pd.concat(partes, ignore_index=True)
        usadas["estoque_insumos"] = "; ".join(f"{r[0]} › {r[1]}" for r in fontes)
    pedidos = None
    if ref["pedidos"]:
        pedidos = normalizar_pedidos(pl.ler(ref["pedidos"]))
        if bom is not None:
            pedidos = classificar_pedidos(pedidos, bom)
    cadastro = normalizar_cadastro(pl.ler(ref["cadastro"])) if ref["cadastro"] else None

    # Lead time por fornecedor
    if ref["lead_time"]:
        lt = normalizar_lead_time(pl.ler(ref["lead_time"]))
    else:
        forn = sorted(set(cadastro["fornecedor"].dropna())) if cadastro is not None else []
        lt = pd.DataFrame({"fornecedor": forn, "lead_time_dias": np.nan, "lote_minimo": np.nan, "multiplo": np.nan})
        if forn:
            avisos.append(("aviso", "Nenhuma aba de Lead Time encontrada: todos os fornecedores usam o lead time padrão "
                                    "(ajuste na aba Parâmetros)."))

    # Capacidade de pallets por fábrica
    if ref["capacidade"]:
        cap = normalizar_capacidade(pl.ler(ref["capacidade"]))
    else:
        fabs = sorted(skus["fabrica"].unique())
        cap = pd.DataFrame({"fabrica": fabs,
                            "capacidade_posicoes": [CAPACIDADE_PADRAO.get(f, np.nan) for f in fabs]})
        avisos.append(("aviso", "Capacidades de pallet por fábrica vieram da aba 'Simulador' (Plenno/LiveMax 50, Bioz 60) — "
                                "confirme na aba Parâmetros."))

    pc = None
    ref_pc = pl.aba_por_nome("Pedidos_Criados", "Pedidos criados")
    if ref_pc:
        pc = normalizar_pedidos_criados(pl.ler(ref_pc))
        usadas["pedidos_criados"] = f"{ref_pc[0]} › {ref_pc[1]}"

    d = Dados(skus=skus, metas=metas, bom=bom, estoque=estoque, pedidos=pedidos, cadastro=cadastro,
              lead_time=lt, capacidade=cap, dups_estoque=dups, pedidos_criados=pc, abas_usadas=usadas, avisos=avisos)
    d.avisos += checar_qualidade(d)
    return d


def aplicar_meta(skus: pd.DataFrame, metas: Optional[pd.DataFrame], mes: Optional[str]) -> Tuple[pd.DataFrame, List[str]]:
    """
    Troca a meta pela de um mês da aba 'Meta de vendas' (casando pela descrição do produto).
    Retorna (skus, lista de SKUs sem correspondência — esses mantêm a meta original).
    """
    if not mes or metas is None or mes not in metas.columns:
        return skus, []
    s = skus.copy()
    novo = s["descricao"].map(norm).map(metas[mes])
    faltantes = s.loc[novo.isna(), "sku"].tolist()
    s["meta_mensal"] = novo.fillna(s["meta_mensal"])
    return s, faltantes


# --------------------------------------------------------------------------- #
# Qualidade dos dados
# --------------------------------------------------------------------------- #
def checar_qualidade(d: Dados) -> List[Tuple[str, str]]:
    av: List[Tuple[str, str]] = []
    s = d.skus
    if (s["un_por_cx"].isna() | (s["un_por_cx"] <= 0)).any():
        lst = ", ".join(s.loc[s["un_por_cx"].isna() | (s["un_por_cx"] <= 0), "sku"])
        av.append(("erro", f"SKUs sem unidades por caixa (não calculam caixas/pallets): {lst}"))
    if (s["cx_por_pallet"] <= 0).any():
        lst = ", ".join(s.loc[s["cx_por_pallet"] <= 0, "sku"])
        av.append(("aviso", f"SKUs sem caixas por pallet (sem arredondamento de pallet): {lst}"))
    if (s["meta_mensal"] <= 0).any():
        lst = ", ".join(s.loc[s["meta_mensal"] <= 0, "sku"])
        av.append(("aviso", f"SKUs sem meta mensal (cobertura não calculável): {lst}"))
    if (s["estoque_un"] < 0).any():
        lst = ", ".join(s.loc[s["estoque_un"] < 0, "sku"])
        av.append(("erro", f"SKUs com estoque negativo: {lst}"))

    if d.bom is not None:
        sem_bom = sorted(set(s["sku"]) - set(d.bom["kit_sku"]))
        if sem_bom:
            av.append(("aviso", f"SKUs sem estrutura (BOM) — não geram necessidade de insumos: {', '.join(sem_bom)}"))
        un_div = (d.bom.groupby("comp_sku")["un"].nunique() > 1)
        if un_div.any():
            av.append(("aviso", "Insumos com unidade diferente entre linhas da BOM (confira conversão): "
                                + ", ".join(un_div[un_div].index)))
        mi = d.bom.loc[d.bom["un"].str.upper() == "MI", "comp_sku"].unique()
        if len(mi):
            av.append(("aviso", "Insumos em MILHEIRO ('MI') na BOM: " + ", ".join(mi) +
                                " — confirme se estoque e pedidos estão em milheiros ou em unidades."))
    else:
        av.append(("erro", "Estrutura de produtos (BOM) não encontrada — o MRP não pode ser calculado."))

    if d.bom is not None and d.estoque is not None:
        sem_est = sorted(set(d.bom["comp_sku"]) - set(d.estoque["comp_sku"]))
        if sem_est:
            av.append(("aviso", f"{len(sem_est)} insumos da BOM não têm linha na planilha de estoque — "
                                f"estoque tratado como ZERO (a compra pode estar superestimada): "
                                + ", ".join(sem_est)))
    if d.estoque is None:
        av.append(("erro", "Estoque de insumos não encontrado — o MRP considera estoque zero."))
    if d.dups_estoque:
        av.append(("info", f"{d.dups_estoque} linha(s) duplicada(s) no estoque de insumos foram removidas "
                           f"(mesmo SKU, mesmo código no terceiro e mesmo saldo)."))

    if d.pedidos is not None:
        p = d.pedidos
        if "tipo" in p.columns:
            n_pa = int((p["tipo"] == "produto_acabado").sum())
            if n_pa:
                av.append(("info", f"{n_pa} linhas de pedido são de produto acabado (Live Max/Plenno) e foram ignoradas "
                                   f"no MRP de insumos."))
            nm = p[p["tipo"] == "nao_mapeado"]
            if len(nm):
                av.append(("aviso", f"{len(nm)} linhas de pedido não casam com nenhum insumo da BOM (ignoradas): "
                                    + "; ".join(sorted(set(nm["item"]))[:8])))
            comp = p[p["tipo"] == "componente"]
            sem_prev = int(comp["prevista"].isna().sum())
            if sem_prev:
                av.append(("aviso", f"{sem_prev} de {len(comp)} pedidos de insumo estão sem Data Prevista — "
                                    f"entram como 'a receber' (sem checagem de prazo)."))
    else:
        av.append(("aviso", "Pedidos em trânsito não encontrados — o MRP considera zero em trânsito."))

    if d.cadastro is not None and d.bom is not None:
        sem_cad = sorted(set(d.bom["comp_sku"]) - set(d.cadastro["comp_sku"]))
        if sem_cad:
            av.append(("aviso", f"Insumos sem fornecedor/preço no cadastro: {', '.join(sem_cad)}"))
        zero = d.cadastro[(d.cadastro["preco"].isna()) | (d.cadastro["preco"] <= 0)]
        zero = zero[zero["comp_sku"].isin(d.bom["comp_sku"])]
        if len(zero):
            av.append(("aviso", f"{len(zero)} insumos com preço zerado/vazio (valor estimado de compra subestimado): "
                                + ", ".join(zero["comp_sku"])))
    return av
