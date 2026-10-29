"""
Engine de cálculo — Simulador de Reposição (PCP) e MRP de Compras.

Módulo puro (sem Streamlit): pode ser usado no app, em testes, em scripts
agendados ou ser portado para Apps Script / Cloud Functions.

Regras (iguais às da planilha "Calculo Adicional" do Dados ponto Reposição.xlsx):

  Venda_Diaria            = Meta_Mensal / 30
  Dias_Cobertura_Atual    = Estoque / Venda_Diaria
  Dias_Cobertura_Pos_Prod = (Estoque + Em_Producao) / Venda_Diaria
  Necessario_Unidades     = MAX(0, ROUNDUP(Dias_Alvo * Venda_Diaria - (Estoque + Em_Producao)))
  Necessario_Caixas       = CEILING(Necessario_Unidades / Un_por_Caixa)
  Caixas_Multiplo_Pallet  = CEILING(Necessario_Caixas, Caixas_por_Pallet)
  Unidades_a_Produzir     = Caixas_Multiplo_Pallet * Un_por_Caixa
  Pallets                 = Caixas_Multiplo_Pallet / Caixas_por_Pallet
  Dias_Cobertura_Final    = ROUND((Estoque + Em_Producao + Unidades_a_Produzir) / Venda_Diaria)

  MRP:
  Consumo_Bruto           = Σ (Unidades_a_Produzir_SKU * Qtd_BOM)
  Necessidade_Liquida     = Consumo_Bruto - (Estoque_Insumo + Pedidos_em_Transito)
  Data_Limite_Pedido      = Data_Necessaria - Lead_Time_do_Fornecedor
"""
from __future__ import annotations

import re
import unicodedata
from datetime import date, timedelta
from typing import Iterable

import numpy as np
import pandas as pd

EPS = 1e-9          # tolera ruído de ponto flutuante antes do ROUNDUP/CEILING
DIAS_MES = 30

UNIDADES_DISCRETAS = {"UND", "UN", "UNT", "PC", "PCT", "CX", "PÇ", "PCS", "MI"}

FAIXA_SEM_META = "Sem meta"
FAIXAS_ORDEM = [
    "Crítico (< 5 dias)",
    "Atenção (5 a 10 dias)",
    "Moderado (10 a 30 dias)",
    "Adequado (30 a 60 dias)",
    "Excesso (> 60 dias)",
    FAIXA_SEM_META,
]


# --------------------------------------------------------------------------- #
# Utilitários
# --------------------------------------------------------------------------- #
def norm(txt) -> str:
    """Normaliza texto para casamento: minúsculo, sem acento, sem espaços duplos."""
    if txt is None or (isinstance(txt, float) and np.isnan(txt)):
        return ""
    s = unicodedata.normalize("NFKD", str(txt))
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = re.sub(r"\s+", " ", s).strip().lower()
    return s


def _ceil(x):
    """ROUNDUP/CEILING do Excel, imune a ruído de ponto flutuante."""
    return np.ceil(np.asarray(x, dtype=float) - EPS) + 0.0


def _round_half_up(x):
    """ROUND do Excel (meio sempre sobe) — diferente do round() do Python."""
    return np.floor(np.asarray(x, dtype=float) + 0.5)


def _num(df: pd.DataFrame, col: str, default=np.nan) -> pd.Series:
    if col not in df.columns:
        return pd.Series(default, index=df.index, dtype=float)
    return pd.to_numeric(df[col], errors="coerce")


def faixa_cobertura(dias: float, tem_meta: bool = True) -> str:
    if not tem_meta or dias is None or (isinstance(dias, float) and np.isnan(dias)):
        return FAIXA_SEM_META
    if dias < 5:
        return FAIXAS_ORDEM[0]
    if dias < 10:
        return FAIXAS_ORDEM[1]
    if dias < 30:
        return FAIXAS_ORDEM[2]
    if dias <= 60:
        return FAIXAS_ORDEM[3]
    return FAIXAS_ORDEM[4]


# --------------------------------------------------------------------------- #
# Módulo A — Simulador de reposição de produtos acabados
# --------------------------------------------------------------------------- #
def simular_reposicao(
    skus: pd.DataFrame,
    dias_alvo: float,
    mult_demanda: float = 1.0,
    descontar_reservado: bool = False,
) -> pd.DataFrame:
    """
    skus: colunas obrigatórias
        sku, descricao, fabrica, un_por_cx, cx_por_pallet,
        estoque_un, em_producao, meta_mensal
      opcionais
        reservado_cx (descontado do estoque se descontar_reservado=True)
        ajuste_cx    (se preenchido, substitui as caixas calculadas — sem arredondar pallet)
    """
    df = skus.copy().reset_index(drop=True)

    un_cx = _num(df, "un_por_cx")
    un_cx = un_cx.where(un_cx > 0)                      # 0/NaN = dado inválido
    cxp = _num(df, "cx_por_pallet").fillna(0).clip(lower=0)
    est = _num(df, "estoque_un").fillna(0)
    prod = _num(df, "em_producao").fillna(0)
    meta = _num(df, "meta_mensal").fillna(0) * float(mult_demanda)
    ajuste = _num(df, "ajuste_cx")

    if descontar_reservado:
        reservado_un = _num(df, "reservado_cx").fillna(0) * un_cx.fillna(1)
        est = (est - reservado_un).clip(lower=0)

    venda = meta / DIAS_MES
    tem_meta = venda > 0
    venda_safe = venda.where(tem_meta)                  # NaN onde não há meta

    cob_atual = est / venda_safe
    cob_pos = (est + prod) / venda_safe

    # Necessidade para atingir os dias-alvo
    nec_un = np.maximum(0, _ceil(dias_alvo * meta / DIAS_MES - (est + prod)))
    nec_un = pd.Series(nec_un, index=df.index)
    nec_cx = pd.Series(_ceil(nec_un / un_cx), index=df.index)          # NaN se un_cx inválido

    # Múltiplo de pallet (sem pallet cadastrado → mantém as caixas)
    cxp_safe = cxp.where(cxp > 0, 1)
    cx_mult = pd.Series(
        np.where(cxp > 0, _ceil(nec_cx / cxp_safe) * cxp, nec_cx), index=df.index
    )
    cx_final = cx_mult.where(ajuste.isna(), ajuste)
    un_final = cx_final * un_cx
    un_pallet = un_cx * cxp.where(cxp > 0)
    pallets = (cx_final / cxp.where(cxp > 0))

    cob_final = pd.Series(_round_half_up((est + prod + un_final) / venda_safe), index=df.index)

    # Posições de pallet (ROUNDUP, cada local fecha a sua própria posição)
    def _pos(un):
        return pd.Series(np.where(un_pallet > 0, _ceil(un / un_pallet.where(un_pallet > 0, 1)), np.nan),
                         index=df.index)

    pos_estoque = _pos(est)
    pos_em_prod = _pos(prod)
    pos_producao = _pos(un_final.fillna(0))

    # Alertas por linha
    alertas = []
    for i in df.index:
        a = []
        if pd.isna(un_cx.iloc[i]):
            a.append("sem unid./caixa")
        if cxp.iloc[i] <= 0:
            a.append("sem pallet cadastrado")
        if not tem_meta.iloc[i]:
            a.append("sem meta")
        if est.iloc[i] < 0:
            a.append("estoque negativo")
        alertas.append("; ".join(a))

    out = pd.DataFrame({
        "sku": df["sku"],
        "descricao": df["descricao"],
        "fabrica": df["fabrica"],
        "un_por_cx": un_cx,
        "cx_por_pallet": cxp.where(cxp > 0),
        "un_por_pallet": un_pallet,
        "estoque_un": est,
        "estoque_cx": est / un_cx,
        "em_producao": prod,
        "meta_mensal": meta,
        "venda_diaria": venda,
        "cob_atual_dias": cob_atual,
        "cob_pos_prod_dias": cob_pos,
        "faixa": [faixa_cobertura(c, t) for c, t in zip(cob_atual, tem_meta)],
        "dias_alvo": float(dias_alvo),
        "nec_un": nec_un,
        "nec_cx": nec_cx,
        "cx_mult_pallet": cx_final,
        "un_a_produzir": un_final,
        "pallets": pallets,
        "cob_final_dias": cob_final,
        "pos_estoque": pos_estoque,
        "pos_em_producao": pos_em_prod,
        "pos_producao": pos_producao,
        "ajuste_manual": ajuste.notna(),
        "alerta_dados": alertas,
    })
    return out


# --------------------------------------------------------------------------- #
# Armazém / capacidade de pallets
# --------------------------------------------------------------------------- #
BASES_OCUPACAO = {
    "producao": "Somente produção simulada",
    "producao+op": "Produção simulada + em produção",
    "total": "Estoque atual + em produção + simulada",
}


def ocupacao_armazem(sim: pd.DataFrame, capacidades: pd.DataFrame, base: str = "producao") -> pd.DataFrame:
    """
    capacidades: colunas fabrica, capacidade_posicoes.
    Retorna uma linha por fábrica com ocupação, capacidade, saldo e status.
    """
    g = sim.groupby("fabrica", dropna=False).agg(
        pos_estoque=("pos_estoque", "sum"),
        pos_em_producao=("pos_em_producao", "sum"),
        pos_producao=("pos_producao", "sum"),
        skus_a_produzir=("un_a_produzir", lambda s: int((s.fillna(0) > 0).sum())),
    ).reset_index()

    if base == "producao":
        g["ocupacao"] = g["pos_producao"]
    elif base == "producao+op":
        g["ocupacao"] = g["pos_producao"] + g["pos_em_producao"]
    else:
        g["ocupacao"] = g["pos_producao"] + g["pos_em_producao"] + g["pos_estoque"]

    cap = capacidades.rename(columns=str).copy()
    g = g.merge(cap[["fabrica", "capacidade_posicoes"]], on="fabrica", how="left")
    g["saldo"] = g["capacidade_posicoes"] - g["ocupacao"]

    def _status(r):
        if pd.isna(r["capacidade_posicoes"]):
            return "SEM CAPACIDADE INFORMADA"
        if r["ocupacao"] > r["capacidade_posicoes"]:
            return "ESTOURA CAPACIDADE"
        if r["ocupacao"] >= 0.9 * r["capacidade_posicoes"]:
            return "ATENÇÃO (≥ 90%)"
        return "OK"

    g["status"] = g.apply(_status, axis=1)
    g["uso_pct"] = g["ocupacao"] / g["capacidade_posicoes"]
    return g


# --------------------------------------------------------------------------- #
# Módulo B — MRP
# --------------------------------------------------------------------------- #
def classificar_pedidos(pedidos: pd.DataFrame, bom: pd.DataFrame) -> pd.DataFrame:
    """
    Liga cada linha de pedido a um componente da BOM (por ID e, se falhar, pela descrição)
    e classifica: 'componente' | 'produto_acabado' | 'nao_mapeado'.

    pedidos: oc, data, prevista, fornecedor, item, qtd, cod_id (opcional)
    bom:     comp_sku, comp_id, comp_desc
    """
    p = pedidos.copy().reset_index(drop=True)
    comps = bom[["comp_sku", "comp_id", "comp_desc"]].drop_duplicates("comp_sku")
    by_id = {str(int(i)): s for i, s in zip(pd.to_numeric(comps["comp_id"], errors="coerce"), comps["comp_sku"])
             if pd.notna(i)}
    by_desc = {norm(d): s for d, s in zip(comps["comp_desc"], comps["comp_sku"])}

    comp_sku, tipo = [], []
    for _, r in p.iterrows():
        sku = None
        cid = r.get("cod_id")
        try:
            sku = by_id.get(str(int(float(cid))))
        except (TypeError, ValueError):
            sku = None
        if sku is None:
            sku = by_desc.get(norm(r.get("item")))
        if sku is not None:
            comp_sku.append(sku)
            tipo.append("componente")
        else:
            comp_sku.append(None)
            nome = norm(r.get("item"))
            tipo.append("produto_acabado" if ("bioz green" in nome or "weasy" in nome) else "nao_mapeado")
    p["comp_sku"] = comp_sku
    p["tipo"] = tipo
    return p


def consumo_detalhado(producao: pd.DataFrame, bom: pd.DataFrame) -> pd.DataFrame:
    """Consumo bruto por (insumo, SKU acabado). producao: sku, un_a_produzir."""
    prod = producao[["sku", "un_a_produzir"]].copy()
    prod = prod[prod["un_a_produzir"].fillna(0) > 0]
    d = bom.merge(prod, left_on="kit_sku", right_on="sku", how="inner")
    d["consumo"] = d["qtd"] * d["un_a_produzir"]
    return d[["comp_sku", "comp_desc", "un", "kit_sku", "un_a_produzir", "qtd", "consumo"]]


def calcular_mrp(
    producao: pd.DataFrame,
    bom: pd.DataFrame,
    estoque: pd.DataFrame,
    pedidos: pd.DataFrame,
    cadastro: pd.DataFrame,
    parametros_forn: pd.DataFrame,
    data_necessaria: date,
    hoje: date,
    lt_padrao: int = 30,
    transito_ate_data: bool = True,
    fabricas: Iterable[str] | None = None,
) -> pd.DataFrame:
    """
    producao:        sku, fabrica, un_a_produzir      (saída de simular_reposicao)
    bom:             kit_sku, comp_sku, comp_id, comp_desc, qtd, un
    estoque:         comp_sku, estoque                (já deduplicado)
    pedidos:         saída de classificar_pedidos     (usa tipo == 'componente')
    cadastro:        comp_sku, descricao, unidade, preco, fornecedor
    parametros_forn: fornecedor, lead_time_dias, lote_minimo, multiplo
    """
    prod = producao.copy()
    if fabricas is not None:
        prod = prod[prod["fabrica"].isin(list(fabricas))]

    det = consumo_detalhado(prod, bom)
    cols = ["comp_sku", "descricao", "unidade", "fornecedor", "consumo_bruto", "n_skus", "estoque",
            "transito_total", "transito_considerado", "transito_fora_prazo", "chegada_prevista",
            "necessidade_liquida", "qtd_comprar", "qtd_sugerida", "preco", "valor_estimado",
            "lead_time_dias", "lead_time_padrao", "data_necessaria", "data_limite_pedido",
            "dias_ate_limite", "status", "sem_estoque_cadastrado", "sem_fornecedor", "sem_preco"]
    if det.empty:
        return pd.DataFrame(columns=cols)

    base = det.groupby("comp_sku").agg(
        descricao_bom=("comp_desc", "first"),
        unidade_bom=("un", "first"),
        consumo_bruto=("consumo", "sum"),
        n_skus=("kit_sku", "nunique"),
    ).reset_index()

    # Estoque do insumo
    est = estoque.groupby("comp_sku", as_index=False)["estoque"].sum()
    base = base.merge(est, on="comp_sku", how="left")
    base["sem_estoque_cadastrado"] = base["estoque"].isna()
    base["estoque"] = base["estoque"].fillna(0.0)

    # Pedidos em trânsito
    ped = pedidos[pedidos["tipo"] == "componente"].copy()
    ped["qtd"] = pd.to_numeric(ped["qtd"], errors="coerce").fillna(0)
    ped["prevista"] = pd.to_datetime(ped["prevista"], errors="coerce")
    limite = pd.Timestamp(data_necessaria)
    if transito_ate_data:
        no_prazo = ped["prevista"].isna() | (ped["prevista"] <= limite)   # sem data prevista = conta
    else:
        no_prazo = pd.Series(True, index=ped.index)
    ped["qtd_prazo"] = ped["qtd"].where(no_prazo, 0.0)
    tr = ped.groupby("comp_sku").agg(
        transito_total=("qtd", "sum"),
        transito_considerado=("qtd_prazo", "sum"),
        chegada_prevista=("prevista", "max"),
    ).reset_index()
    tr["transito_fora_prazo"] = tr["transito_total"] - tr["transito_considerado"]
    base = base.merge(tr, on="comp_sku", how="left")
    for c in ("transito_total", "transito_considerado", "transito_fora_prazo"):
        base[c] = base[c].fillna(0.0)

    # Cadastro (fornecedor, preço, unidade)
    cad = cadastro.drop_duplicates("comp_sku")[["comp_sku", "descricao", "unidade", "preco", "fornecedor"]]
    base = base.merge(cad, on="comp_sku", how="left")
    base["descricao"] = base["descricao"].replace("", np.nan).fillna(base["descricao_bom"])
    base["unidade"] = base["unidade_bom"].replace("", np.nan).fillna(base["unidade"])
    base["sem_fornecedor"] = base["fornecedor"].isna()
    base["fornecedor"] = base["fornecedor"].fillna("(sem fornecedor)")
    base["preco"] = pd.to_numeric(base["preco"], errors="coerce")
    base["sem_preco"] = base["preco"].isna() | (base["preco"] <= 0)

    # Lead time / lote mínimo / múltiplo por fornecedor
    pf = parametros_forn.copy()
    pf["_k"] = pf["fornecedor"].map(norm)
    pf = pf.drop_duplicates("_k").set_index("_k")
    key = base["fornecedor"].map(norm)
    lt = key.map(pf["lead_time_dias"]) if len(pf) else pd.Series(np.nan, index=base.index)
    base["lead_time_padrao"] = pd.to_numeric(lt, errors="coerce").isna()
    base["lead_time_dias"] = pd.to_numeric(lt, errors="coerce").fillna(lt_padrao).astype(int)
    lote = pd.to_numeric(key.map(pf["lote_minimo"]), errors="coerce").fillna(0) if len(pf) else 0
    mult = pd.to_numeric(key.map(pf["multiplo"]), errors="coerce").fillna(0) if len(pf) else 0

    # Necessidade líquida e quantidade a comprar
    base["necessidade_liquida"] = base["consumo_bruto"] - (base["estoque"] + base["transito_considerado"])
    base["qtd_comprar"] = base["necessidade_liquida"].clip(lower=0)

    # Itens contáveis (UND/PC...) não se compram fracionados: arredonda para cima
    discreta = base["unidade"].astype(str).str.upper().str.strip().isin(UNIDADES_DISCRETAS)
    sug = pd.Series(np.where(discreta, _ceil(base["qtd_comprar"]), base["qtd_comprar"]), index=base.index)
    m = pd.Series(mult, index=base.index) if np.isscalar(mult) else mult
    l = pd.Series(lote, index=base.index) if np.isscalar(lote) else lote
    sug = pd.Series(np.where(m > 0, _ceil(sug / m.where(m > 0, 1)) * m, sug), index=base.index)
    sug = pd.Series(np.where((l > 0) & (sug > 0), np.maximum(sug, l), sug), index=base.index)
    base["qtd_sugerida"] = sug
    base["valor_estimado"] = base["qtd_sugerida"] * base["preco"].fillna(0)

    # Datas
    nec = pd.Timestamp(data_necessaria)
    base["data_necessaria"] = nec
    base["data_limite_pedido"] = nec - pd.to_timedelta(base["lead_time_dias"], unit="D")
    base["dias_ate_limite"] = (base["data_limite_pedido"] - pd.Timestamp(hoje)).dt.days

    def _status(r):
        if r["qtd_comprar"] <= 0:
            return "COBERTO"
        if r["dias_ate_limite"] < 0:
            return "ATRASADO"
        if r["dias_ate_limite"] <= 7:
            return "PEDIR ESTA SEMANA"
        return "PROGRAMAR"

    base["status"] = base.apply(_status, axis=1)
    ordem = {"ATRASADO": 0, "PEDIR ESTA SEMANA": 1, "PROGRAMAR": 2, "COBERTO": 3}
    base = base.sort_values(
        by=["status", "data_limite_pedido", "fornecedor"], key=lambda c: c.map(ordem) if c.name == "status" else c
    ).reset_index(drop=True)
    return base[cols]


# --------------------------------------------------------------------------- #
# Cenários lado a lado (60 / 75 / 80 dias...)
# --------------------------------------------------------------------------- #
def comparar_cenarios(
    skus: pd.DataFrame,
    dias_list: Iterable[float],
    mrp_kwargs: dict,
    mult_demanda: float = 1.0,
    descontar_reservado: bool = False,
) -> pd.DataFrame:
    linhas = []
    for d in dias_list:
        sim = simular_reposicao(skus, d, mult_demanda, descontar_reservado)
        mrp = calcular_mrp(sim, **mrp_kwargs)
        linhas.append({
            "dias_alvo": d,
            "skus_a_produzir": int((sim["un_a_produzir"].fillna(0) > 0).sum()),
            "unidades": float(sim["un_a_produzir"].sum()),
            "caixas": float(sim["cx_mult_pallet"].sum()),
            "pallets": float(sim["pallets"].sum()),
            "posicoes_pallet": float(sim["pos_producao"].sum()),
            "insumos_a_comprar": int((mrp["qtd_comprar"] > 0).sum()) if len(mrp) else 0,
            "insumos_atrasados": int((mrp["status"] == "ATRASADO").sum()) if len(mrp) else 0,
            "valor_compras": float(mrp["valor_estimado"].sum()) if len(mrp) else 0.0,
        })
    return pd.DataFrame(linhas)


# --------------------------------------------------------------------------- #
# Programação viável: o que o terceiro JÁ consegue produzir com o material que tem
# --------------------------------------------------------------------------- #
def producao_viavel(
    sim: pd.DataFrame,
    bom: pd.DataFrame,
    estoque_local: pd.DataFrame,
    fabrica: str | None = None,
    pallet_fechado: bool = True,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Dado o plano (sim = saída de simular_reposicao) e o saldo de insumos que o terceiro tem em mãos
    (estoque_local: comp_sku, estoque), calcula quanto de cada SKU dá para produzir AGORA.

    Os SKUs são atendidos por prioridade (coluna `prioridade` se existir; senão menor cobertura primeiro);
    cada um consome o saldo compartilhado, então o que um produz deixa de existir para o próximo.
    Arredondamentos: sempre em CAIXAS fechadas e, com pallet_fechado=True (padrão) e pallet cadastrado
    (coluna cx_por_pallet), em PALLETS fechados — exceto quando o material cobre o plano inteiro.

    Retorna (por_sku, por_insumo):
      por_sku:    sku, descricao, fabrica, cob_atual_dias, plano_cx, plano_un, viavel_cx, viavel_un, falta_un,
                  pct_viavel, gargalo, status  (PRODUZ TUDO | PRODUZ PARTE | NÃO PRODUZ | SEM ESTRUTURA)
      por_insumo: comp_sku, descricao, unidade, necessario_plano, saldo_terceiro, consumo_viavel,
                  saldo_apos_viavel, falta_para_plano   (falta_para_plano = o que Bioz precisa enviar/comprar)
    """
    df = sim.copy()
    if fabrica is not None:
        df = df[df["fabrica"] == fabrica]
    df = df[df["un_a_produzir"].fillna(0) > 0].copy()
    if "prioridade" in df.columns and df["prioridade"].notna().any():          # prioridade definida pelo usuário
        df["_ord"] = pd.to_numeric(df["prioridade"], errors="coerce").fillna(1e9)
    else:                                                                      # padrão: menor cobertura primeiro
        df["_ord"] = df["cob_atual_dias"].fillna(1e9)
    df = df.sort_values(["_ord", "sku"])

    saldo0 = estoque_local.groupby("comp_sku")["estoque"].sum().to_dict()
    saldo = dict(saldo0)
    desc = bom.drop_duplicates("comp_sku").set_index("comp_sku")[["comp_desc", "un"]]
    estrut = {k: g[["comp_sku", "qtd"]].values.tolist() for k, g in bom.groupby("kit_sku")}

    necessario: dict = {}
    consumido: dict = {}
    linhas = []
    for r in df.itertuples(index=False):
        comps = estrut.get(r.sku)
        un_cx = float(r.un_por_cx) if pd.notna(r.un_por_cx) and r.un_por_cx > 0 else np.nan
        plano_un = float(r.un_a_produzir)
        plano_cx = plano_un / un_cx if un_cx == un_cx else np.nan
        if not comps or un_cx != un_cx:
            linhas.append(dict(sku=r.sku, descricao=r.descricao, fabrica=r.fabrica, cob_atual_dias=r.cob_atual_dias,
                               plano_cx=plano_cx, plano_un=plano_un, viavel_cx=0.0, viavel_un=0.0, falta_un=plano_un,
                               pct_viavel=0.0, gargalo="", status="SEM ESTRUTURA", cx_por_pallet=np.nan))
            continue
        # insumos necessários para o plano completo
        for c, q in comps:
            necessario[c] = necessario.get(c, 0.0) + q * plano_un
        # quanto cada insumo permite produzir (em unidades)
        lim_un, gargalo = np.inf, ""
        for c, q in comps:
            if q <= 0:
                continue
            pode = np.floor(saldo.get(c, 0.0) / q + EPS)
            if pode < lim_un:
                lim_un, gargalo = pode, c
        cxp = getattr(r, "cx_por_pallet", np.nan)
        cxp = float(cxp) if pd.notna(cxp) and cxp > 0 else np.nan
        viavel_cx = float(min(plano_cx, np.floor(lim_un / un_cx + EPS))) if np.isfinite(lim_un) else float(plano_cx)
        viavel_cx = max(viavel_cx, 0.0)
        if pallet_fechado and cxp == cxp and viavel_cx < plano_cx - EPS:      # material não cobre tudo → pallets fechados
            viavel_cx = float(np.floor(viavel_cx / cxp + EPS) * cxp)
        viavel_un = viavel_cx * un_cx
        for c, q in comps:
            saldo[c] = saldo.get(c, 0.0) - q * viavel_un
            consumido[c] = consumido.get(c, 0.0) + q * viavel_un
        falta = plano_un - viavel_un
        status = "PRODUZ TUDO" if falta <= EPS else ("NÃO PRODUZ" if viavel_un <= 0 else "PRODUZ PARTE")
        linhas.append(dict(sku=r.sku, descricao=r.descricao, fabrica=r.fabrica, cob_atual_dias=r.cob_atual_dias,
                           plano_cx=plano_cx, plano_un=plano_un, viavel_cx=viavel_cx, viavel_un=viavel_un,
                           falta_un=falta, pct_viavel=viavel_un / plano_un if plano_un else 1.0,
                           gargalo="" if status == "PRODUZ TUDO" else gargalo, status=status, cx_por_pallet=cxp))
    por_sku = pd.DataFrame(linhas, columns=["sku", "descricao", "fabrica", "cob_atual_dias", "plano_cx", "plano_un",
                                            "viavel_cx", "viavel_un", "falta_un", "pct_viavel", "gargalo", "status",
                                            "cx_por_pallet"])
    ins = []
    for c, nec in necessario.items():
        s0 = saldo0.get(c, 0.0)
        ins.append(dict(comp_sku=c, descricao=desc["comp_desc"].get(c, ""), unidade=desc["un"].get(c, ""),
                        necessario_plano=nec, saldo_terceiro=s0, consumo_viavel=consumido.get(c, 0.0),
                        saldo_apos_viavel=saldo.get(c, s0), falta_para_plano=max(0.0, nec - s0),
                        sem_saldo_informado=c not in saldo0))
    por_insumo = pd.DataFrame(ins, columns=["comp_sku", "descricao", "unidade", "necessario_plano", "saldo_terceiro",
                                            "consumo_viavel", "saldo_apos_viavel", "falta_para_plano",
                                            "sem_saldo_informado"])
    por_insumo = por_insumo.sort_values("falta_para_plano", ascending=False).reset_index(drop=True)
    return por_sku, por_insumo


def cronograma_terceiro(
    sim: pd.DataFrame,
    bom: pd.DataFrame,
    estoque_local: pd.DataFrame,
    entradas: pd.DataFrame | None,
    fabrica: str,
    hoje,
    dias_apos_chegada: int = 0,
    pallet_fechado: bool = True,
):
    """
    Simula QUANDO cada SKU pode ser produzido no terceiro, em "ondas":
      onda 0 = hoje, com o saldo que ele já tem;  ondas seguintes = cada data de chegada de material
      (pedidos de compra criados: entradas = comp_sku, qtd, data).
    Em cada onda produz-se, na ordem de prioridade, o que o saldo permitir; o resto espera a próxima chegada.

    Retorna (por_sku, ondas, faltas_sem_pedido):
      por_sku: sku, descricao, plano_cx, agora_cx, depois_cx, sem_pedido_cx, primeira_data, ultima_data, status
      ondas:   data, sku, descricao, cx, un           (o que é produzido em cada data)
      faltas:  comp_sku, descricao, unidade, falta    (material do que ainda NÃO tem pedido que cubra)
    """
    hoje = pd.Timestamp(hoje).normalize()
    plano = sim[(sim["fabrica"] == fabrica) & (sim["un_a_produzir"].fillna(0) > 0)].copy()
    restante = plano.set_index("sku")["un_a_produzir"].astype(float).to_dict()
    saldo = estoque_local.groupby("comp_sku")["estoque"].sum().to_dict()
    consumo_un = {k: g[["comp_sku", "qtd"]].values.tolist() for k, g in bom.groupby("kit_sku")}

    chegadas = []
    if entradas is not None and len(entradas):
        en = entradas.copy()
        en["data"] = pd.to_datetime(en["data"], errors="coerce").dt.normalize()
        en["qtd"] = pd.to_numeric(en["qtd"], errors="coerce").fillna(0)
        en = en.dropna(subset=["data", "comp_sku"])
        en["data"] = en["data"].where(en["data"] > hoje, hoje)                 # já chegou = disponível hoje
        chegadas = sorted(en.groupby("data").apply(lambda g: g[["comp_sku", "qtd"]].values.tolist()).items())
    datas = [hoje] + [d for d, _ in chegadas if d != hoje]
    chegou = {d: itens for d, itens in chegadas}
    if hoje in chegou:
        for c, q in chegou[hoje]:
            saldo[c] = saldo.get(c, 0.0) + q

    ondas = []
    for d in datas:
        if d != hoje:
            for c, q in chegou.get(d, []):
                saldo[c] = saldo.get(c, 0.0) + q
        pend = plano[plano["sku"].map(lambda k: restante.get(k, 0) > EPS)].copy()
        if pend.empty:
            break
        pend["un_a_produzir"] = pend["sku"].map(restante)
        est = pd.DataFrame({"comp_sku": list(saldo), "estoque": list(saldo.values())})
        r, _ = producao_viavel(pend, bom, est, fabrica, pallet_fechado)
        for x in r[r["viavel_un"] > 0].itertuples(index=False):
            ondas.append(dict(data=d + pd.Timedelta(days=dias_apos_chegada if d != hoje else 0), onda=d,
                              sku=x.sku, descricao=x.descricao, cx=x.viavel_cx, un=x.viavel_un,
                              pallets=(x.viavel_cx / x.cx_por_pallet) if pd.notna(x.cx_por_pallet) else np.nan))
            restante[x.sku] -= x.viavel_un
            for c, q in consumo_un.get(x.sku, []):
                saldo[c] = saldo.get(c, 0.0) - q * x.viavel_un
    ondas_df = pd.DataFrame(ondas, columns=["data", "onda", "sku", "descricao", "cx", "un", "pallets"])

    linhas = []
    for p in plano.itertuples(index=False):
        un_cx = float(p.un_por_cx) if pd.notna(p.un_por_cx) and p.un_por_cx > 0 else np.nan
        o = ondas_df[ondas_df["sku"] == p.sku]
        agora = float(o[o["onda"] == hoje]["cx"].sum())
        depois = float(o[o["onda"] != hoje]["cx"].sum())
        plano_cx = float(p.un_a_produzir) / un_cx if un_cx == un_cx else np.nan
        sem = max(0.0, restante.get(p.sku, 0.0) / un_cx) if un_cx == un_cx else np.nan
        tem_bom = p.sku in consumo_un
        prim = o[o["onda"] != hoje]["data"].min() if depois > 0 else pd.NaT
        ult = o["data"].max() if len(o) else pd.NaT
        if not tem_bom:
            st = "SEM ESTRUTURA"
        elif sem and sem > EPS:
            st = "SEM PEDIDO QUE COBRE" if (agora + depois) <= EPS else "PARTE SEM PEDIDO"
        elif depois <= EPS:
            st = "PRODUZ TUDO AGORA"
        elif agora > EPS:
            st = "PARTE AGORA, RESTO COM PEDIDO"
        else:
            st = "AGUARDA MATERIAL"
        cxp_p = getattr(p, "cx_por_pallet", np.nan)
        linhas.append(dict(sku=p.sku, descricao=p.descricao, cob_atual_dias=p.cob_atual_dias, plano_cx=plano_cx,
                           cx_por_pallet=cxp_p if pd.notna(cxp_p) and cxp_p > 0 else np.nan,
                           agora_cx=agora, depois_cx=depois, sem_pedido_cx=sem, primeira_data=prim, ultima_data=ult,
                           status=st))
    por_sku = pd.DataFrame(linhas)
    if len(por_sku):                                                            # mantém a ordem de prioridade usada
        ordem = {s: i for i, s in enumerate(
            (plano.assign(_o=pd.to_numeric(plano["prioridade"], errors="coerce") if "prioridade" in plano.columns
                          else plano["cob_atual_dias"]).sort_values(["_o", "sku"], na_position="last")["sku"]))}
        por_sku = por_sku.sort_values("sku", key=lambda c: c.map(ordem)).reset_index(drop=True)

    # material que ainda falta para o que não tem pedido
    pend = plano[plano["sku"].map(lambda k: restante.get(k, 0) > EPS)].copy()
    faltas = pd.DataFrame(columns=["comp_sku", "descricao", "unidade", "falta"])
    if len(pend):
        pend["un_a_produzir"] = pend["sku"].map(restante)
        est = pd.DataFrame({"comp_sku": list(saldo), "estoque": [max(v, 0.0) for v in saldo.values()]})
        _, ins = producao_viavel(pend, bom, est, fabrica, pallet_fechado)
        f = ins[ins["falta_para_plano"] > EPS]
        faltas = f[["comp_sku", "descricao", "unidade", "falta_para_plano"]].rename(columns={"falta_para_plano": "falta"})
    return por_sku, ondas_df, faltas.reset_index(drop=True)
