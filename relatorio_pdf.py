"""PDF da programação de produção para enviar ao terceiro (identidade visual Bioz Green)."""
from __future__ import annotations

import io
from pathlib import Path

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import Image, KeepTogether, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

VERDE = colors.HexColor("#0B592F")
VERDE_CLARO = colors.HexColor("#8CA70A")
AMARELO = colors.HexColor("#EEB627")
LARANJA = colors.HexColor("#DF6F3B")
CREME = colors.HexColor("#F0E7D7")
LINHA = colors.HexColor("#D9D9D9")
CINZA = colors.HexColor("#3D3D3D")
LOGO = Path(__file__).parent / "assets" / "logo-bioz-green-small.png"

P = ParagraphStyle("p", fontName="Helvetica", fontSize=8.5, leading=10.5, textColor=CINZA)
PB = ParagraphStyle("pb", parent=P, fontName="Helvetica-Bold")
PH = ParagraphStyle("ph", parent=P, fontName="Helvetica-Bold", textColor=colors.white, alignment=1)
H1 = ParagraphStyle("h1", fontName="Helvetica-Bold", fontSize=18, leading=22, textColor=VERDE)
H2 = ParagraphStyle("h2", fontName="Helvetica-Bold", fontSize=11.5, leading=14, textColor=VERDE, spaceBefore=10, spaceAfter=4)
SUB = ParagraphStyle("sub", fontName="Helvetica-Bold", fontSize=9.5, textColor=VERDE_CLARO)
SMALL = ParagraphStyle("small", parent=P, fontSize=7.5, textColor=colors.HexColor("#595959"))


def _n(v, casas=0) -> str:
    try:
        s = f"{float(v):,.{casas}f}"
    except (TypeError, ValueError):
        return "—"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def _d(v) -> str:
    return "—" if pd.isna(v) else pd.Timestamp(v).strftime("%d/%m/%Y")


def _tabela(cab, linhas, larguras, alinhar_dir=()):
    dados = [[Paragraph(c, PH) for c in cab]] + linhas
    t = Table(dados, colWidths=larguras, repeatRows=1)
    est = [("BACKGROUND", (0, 0), (-1, 0), VERDE), ("GRID", (0, 0), (-1, -1), 0.4, LINHA),
           ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
           ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#FAF7F0")])]
    for j in alinhar_dir:
        est.append(("ALIGN", (j, 1), (j, -1), "RIGHT"))
    t.setStyle(TableStyle(est))
    return t


def gerar_pdf_programacao(fabrica: str, por_sku: pd.DataFrame, ondas: pd.DataFrame, faltas: pd.DataFrame,
                          hoje, meta: dict | None = None, observacao: str = "") -> bytes:
    """
    por_sku/ondas/faltas = saída de engine.cronograma_terceiro. meta: {"Horizonte": "40 dias", "Saldo de material": "02/10/2026"…}
    """
    hoje = pd.Timestamp(hoje).normalize()
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=landscape(A4), leftMargin=14 * mm, rightMargin=14 * mm, topMargin=12 * mm,
                            bottomMargin=14 * mm, title=f"Programação de produção — {fabrica}", author="Bioz Green — Compras")
    W = landscape(A4)[0] - 28 * mm
    el = []

    # Cabeçalho de marca
    logo = Image(str(LOGO), width=17 * mm, height=13.6 * mm) if LOGO.exists() else Spacer(1, 1)
    cab = Table([[logo, [Paragraph(f"Programação de produção — {fabrica}", H1), Paragraph("Bioz Green — Pelo Planeta", SUB)],
                  Paragraph(f"<b>Emissão:</b> {hoje:%d/%m/%Y}<br/>" + "<br/>".join(f"<b>{k}:</b> {v}" for k, v in (meta or {}).items()), P)]],
                colWidths=[22 * mm, W - 22 * mm - 80 * mm, 80 * mm])
    cab.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("LEFTPADDING", (0, 0), (-1, -1), 0)]))
    faixa = Table([[""]], colWidths=[W], rowHeights=[3 * mm])
    faixa.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), VERDE)]))
    fio = Table([[""]], colWidths=[W], rowHeights=[1 * mm])
    fio.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), VERDE_CLARO)]))
    el += [faixa, fio, Spacer(1, 3 * mm), cab, Spacer(1, 2 * mm)]

    if observacao.strip():
        o = Table([[Paragraph(observacao.strip().replace("\n", "<br/>"), P)]], colWidths=[W])
        o.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), CREME), ("BOX", (0, 0), (-1, -1), 0.5, LINHA),
                               ("LEFTPADDING", (0, 0), (-1, -1), 6), ("TOPPADDING", (0, 0), (-1, -1), 5), ("BOTTOMPADDING", (0, 0), (-1, -1), 5)]))
        el += [o, Spacer(1, 2 * mm)]

    ordem = {s: i + 1 for i, s in enumerate(por_sku["sku"])} if len(por_sku) else {}
    onda0 = ondas[ondas["onda"] == hoje] if len(ondas) else ondas
    plano_cx = por_sku.set_index("sku")["plano_cx"].to_dict() if len(por_sku) else {}

    # 1) produzir agora
    el.append(Paragraph("1. Pode produzir agora — material já disponível no estoque", H2))
    if len(onda0):
        ag = onda0.groupby(["sku", "descricao"], sort=False, as_index=False).agg(cx=("cx", "sum"), pallets=("pallets", "sum"),
                                                                                  un=("un", "sum"))
        ag["ord"] = ag["sku"].map(ordem)
        ag = ag.sort_values("ord")
        linhas = []
        for r in ag.itertuples(index=False):
            pl = plano_cx.get(r.sku, r.cx)
            obs = "Programação completa" if r.cx >= pl - 1e-6 else f"Parcial: {_n(r.cx)} de {_n(pl)} cx (restante aguarda material)"
            linhas.append([Paragraph(str(r.ord), P), Paragraph(r.sku, PB), Paragraph(r.descricao, P),
                           Paragraph(_n(r.cx), PB), Paragraph(_n(r.pallets) if pd.notna(r.pallets) else "—", P),
                           Paragraph(_n(r.un), P), Paragraph(obs, P)])
        el.append(_tabela(["Prioridade", "SKU", "Descrição", "Caixas", "Pallets", "Unidades", "Observação"], linhas,
                          [20 * mm, 30 * mm, W - 20 * mm - 30 * mm - 20 * mm - 18 * mm - 24 * mm - 70 * mm, 20 * mm, 18 * mm,
                           24 * mm, 70 * mm], (3, 4, 5)))
        el.append(Paragraph(f"Total: <b>{_n(ag['cx'].sum())}</b> caixas · <b>{_n(ag['pallets'].sum())}</b> pallets · "
                            f"<b>{_n(ag['un'].sum())}</b> unidades. Produção em caixas e pallets fechados, na ordem de prioridade acima.", SMALL))
    else:
        el.append(Paragraph("Nenhum item pode ser produzido agora com o saldo de material informado.", P))

    # 2) aguardando material (com data)
    depois = ondas[ondas["onda"] != hoje] if len(ondas) else ondas
    el.append(Paragraph("2. Libera conforme a chegada do material (pedidos de compra emitidos)", H2))
    if len(depois):
        d2 = depois.sort_values(["data", "sku"])
        linhas = [[Paragraph(_d(r.data), PB), Paragraph(r.sku, PB), Paragraph(r.descricao, P), Paragraph(_n(r.cx), PB),
                   Paragraph(_n(r.pallets) if pd.notna(r.pallets) else "—", P), Paragraph(_n(r.un), P)]
                  for r in d2.itertuples(index=False)]
        el.append(_tabela(["Pode iniciar a partir de", "SKU", "Descrição", "Caixas", "Pallets", "Unidades"], linhas,
                          [38 * mm, 30 * mm, W - 38 * mm - 30 * mm - 22 * mm - 18 * mm - 26 * mm, 22 * mm, 18 * mm, 26 * mm],
                          (3, 4, 5)))
        el.append(Paragraph("As datas são previsões baseadas nas datas de entrega dos pedidos de compra; serão atualizadas conforme confirmação.", SMALL))
    else:
        el.append(Paragraph("Nenhum item aguardando material com pedido já emitido.", P))

    # 3) sem pedido
    pend = por_sku[por_sku["sem_pedido_cx"].fillna(0) > 1e-9] if len(por_sku) else por_sku
    el.append(Paragraph("3. Ainda sem material previsto (pedido de compra em andamento)", H2))
    if len(pend):
        linhas = [[Paragraph(str(ordem.get(r.sku, "")), P), Paragraph(r.sku, PB), Paragraph(r.descricao, P),
                   Paragraph(_n(r.sem_pedido_cx), PB)] for r in pend.itertuples(index=False)]
        el.append(_tabela(["Prioridade", "SKU", "Descrição", "Caixas pendentes"], linhas,
                          [20 * mm, 30 * mm, W - 20 * mm - 30 * mm - 34 * mm, 34 * mm], (3,)))
        if len(faltas):
            lf = [[Paragraph(r.comp_sku, PB), Paragraph(str(r.descricao), P), Paragraph(str(r.unidade), P), Paragraph(_n(r.falta, 1), PB)]
                  for r in faltas.itertuples(index=False)]
            el.append(Spacer(1, 2 * mm))
            el.append(KeepTogether([Paragraph("Material em falta para esses itens", PB), Spacer(1, 1 * mm),
                                    _tabela(["Insumo", "Descrição", "Un.", "Falta"], lf,
                                            [38 * mm, W - 38 * mm - 16 * mm - 28 * mm, 16 * mm, 28 * mm], (3,))]))
    else:
        el.append(Paragraph("Todos os itens do plano têm material disponível ou previsto.", P))

    def rodape(canvas, d):
        canvas.saveState()
        canvas.setStrokeColor(LINHA)
        canvas.line(14 * mm, 10 * mm, landscape(A4)[0] - 14 * mm, 10 * mm)
        canvas.setFont("Helvetica", 7.5)
        canvas.setFillColor(colors.HexColor("#595959"))
        canvas.drawString(14 * mm, 6.5 * mm, "Bioz Green — Pelo Planeta · Compras · compras@biozgreen.eco.br")
        canvas.drawRightString(landscape(A4)[0] - 14 * mm, 6.5 * mm, f"Página {d.page}")
        canvas.restoreState()

    doc.build(el, onFirstPage=rodape, onLaterPages=rodape)
    return buf.getvalue()
