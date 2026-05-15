"""Pure functions that compute derived values for the dashboard.

These take loader output (dicts/lists) and produce the final shape consumed by index.html.
"""

# Map CAPAG note -> badge color (mirrors the dashboard's color scheme)
NOTE_BADGE = {"A": "green", "B": "amber", "C": "red", "D": "red", "": ""}


def fmt_bi(value_reais):
    """Format reais as 'R$ X,Y bi' string. None-safe."""
    if value_reais is None:
        return ""
    bi = value_reais / 1e9
    return f"R$ {bi:.1f} bi".replace(".", ",")


def fmt_pct(decimal_or_pct, decimals=1):
    """Format a number as percent. Accepts either 0.42 or 42 (heuristic on magnitude)."""
    if decimal_or_pct is None:
        return ""
    v = decimal_or_pct * 100 if abs(decimal_or_pct) <= 2 else decimal_or_pct
    return f"{v:.{decimals}f}%".replace(".", ",")


def build_capag_history(capag_history_dict):
    """Convert the CAPAG history dict {year: row} into the dashboard's expected shape.

    Output: list of dicts with keys ano, endividamento, poupanca, liquidez, consolidado.
    """
    rows = []
    for year in sorted(capag_history_dict.keys()):
        r = capag_history_dict[year]
        rows.append({
            "ano": str(year),
            "endividamento": r["endividamento_nota"] or "ND",
            "poupanca": r["poupanca_nota"] or "ND",
            "liquidez": r["liquidez_nota"] or "ND",
            "consolidado": r["consolidado"] or "ND",
        })
    return rows


def build_capag_current(capag_history_dict, year):
    """Build the current CAPAG indicator table for the given year (the 4 rows shown in the dashboard)."""
    if year not in capag_history_dict:
        return []
    r = capag_history_dict[year]
    return [
        {
            "indicador": "Endividamento (% da RCL)",
            "nota": r["endividamento_nota"],
            "badge": NOTE_BADGE.get(r["endividamento_nota"], ""),
            "valor": r["endividamento_valor"],
            "obs": "",
            "conceito": "Mede a dívida consolidada do estado em relação à receita corrente líquida. Quanto menor, melhor. Abaixo de 60% da RCL = nota A.",
        },
        {
            "indicador": "Poupança Corrente (Despesa/Receita)",
            "nota": r["poupanca_nota"],
            "badge": NOTE_BADGE.get(r["poupanca_nota"], ""),
            "valor": r["poupanca_valor"],
            "obs": "",
            "conceito": "Mede quanto das receitas correntes é consumido pelas despesas correntes. Acima de 95% significa que o estado gasta quase tudo que arrecada só para se manter, sem sobrar para investir.",
        },
        {
            "indicador": "Liquidez Relativa (Dívida/RCL)",
            "nota": r["liquidez_nota"],
            "badge": NOTE_BADGE.get(r["liquidez_nota"], ""),
            "valor": r["liquidez_valor"],
            "obs": "",
            "conceito": "Mede se o estado tem recursos disponíveis em caixa para honrar suas obrigações de curto prazo. Entre 0% e 5% = nota B.",
        },
    ]


def build_kpis(orcamento_df_bi, fcdf_bi, investimento_rcl_pct, capag_nota,
               caixa_mi, beneficios_bi, beneficios_pct_impostos):
    """Assemble the top KPI grid the dashboard expects (chave/valor_bilhoes/sub/cor)."""
    kpis = [
        {"chave": "orcamento_df", "valor_bilhoes": f"R$ {orcamento_df_bi:.1f} bi".replace(".", ","), "sub": "", "cor": "blue"},
        {"chave": "fcdf", "valor_bilhoes": f"R$ {fcdf_bi:.1f} bi".replace(".", ","), "sub": "Lei orçamentária da União", "cor": "blue"},
        {"chave": "investimento_rcl", "valor_bilhoes": f"{investimento_rcl_pct:.1f}%".replace(".", ","), "sub": "% da RCL", "cor": "amber"},
        {"chave": "capag", "valor_bilhoes": capag_nota, "sub": "Nota consolidada", "cor": "red" if capag_nota in ("C", "D") else "amber" if capag_nota == "B" else "green"},
        {"chave": "caixa", "valor_bilhoes": f"R$ {caixa_mi:.0f} mi", "sub": "Disponibilidade líquida (dez)", "cor": "red" if caixa_mi < 1000 else "amber"},
        {"chave": "beneficios", "valor_bilhoes": f"R$ {beneficios_bi:.1f} bi".replace(".", ","), "sub": f"{beneficios_pct_impostos:.0f}% da receita de impostos", "cor": "amber"},
    ]
    return kpis
