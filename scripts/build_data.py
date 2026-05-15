"""Orchestrator: builds data.json by combining live API data with the manual XLSX.

Strategy
--------
The dashboard's index.html consumes a JSON whose top-level keys mirror the previous Excel
sheet names (kpis, investimentos, caixa, capag, capag_historico, pessoal, clp_ranking,
ppps, ppps_projecao). We start from the manual XLSX as a baseline and overlay API-derived
values when the loaders succeed. If an API call fails, the manual value is preserved.

Run:
    python scripts/build_data.py
"""
import json
import os
import re
import sys
import datetime
import traceback

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)


def _load_dotenv(path):
    """Tiny inline .env reader (KEY=value, ignores comments/blanks). Avoids dotenv dep."""
    if not os.path.exists(path):
        return
    with open(path, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            k, v = line.split("=", 1)
            k, v = k.strip(), v.strip().strip('"').strip("'")
            if k and k not in os.environ:
                os.environ[k] = v


_load_dotenv(os.path.join(ROOT, ".env"))

from loaders import siconfi, ibge, capag, manual  # noqa: E402
# Note: transparencia loader não é usado em MT (FCDF é exclusivo do DF).
import transforms  # noqa: E402

UF = "MT"  # cod_ibge=51 (definido em loaders/siconfi.py como MT_ID_ENTE)

MANUAL_XLSX = os.path.join(ROOT, "manual", "panorama_manual.xlsx")
CAPAG_DIR = os.path.join(ROOT, "capag")
OUTPUT = os.path.join(ROOT, "data.json")

# Reference periods for the current dashboard snapshot.
# We use TWO RREO snapshots for different reasons:
#   - RREO_CURRENT (2026/1): latest published, used for "Orçamento de MT em 2026"
#   - RREO_CLOSED  (2025/6): closing of 2025, used for investimento liquidado + receita
#     de impostos realizada (so we compare full-year fiscal values, not partial 2026).
RREO_CURRENT_YEAR = 2026
RREO_CURRENT_BIM = 1
RREO_CLOSED_YEAR = 2025
RREO_CLOSED_BIM = 6
DEFAULT_RGF_YEAR = 2025
DEFAULT_RGF_QUAD = 3
DEFAULT_POP_YEAR = 2024
# IBGE UF code para Mato Grosso (51 cod_ibge é p/ SICONFI; pop IBGE também usa 51)
MT_IBGE_UF = 51


def safely(label, fn, fallback):
    """Try fn(); on exception or empty return, log and return fallback."""
    try:
        result = fn()
    except Exception as e:
        print(f"  [{label}] FAILED: {type(e).__name__}: {e} — using fallback")
        return fallback, False
    if result is None or (hasattr(result, "__len__") and len(result) == 0):
        print(f"  [{label}] empty — using fallback")
        return fallback, False
    return result, True


def load_all_manual_sheets():
    """Read every sheet from manual.xlsx, normalizing the `chave` column.

    Strips parenthetical suffixes from `chave` values (e.g. 'beneficios (R$ 1,00)'
    becomes 'beneficios') so the user can document units inline without breaking
    the chave→label mapping in index.html.
    """
    import openpyxl
    wb = openpyxl.load_workbook(MANUAL_XLSX, data_only=True)
    sheets = {sheet: manual.read_sheet(wb, sheet) for sheet in wb.sheetnames}
    for rows in sheets.values():
        for row in rows:
            ch = row.get("chave")
            if isinstance(ch, str):
                row["chave"] = re.sub(r"\s*\([^)]*\)\s*$", "", ch).strip()
    return sheets


def fmt_bi(reais, decimals=1):
    return f"R$ {reais/1e9:.{decimals}f} bi".replace(".", ",")


def _coerce_to_reais(value):
    """Accept either a numeric value (in reais) or a string like 'R$ 11,3 bi' and
    return a float in reais. Returns None if the input isn't parseable."""
    if value is None:
        return None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)
    s = str(value).strip()
    m = re.search(r"([\d.,]+)\s*bi", s, re.IGNORECASE)
    if m:
        try:
            return float(m.group(1).replace(".", "").replace(",", ".")) * 1e9
        except ValueError:
            return None
    # Try parsing plain BR number (e.g. "13400000000" or "13.400.000.000")
    s_clean = s.replace("R$", "").replace(" ", "").replace(".", "").replace(",", ".")
    try:
        return float(s_clean)
    except ValueError:
        return None


def fmt_mi(reais):
    return f"R$ {reais/1e6:.0f} mi"


def update_kpi(kpis, chave, **fields):
    """Find the kpi row by `chave` and update the given fields in place."""
    for row in kpis:
        if row.get("chave") == chave:
            row.update(fields)
            return True
    return False


def update_pessoal(pessoal_rows, chave, value):
    for row in pessoal_rows:
        if row.get("chave") == chave:
            row["valor"] = value
            return True
    return False


def build():
    print(f"Building data.json (root={ROOT})\n")
    sources = {}

    # ---- 1. Manual baseline ----
    print("[1/4] Loading manual XLSX baseline...")
    sheets = load_all_manual_sheets()
    print(f"  sheets: {list(sheets.keys())}")
    for k in sheets:
        sources[k] = "manual"

    # ---- 2. CAPAG (CSV local + obs/conceito do manual) ----
    print(f"\n[2/4] Loading CAPAG (CSVs) para UF={UF}...")
    capag_history, ok = safely("CAPAG", lambda: capag.load_history(CAPAG_DIR, uf=UF), {})
    if ok:
        # Indicators (notas + valores) vêm do CSV; mas obs e conceito ficam no
        # manual.xlsx para ser editável (mensagens de "Para nota A: < 60%..." e
        # as explicações longas). Faz merge por `indicador`.
        manual_capag = sheets.get("capag", [])
        manual_by_indicador = {str(r.get("indicador") or "").strip(): r for r in manual_capag}
        csv_capag = transforms.build_capag_current(capag_history, RREO_CLOSED_YEAR)
        for row in csv_capag:
            m = manual_by_indicador.get(row.get("indicador", ""))
            if m:
                if m.get("obs"):
                    row["obs"] = m["obs"]
                if m.get("conceito"):
                    row["conceito"] = m["conceito"]
        sheets["capag"] = csv_capag
        sheets["capag_historico"] = transforms.build_capag_history(capag_history)
        sources["capag"] = "csv-local+manual(obs)"
        sources["capag_historico"] = "csv-local"
        print(f"  override CAPAG (years {sorted(capag_history.keys())})")
    capag_nota = capag_history.get(RREO_CLOSED_YEAR, {}).get("consolidado") if capag_history else None

    # ---- 3. IBGE população de MT ----
    print(f"\n[3/4] Loading IBGE população MT {DEFAULT_POP_YEAR}...")
    populacao_mt, ok_pop = safely(
        "IBGE",
        lambda: ibge.fetch_population(DEFAULT_POP_YEAR, uf=MT_IBGE_UF),
        None,
    )
    if ok_pop:
        print(f"  população MT {DEFAULT_POP_YEAR}: {populacao_mt:,}")

    # ---- 4. SICONFI (RREO atual + fechamento) + RGF ----
    print(f"\n[4/4] Loading SICONFI APIs (RREO atual+fechamento, RGF {DEFAULT_RGF_YEAR}/{DEFAULT_RGF_QUAD}º quad)...")

    # RREO atual (orçamento MT 2026)
    rreo_current, ok_rreo_current = safely(
        f"RREO {RREO_CURRENT_YEAR}/{RREO_CURRENT_BIM} Anexo 01",
        lambda: siconfi.fetch_rreo(RREO_CURRENT_YEAR, RREO_CURRENT_BIM, "RREO-Anexo 01"),
        [],
    )
    # RREO fechamento (investimento e receita de impostos do ano completo anterior)
    rreo_balanco, ok_rreo = safely(
        f"RREO {RREO_CLOSED_YEAR}/{RREO_CLOSED_BIM} Anexo 01",
        lambda: siconfi.fetch_rreo(RREO_CLOSED_YEAR, RREO_CLOSED_BIM, "RREO-Anexo 01"),
        [],
    )
    rgf_dtp, ok_rgf_dtp = safely(
        "RGF Anexo 01",
        lambda: siconfi.fetch_rgf(DEFAULT_RGF_YEAR, DEFAULT_RGF_QUAD, "RGF-Anexo 01"),
        [],
    )
    rgf_caixa, ok_rgf_caixa = safely(
        "RGF Anexo 05",
        lambda: siconfi.fetch_rgf(DEFAULT_RGF_YEAR, DEFAULT_RGF_QUAD, "RGF-Anexo 05"),
        [],
    )

    # Extract specific values & overlay into kpis/pessoal
    rreo_current_vals = siconfi.extract_rreo_balanco(rreo_current) if ok_rreo_current else {}
    rreo_vals = siconfi.extract_rreo_balanco(rreo_balanco) if ok_rreo else {}
    rgf_dtp_vals = siconfi.extract_rgf_dtp(rgf_dtp) if ok_rgf_dtp else {}
    rgf_caixa_vals = siconfi.extract_rgf_caixa(rgf_caixa) if ok_rgf_caixa else {}

    print("\nValores extraídos do SICONFI:")
    for label, vals in [("RREO atual", rreo_current_vals), ("RREO fechamento", rreo_vals),
                         ("RGF DTP", rgf_dtp_vals), ("RGF Caixa", rgf_caixa_vals)]:
        for k, v in vals.items():
            if v is not None:
                unit = "" if k.startswith("pct") else " R$" if isinstance(v, (int, float)) and abs(v) > 1e6 else ""
                print(f"  [{label}] {k} = {v}{unit}")

    # Overlay on KPIs and pessoal sheet
    kpis = sheets.get("kpis", [])
    # Orçamento MT: TOTAL DAS DESPESAS (XII) do RREO atual (2026/1).
    # Mantemos o chave 'orcamento_df' por compatibilidade com o front (genérico).
    if rreo_current_vals.get("orcamento_total_dotacao"):
        update_kpi(kpis, "orcamento_df",
                   valor_bilhoes=fmt_bi(rreo_current_vals["orcamento_total_dotacao"]),
                   sub=f"Dotação atualizada {RREO_CURRENT_YEAR} (RREO {RREO_CURRENT_BIM}º bim)")
        sources["kpi.orcamento_df"] = "siconfi-rreo-atual"

    rcl = rgf_dtp_vals.get("rcl")
    # Investimento e receita de impostos vêm do fechamento (2025/6º bim)
    inv = rreo_vals.get("investimento_liquidado")
    if rcl and inv:
        pct = (inv / rcl) * 100
        # Sub: posição do MT no ranking nacional (vem do manual sheet "investimentos")
        invest_rows = sheets.get("investimentos", [])
        mt_row = next((r for r in invest_rows if str(r.get("estado", "")).strip() == UF), None)
        sub_invest = ""
        if mt_row and mt_row.get("ranking"):
            sub_invest = f"{int(mt_row['ranking'])}º lugar no Brasil"
        update_kpi(kpis, "investimento_rcl",
                   valor_bilhoes=f"{pct:.1f}%".replace(".", ","),
                   sub=sub_invest)
        sources["kpi.investimento_rcl"] = "siconfi-rreo+rgf"

    if capag_nota:
        update_kpi(kpis, "capag", valor_bilhoes=capag_nota)
        sources["kpi.capag"] = "csv-local"

    # NOTE: 'caixa' KPI value depends on the right TOTAL row (vinculados + não vinculados)
    # which we still need to confirm. Leaving as manual until then.

    # Pessoal sheet from RGF DTP
    pessoal = sheets.get("pessoal", [])
    if rgf_dtp_vals.get("pct_pessoal_rcl_ajustada") is not None:
        update_pessoal(pessoal, "atual_pct", round(rgf_dtp_vals["pct_pessoal_rcl_ajustada"], 2))
        sources["pessoal.atual_pct"] = "siconfi-rgf"
    if rcl:
        update_pessoal(pessoal, "rcl_bi", round(rcl / 1e9, 2))
        sources["pessoal.rcl_bi"] = "siconfi-rgf"

    # KPI beneficios: total vem do manual (Beneficiômetro é Qlik sem API).
    # Aceita valor_bilhoes como número (em reais) ou string já formatada.
    # O subtítulo "% da receita de impostos" é calculado a partir da
    # receita_impostos_realizada extraída do RREO de fechamento.
    receita_impostos = rreo_vals.get("receita_impostos_realizada")
    if receita_impostos:
        for row in kpis:
            if row.get("chave") != "beneficios":
                continue
            beneficios_reais = _coerce_to_reais(row.get("valor_bilhoes"))
            if beneficios_reais is None:
                break
            pct = beneficios_reais / receita_impostos * 100
            row["sub"] = f"{pct:.0f}% da receita de impostos"
            sources["kpi.beneficios.sub"] = "siconfi-rreo-calc"
            break

    # Normalize KPI valor_bilhoes coming from the manual XLSX to the dashboard's
    # "R$ X,Y bi" display format:
    #   - numbers (int/float in reais) → fmt_bi
    #   - strings already in "R$ X,Y bi" → kept as-is
    #   - strings that look like raw BR-formatted numbers ("13447362101,71") → fmt_bi
    #   - non-monetary strings ("C", "4,7%", "R$ 570 mi") → kept as-is
    for row in kpis:
        v = row.get("valor_bilhoes")
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            row["valor_bilhoes"] = fmt_bi(v)
        elif isinstance(v, str):
            s = v.strip()
            # Skip strings that already look formatted or are non-monetary
            if re.search(r"\bbi\b", s, re.IGNORECASE) or re.search(r"\bmi\b", s, re.IGNORECASE):
                continue
            if "%" in s or len(s) <= 2:
                continue
            # Try to coerce — if parseable to a meaningful reais value, reformat
            n = _coerce_to_reais(s)
            if n is not None and abs(n) >= 1e6:  # ≥ 1 milhão = monetary, not a small number
                row["valor_bilhoes"] = fmt_bi(n)

    # KPI caixa: TOTAL (IV) do RGF atual
    caixa_total = rgf_caixa_vals.get("caixa_liquido_total")
    if caixa_total is not None:
        # "bi" para |caixa| >= 1 bi; "mi" caso contrário (mais legível para valores pequenos)
        if abs(caixa_total) >= 1e9:
            label = fmt_bi(caixa_total)
        else:
            label = f"R$ {caixa_total/1e6:.0f} mi"
        cor = "red" if caixa_total < 1e9 else "amber" if caixa_total < 3e9 else "green"
        update_kpi(kpis, "caixa",
                   valor_bilhoes=label,
                   sub=f"Disponibilidade líquida — dez/{DEFAULT_RGF_YEAR}",
                   cor=cor)
        sources["kpi.caixa"] = "siconfi-rgf-anexo05"

    # Tabela histórica de caixa — busca RGF Anexo 05 para cada ano de 2021 até DEFAULT_RGF_YEAR
    print(f"\nFetching caixa history {2021}..{DEFAULT_RGF_YEAR}")
    caixa_history = safely(
        "caixa history",
        lambda: siconfi.fetch_caixa_history(range(2021, DEFAULT_RGF_YEAR + 1)),
        [],
    )
    if isinstance(caixa_history, tuple):
        caixa_history, _ok = caixa_history
    if caixa_history:
        sheets["caixa"] = caixa_history
        sources["caixa"] = "siconfi-rgf-anexo05"
        print(f"  caixa history: {len(caixa_history)} anos")

    # Posição geral do MT no ranking de competitividade CLP (texto livre vindo
    # do manual.xlsx, aba 'clp_meta' linha 'pos_geral'; cai p/ padrão se ausente).
    clp_meta_rows = sheets.get("clp_meta", [])
    clp_pos_geral = ""
    for r in clp_meta_rows:
        if str(r.get("chave", "")).strip() == "pos_geral":
            clp_pos_geral = str(r.get("valor", "")).strip()
            break

    # ---- Compose final data.json ----
    data = {
        "_meta": {
            "uf": UF,
            "updated_at": datetime.datetime.now().isoformat(timespec="seconds"),
            "rreo_current_period": f"{RREO_CURRENT_YEAR}/{RREO_CURRENT_BIM}",
            "rreo_closed_period": f"{RREO_CLOSED_YEAR}/{RREO_CLOSED_BIM}",
            "rgf_period": f"{DEFAULT_RGF_YEAR}/{DEFAULT_RGF_QUAD}",
            "populacao_year": DEFAULT_POP_YEAR,
            "populacao_mt": populacao_mt,
            "capag_nota_consolidada": capag_nota,
            "clp_pos_geral": clp_pos_geral,
            "raw_siconfi": {
                "rreo_current": rreo_current_vals,
                "rreo_closed": rreo_vals,
                "rgf_dtp": rgf_dtp_vals,
                "rgf_caixa": rgf_caixa_vals,
            },
            "sources": sources,
        },
        "kpis": kpis,
        "investimentos": sheets.get("investimentos", []),
        "caixa": sheets.get("caixa", []),
        "capag": sheets.get("capag", []),
        "capag_historico": sheets.get("capag_historico", []),
        "pessoal": pessoal,
        "clp_ranking": sheets.get("clp_ranking", []),
    }

    with open(OUTPUT, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2, default=str)
    print(f"\nWrote {OUTPUT} ({os.path.getsize(OUTPUT):,} bytes)")
    print(f"Sources used: {json.dumps(sources, ensure_ascii=False, indent=2)}")


if __name__ == "__main__":
    try:
        build()
    except Exception:
        traceback.print_exc()
        sys.exit(1)
