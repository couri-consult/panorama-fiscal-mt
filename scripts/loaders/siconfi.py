"""SICONFI loader: RREO and RGF for the State of Mato Grosso (Governo de MT, id_ente=51).

Docs: http://apidatalake.tesouro.gov.br/docs/siconfi/

Each row has columns: exercicio, demonstrativo, periodo, periodicidade, instituicao,
cod_ibge, uf, populacao, anexo, esfera, rotulo, coluna, cod_conta, conta, valor.

We filter rows by `cod_conta` (machine-stable account code) + `coluna` (column header) +
sometimes `conta` (when same cod_conta groups multiple lines, like RGF Anexo 05 where
each cod_conta has rows for different "TOTAL DOS RECURSOS..." subtotals).

For MT, id_ente=51 is the cod_ibge da UF (estadual). Cod_siconfi do Executivo é "51EX"
(Cod_instituicoes_siconfi.pdf disponível em siconfi.tesouro.gov.br).
"""
import time
from ._http import get_json

BASE = "https://apidatalake.tesouro.gov.br/ords/siconfi/tt"
# MT as a state-level entity uses cod_ibge=51 (UF code).
MT_ID_ENTE = 51
# Mantido por compatibilidade — qualquer código que importe DF_ID_ENTE recebe o id de MT.
DF_ID_ENTE = MT_ID_ENTE

THROTTLE_SLEEP = 3  # SICONFI's ORDS pool rate-limits aggressively; pause between calls.


def _normalize(rows):
    return [{k: (v.strip() if isinstance(v, str) else v) for k, v in r.items()} for r in rows]


def fetch_rreo(year, periodo, anexo):
    """Fetch a RREO annex for MT at a given bimester."""
    print(f"[SICONFI] RREO {year}/{periodo}º bim — {anexo}")
    data = get_json(
        f"{BASE}/rreo",
        params={
            "an_exercicio": year,
            "nr_periodo": periodo,
            "co_tipo_demonstrativo": "RREO",
            "id_ente": DF_ID_ENTE,
            "no_anexo": anexo,
        },
    )
    items = _normalize(data.get("items", []))
    print(f"  {len(items)} rows")
    time.sleep(THROTTLE_SLEEP)
    return items


def fetch_rgf(year, quad, anexo, poder="E"):
    """Fetch a RGF annex for MT at a given quadrimester (Q periodicity)."""
    print(f"[SICONFI] RGF {year}/{quad}º quad — {anexo} (poder={poder})")
    data = get_json(
        f"{BASE}/rgf",
        params={
            "an_exercicio": year,
            "in_periodicidade": "Q",
            "nr_periodo": quad,
            "co_tipo_demonstrativo": "RGF",
            "co_poder": poder,
            "id_ente": DF_ID_ENTE,
            "no_anexo": anexo,
        },
    )
    items = _normalize(data.get("items", []))
    print(f"  {len(items)} rows")
    time.sleep(THROTTLE_SLEEP)
    return items


def find_value(rows, *, cod_conta, coluna, conta=None):
    """Extract a single numeric value from a SICONFI dataset, filtering by cod_conta + coluna
    (and optionally `conta` to disambiguate when one cod_conta has multiple sub-rows)."""
    for r in rows:
        if r.get("cod_conta") != cod_conta:
            continue
        if r.get("coluna") != coluna:
            continue
        if conta is not None and r.get("conta") != conta:
            continue
        return r.get("valor")
    return None


# ── RREO Anexo 01 — Balanço Orçamentário ─────────────────────────
def extract_rreo_balanco(rows):
    """Extract budget figures from RREO Anexo 01 rows.

    Returns dict with:
      - orcamento_total_dotacao: linha TOTAL DAS DESPESAS (XII) = (X + XI), dotação atualizada
      - investimento_liquidado: linha INVESTIMENTOS, coluna despesas liquidadas até o bimestre
      - receita_impostos_realizada: linha Impostos, coluna receitas realizadas até o bimestre
    All values are in BRL.

    Note: SICONFI's RREO uses current MCASP layout. The line "(VIII)" in older RREO is
    now "DespesasExcetoIntraOrcamentarias"; the total *including intras* is `TotalDespesas`
    (XII = X + XI, where X = SUBTOTAL = VIII + IX and XI = amortização da dívida/refinanciamento).
    """
    return {
        "orcamento_total_dotacao": find_value(
            rows, cod_conta="TotalDespesas", coluna="DOTAÇÃO ATUALIZADA (e)"
        ),
        "investimento_liquidado": find_value(
            rows, cod_conta="Investimentos", coluna="DESPESAS LIQUIDADAS ATÉ O BIMESTRE (h)"
        ),
        "receita_impostos_realizada": find_value(
            rows, cod_conta="Impostos", coluna="Até o Bimestre (c)"
        ),
    }


# ── RGF Anexo 01 — DTP / Apuração do Limite Legal ────────────────
def extract_rgf_dtp(rows):
    """Extract personnel-expense indicators from RGF Anexo 01.

    Returns: rcl, despesa_pessoal_total, pct_pessoal_rcl_ajustada.
    """
    return {
        "rcl": find_value(rows, cod_conta="ReceitaCorrenteLiquidaLimiteLegal", coluna="Valor"),
        "rcl_ajustada": find_value(rows, cod_conta="ReceitaCorrenteLiquidaAjustada", coluna="Valor"),
        "despesa_pessoal_total": find_value(
            rows, cod_conta="DespesaComPessoalTotal", coluna="TOTAL (ÚLTIMOS 12 MESES) (a)"
        ),
        "pct_pessoal_rcl_ajustada": find_value(
            rows, cod_conta="DespesaComPessoalTotal", coluna="% sobre a RCL Ajustada"
        ),
    }


# ── RGF Anexo 05 — Disponibilidade de Caixa ──────────────────────
# Coluna name is enormous and SICONFI stores it verbatim.
CAIXA_COLUNA_LIQUIDA_APOS_RP = (
    "DISPONIBILIDADE DE CAIXA LÍQUIDA (APÓS A INSCRIÇÃO EM RESTOS A PAGAR NÃO "
    "PROCESSADOS DO EXERCÍCIO) (i) = (g - h)"
)
# cod_conta in SICONFI is "DisponibilidadeDeCaixaLiquidaAposRP" (note AposRP suffix).
CAIXA_COD_CONTA = "DisponibilidadeDeCaixaLiquidaAposRP"


def extract_rgf_caixa(rows):
    """Extract cash-availability figures from RGF Anexo 05.

    All values come from the column "DISPONIBILIDADE DE CAIXA LÍQUIDA (APÓS RPNP)".
    - caixa_liquido_total: linha TOTAL — saldo líquido total
        The label of this row changed across years:
          - 2021, 2022: 'TOTAL (III) = (I + II)'   (before RPPS column was split out)
          - 2023+     : 'TOTAL (IV) = (I + II + III)'
    - caixa_liquido_nao_vinculado: linha TOTAL DOS RECURSOS NÃO VINCULADOS (I) — só os
      recursos discricionários (negativo = passivo descoberto)
    """
    total = find_value(
        rows, cod_conta=CAIXA_COD_CONTA, coluna=CAIXA_COLUNA_LIQUIDA_APOS_RP,
        conta="TOTAL (IV) = (I + II + III)",
    )
    if total is None:  # fallback to legacy layout (2021-2022)
        total = find_value(
            rows, cod_conta=CAIXA_COD_CONTA, coluna=CAIXA_COLUNA_LIQUIDA_APOS_RP,
            conta="TOTAL (III) = (I + II)",
        )
    return {
        "caixa_liquido_total": total,
        "caixa_liquido_nao_vinculado": find_value(
            rows, cod_conta=CAIXA_COD_CONTA, coluna=CAIXA_COLUNA_LIQUIDA_APOS_RP,
            conta="TOTAL DOS RECURSOS NÃO VINCULADOS (I)",
        ),
    }


def fetch_caixa_history(years):
    """Fetch RGF Anexo 05 for each year in `years` (using 3º quadrimestre = year close).

    Returns list of dicts in the shape the dashboard's `caixa` sheet expects:
        [{"ano": "2021", "caixa_liquido": 1920209349, "nao_vinculado": 916943191}, ...]
    Years that fail or have no data are skipped (with a warning).
    """
    history = []
    for year in years:
        try:
            rows = fetch_rgf(year, 3, "RGF-Anexo 05")
        except Exception as e:
            print(f"  [caixa history] {year} skipped: {type(e).__name__}: {e}")
            continue
        if not rows:
            print(f"  [caixa history] {year} empty")
            continue
        v = extract_rgf_caixa(rows)
        if v.get("caixa_liquido_total") is None:
            print(f"  [caixa history] {year} extraction returned None")
            continue
        history.append({
            "ano": str(year),
            "caixa_liquido": v["caixa_liquido_total"],
            "nao_vinculado": v["caixa_liquido_nao_vinculado"],
        })
    return history
