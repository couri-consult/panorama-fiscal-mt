"""CAPAG loader: reads local annual CSVs from capag/capagdosestados{YEAR}.csv.

The Tesouro publishes CAPAG once per year. CSV layouts shifted across years:

  2018         : UF;Ind1;Ind2;Ind3;Classificação                     (no individual notes; values as %)
  2019         : header line "Dados gerais em DD/MM/YYYY", data is UF;I1;N1;I2;N2;I3;N3;Class (decimal vals)
  2020         : UF;I1;N1;I2;N2;I3;N3;Classificação;Observação        (decimal vals)
  2021         : UF;I1;N1;I2;N2;I3;N3;Classificação                  (% vals)
  2022, 2023   : UF;I1;N1;I2;N2;I3;N3;Classificação;Observação        (% vals)
  2024, 2025   : UF;I1;N1;I2;N2;I3;N3;Classificação;Qualidade;Obs     (% vals)

Indicador 1 = Endividamento (DC/RCL)
Indicador 2 = Poupança Corrente (DespCorr/RecCorrAjust)
Indicador 3 = Liquidez (ObrigFin/DispCx)
"""
import csv
import os

UF_DF = "DF"


def _parse_row_with_notes(row):
    """Parse a row from the modern format (with individual notas)."""
    # row is a list: [UF, I1, N1, I2, N2, I3, N3, Class, ...]
    return {
        "uf": row[0].strip(),
        "endividamento_valor": row[1].strip(),
        "endividamento_nota": row[2].strip(),
        "poupanca_valor": row[3].strip(),
        "poupanca_nota": row[4].strip(),
        "liquidez_valor": row[5].strip(),
        "liquidez_nota": row[6].strip(),
        "consolidado": row[7].strip(),
    }


def _parse_row_simple(row):
    """Parse the 2018 format: only consolidated note, no per-indicator notes."""
    return {
        "uf": row[0].strip(),
        "endividamento_valor": row[1].strip(),
        "endividamento_nota": "",
        "poupanca_valor": row[2].strip(),
        "poupanca_nota": "",
        "liquidez_valor": row[3].strip(),
        "liquidez_nota": "",
        "consolidado": row[4].strip(),
    }


def _read_csv(path):
    """Read a CAPAG CSV, return list of rows, dispatching on column count."""
    with open(path, encoding="utf-8-sig") as f:
        reader = csv.reader(f, delimiter=";")
        rows = [r for r in reader if r and r[0].strip()]

    # 2019 has a header line "Dados gerais em..."; skip lines until first row starts with UF code
    valid_uf = {"AC", "AL", "AM", "AP", "BA", "CE", "DF", "ES", "GO", "MA", "MG", "MS", "MT", "PA", "PB", "PE", "PI", "PR", "RJ", "RN", "RO", "RR", "RS", "SC", "SE", "SP", "TO"}
    data_rows = [r for r in rows if r[0].strip() in valid_uf]

    parsed = []
    for r in data_rows:
        # column 2 is "Nota 1" in modern format (single letter A/B/C/D), or another value in simple format
        looks_like_note = len(r) >= 8 and r[2].strip() in ("A", "B", "C", "D")
        if looks_like_note:
            parsed.append(_parse_row_with_notes(r))
        else:
            # 2018 format
            if len(r) >= 5:
                parsed.append(_parse_row_simple(r))
    return parsed


def load_history(capag_dir, uf=UF_DF):
    """Load CAPAG history for a given UF. Returns dict {year: parsed_row}."""
    history = {}
    for fname in sorted(os.listdir(capag_dir)):
        if not fname.startswith("capagdosestados") or not fname.endswith(".csv"):
            continue
        try:
            year = int(fname.replace("capagdosestados", "").replace(".csv", ""))
        except ValueError:
            continue
        rows = _read_csv(os.path.join(capag_dir, fname))
        for r in rows:
            if r["uf"] == uf:
                history[year] = r
                break
    return history
