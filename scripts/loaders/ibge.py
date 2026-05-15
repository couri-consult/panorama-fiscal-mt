"""IBGE loader: estimated population by UF.

Uses the SIDRA v3 API: https://servicodados.ibge.gov.br/api/v3/agregados/6579 (Estimativa de população).
Variable 9324 = "População residente estimada".
"""
from ._http import get_json

# UF code 53 = Distrito Federal in IBGE
UF_DF = 53


def fetch_population(year, uf=UF_DF):
    """Return estimated population (int) for a UF in the given year."""
    url = f"https://servicodados.ibge.gov.br/api/v3/agregados/6579/periodos/{year}/variaveis/9324"
    data = get_json(url, params={"localidades": f"N3[{uf}]"})
    try:
        serie = data[0]["resultados"][0]["series"][0]["serie"]
        return int(serie[str(year)])
    except (KeyError, IndexError, ValueError) as e:
        raise RuntimeError(f"IBGE response shape unexpected for UF={uf} year={year}: {e}")
