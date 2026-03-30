import truststore
truststore.inject_into_ssl()

import requests
import pandas as pd
from requests.exceptions import RequestException

# SSL da rede corporativa via store do Windows.
BASE_URL = "https://api.cubeanalytics.autostoresystem.com/v1"
token_key = "pW32GPWhY18r8kazjPMO3lhUmRfEMeOhsQ39e7bEBa_taZ78SrgMwzLrTM7NXLmi"

headers = {
    "API-Authorization": f"Token {token_key}"
}


def api_get(url: str) -> requests.Response:
    """Executa GET com tratamento padrao de erro de conexao da API."""
    try:
        response = requests.get(url, headers=headers, timeout=60)
        response.raise_for_status()
        return response
    except RequestException as exc:
        raise RuntimeError("Não foi possível conectar a API") from exc


# ======================
# Coleta de dados da API
# ======================

# Busca a lista de instalações e pega o ID da primeira
response = api_get(f"{BASE_URL}/installations/")
installations = response.json()["results"]
installation_id = installations[0]["id"]
print(f"Installation ID: {installation_id}")

# Solicita ao usuario o ano a ser coletado
year_input = input("Informe o ano desejado (ex: 2026): ").strip()
if not year_input.isdigit() or len(year_input) != 4:
    print("Erro: ano invalido. Informe ano com 4 digitos.")
    exit(1)

target_year = year_input

# Pagina o endpoint diário e coleta TODAS as páginas
url = f"{BASE_URL}/installations/{installation_id}/uptime/"
all_records = []

while url:
    response_daily = api_get(url)
    payload = response_daily.json()
    results = payload.get("results", [])
    all_records.extend(results)  # Adiciona todos os registros da página
    url = payload.get("next")  # Vai para próxima página

# Filtra por ano após coletar tudo
all_records = [r for r in all_records if r.get("date", "").startswith(target_year)]

print(f"Registros diarios encontrados em {target_year}: {len(all_records)}")

# =================
# Transformacao data
# =================
# Consolida os periodos por dia em uma linha
rows = []
for record in all_records:
    date_str = record.get("date")
    if not date_str:
        continue
    
    periods = record.get("result", {}).get("periods", [])

    up_seconds = sum((p.get("up_seconds") or 0) for p in periods)
    down_seconds = sum((p.get("down_seconds") or 0) for p in periods)
    missing_seconds = sum((p.get("missing_seconds") or 0) for p in periods)
    recovery_seconds = sum((p.get("recovery_seconds") or 0) for p in periods)
    response_seconds = sum((p.get("response_seconds") or 0) for p in periods)
    downtime_events = sum(1 for p in periods if p.get("mode") == "downtime")

    total_observed_seconds = up_seconds + down_seconds + missing_seconds
    uptime_ratio = (
        up_seconds / total_observed_seconds if total_observed_seconds > 0 else None
    )

    rows.append({
        "date": date_str,
        "up_seconds": up_seconds,
        "down_seconds": down_seconds,
        "missing_seconds": missing_seconds,
        "recovery_seconds": recovery_seconds,
        "response_seconds": response_seconds,
        "downtime_events": downtime_events,
        "total_observed_seconds": total_observed_seconds,
        "uptime_ratio": round(uptime_ratio, 6) if uptime_ratio is not None else None,
        "uptime_percent": round(uptime_ratio * 100, 4) if uptime_ratio is not None else None,
    })


# =================
# Saida em arquivo
# =================

if not rows:
    print(f"Nenhum dado diario de uptime encontrado para o ano {target_year}.")
else:
    df = pd.DataFrame(rows)
    if len(df) > 0:
        df = df.sort_values("date").reset_index(drop=True)
    output_path = fr"c:\base\API_AUTOSTORE\uptime_{target_year}_diario.xlsx"
    df.to_excel(output_path, index=False, sheet_name=f"UptimeDiario{target_year}")
    print(f"Exportado {len(df)} dias para: {output_path}")
    print(df.to_string())