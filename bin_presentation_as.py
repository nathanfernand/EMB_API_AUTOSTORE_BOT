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
url = f"{BASE_URL}/installations/{installation_id}/bin-presentations/"
all_records = []

while url:
    resp = api_get(url)
    data = resp.json()
    results = data.get("results", [])
    all_records.extend(results)  # Adiciona todos os registros da página
    url = data.get("next")  # Vai para próxima página

# Filtra por ano após coletar tudo
all_records = [r for r in all_records if r.get("date", "").startswith(target_year)]

print(f"Registros encontrados em {target_year}: {len(all_records)}")

# =================
# Transformacao data
# =================
# Agrega os dados de todas as portas em uma linha por dia
rows = []
for record in all_records:
    if record.get("version") != "2.2.0":
        continue

    date_str = record.get("date")
    if not date_str:
        continue
    
    ports = record.get("result", {}).get("bin_presentations", [])

    total_count        = sum(p.get("count", 0)               for p in ports)
    total_picks        = sum(p.get("picks", 0)               for p in ports)
    total_goods_in     = sum(p.get("goods_in", 0)            for p in ports)
    total_all_bins     = sum(p.get("count_all_bins", 0)      for p in ports)
    total_insp         = sum(p.get("inspection_or_adhoc", 0) for p in ports)

    # Media ponderada pelo volume de operacoes (count).
    def wavg(field):
        total_w = sum(p.get("count", 0) for p in ports if p.get("count", 0) > 0)
        if total_w == 0:
            return None
        return sum(p.get(field, 0) * p.get("count", 0) for p in ports if p.get("count", 0) > 0) / total_w

    rows.append({
        "date":                   date_str,
        "total_count":            total_count,
        "total_picks":            total_picks,
        "total_goods_in":         total_goods_in,
        "total_count_all_bins":   total_all_bins,
        "total_inspection_adhoc": total_insp,
        "avg_wait_bin_seg":       round(wavg("average_wait_bin"),       2) if wavg("average_wait_bin")       else None,
        "avg_wait_user_seg":      round(wavg("average_wait_user"),      2) if wavg("average_wait_user")      else None,
        "avg_waste_time_seg":     round(wavg("average_waste_time"),     2) if wavg("average_waste_time")     else None,
        "avg_wait_bin_port_seg":  round(wavg("average_wait_bin_port"),  2) if wavg("average_wait_bin_port")  else None,
        "avg_wait_bin_robot_seg": round(wavg("average_wait_bin_robot"), 2) if wavg("average_wait_bin_robot") else None,
    })


# =================
# Saida em arquivo
# =================

if not rows:
    print(f"Nenhum dado encontrado para o ano {target_year}.")
else:
    df = pd.DataFrame(rows)
    if len(df) > 0:
        df = df.sort_values("date").reset_index(drop=True)
    output_path = fr"c:\base\API_AUTOSTORE\bin_presentations_{target_year}_diario.xlsx"
    df.to_excel(output_path, index=False, sheet_name=f"BinPresentations{target_year}")
    print(f"Exportado {len(df)} dias para: {output_path}")
    print(df.to_string())

