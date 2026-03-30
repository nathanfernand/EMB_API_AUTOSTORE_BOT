import queue
import sys

from main import fetch_uptime_trend_data, get_installation_id


def read_year() -> int:
	year_input = input("Informe o ano desejado (ex: 2026): ").strip()
	if not year_input.isdigit() or len(year_input) != 4:
		raise ValueError("Erro: ano invalido. Informe ano com 4 digitos.")
	return int(year_input)


def main() -> int:
	log_queue: queue.Queue = queue.Queue()
	installation_id = get_installation_id(log_queue)
	print(f"Installation ID: {installation_id}")

	try:
		target_year = read_year()
	except ValueError as exc:
		print(str(exc))
		return 1

	df = fetch_uptime_trend_data(installation_id, target_year, log_queue)
	if df.empty:
		print(f"Nenhum dado encontrado para o ano {target_year}.")
		return 0

	output_path = fr"c:\base\API_AUTOSTORE\uptime_trend_{target_year}_diario.xlsx"
	df.to_excel(output_path, index=False, sheet_name="UptimeTrend")
	print(f"Exportado {len(df)} registros para: {output_path}")
	print(df.to_string())
	return 0


if __name__ == "__main__":
	sys.exit(main())