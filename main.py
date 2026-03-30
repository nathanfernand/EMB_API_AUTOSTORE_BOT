import os
import queue
import shutil
import tempfile
import threading
from datetime import datetime

import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.scrolledtext import ScrolledText

import pandas as pd
import requests
import truststore
from requests.exceptions import RequestException


# SSL da rede corporativa via store do Windows.
truststore.inject_into_ssl()

# Configuracao base da API.
BASE_URL = "https://api.cubeanalytics.autostoresystem.com/v1"
TOKEN = "pW32GPWhY18r8kazjPMO3lhUmRfEMeOhsQ39e7bEBa_taZ78SrgMwzLrTM7NXLmi"
HEADERS = {"API-Authorization": f"Token {TOKEN}"}
DEFAULT_YEAR = datetime.now().year
EXPORT_FILE = r"c:\base\API_AUTOSTORE\api_autostore_export_2026.xlsx"


def api_get(url: str, *, headers: dict, timeout: int) -> requests.Response:
	"""Executa GET com tratamento padrao de erro de conexao da API."""
	try:
		response = requests.get(url, headers=headers, timeout=timeout)
		response.raise_for_status()
		return response
	except RequestException as exc:
		raise RuntimeError("Não foi possível conectar a API") from exc


# =========================
# App principal / Navegacao
# =========================
class App(tk.Tk):
	"""Janela principal com roteamento entre as paginas."""
	def __init__(self) -> None:
		super().__init__()
		self.title("AutoStore - Atualizacao de Dados")
		self.geometry("900x550")
		self.minsize(800, 500)
		self.center_window(900, 550)

		self.log_queue: queue.Queue = queue.Queue()
		self.worker_thread: threading.Thread | None = None

		container = tk.Frame(self)
		container.pack(fill="both", expand=True)
		container.grid_rowconfigure(0, weight=1)
		container.grid_columnconfigure(0, weight=1)

		self.frames: dict[type[tk.Frame], tk.Frame] = {}
		for page_cls in (WelcomePage, ActionsPage):
			frame = page_cls(parent=container, controller=self)
			self.frames[page_cls] = frame
			frame.grid(row=0, column=0, sticky="nsew")

		self.show_frame(WelcomePage)
		self.after(100, self.process_log_queue)

	def center_window(self, width: int, height: int) -> None:
		self.update_idletasks()
		screen_w = self.winfo_screenwidth()
		screen_h = self.winfo_screenheight()
		x = (screen_w // 2) - (width // 2)
		y = (screen_h // 2) - (height // 2)
		self.geometry(f"{width}x{height}+{x}+{y}")

	def show_frame(self, page_cls: type[tk.Frame]) -> None:
		frame = self.frames[page_cls]
		frame.tkraise()

	def process_log_queue(self) -> None:
		"""Consome eventos da thread e atualiza a UI no thread principal."""
		try:
			while True:
				event = self.log_queue.get_nowait()
				event_type = event[0]
				actions_page: ActionsPage = self.frames[ActionsPage]  # type: ignore[assignment]

				if event_type == "log":
					actions_page.append_log(event[1])
				elif event_type == "status":
					actions_page.set_status(event[1], event[2])
				elif event_type == "error_dialog":
					messagebox.showerror("Erro", event[1])
				elif event_type == "done":
					actions_page.set_busy(False)
		except queue.Empty:
			pass
		finally:
			self.after(100, self.process_log_queue)


class WelcomePage(tk.Frame):
	"""Tela inicial de boas-vindas."""
	def __init__(self, parent: tk.Widget, controller: App) -> None:
		super().__init__(parent)
		self.controller = controller

		content = tk.Frame(self)
		content.pack(expand=True)

		title = tk.Label(content, text="Bem-vindo! 🤖", font=("Segoe UI", 28, "bold"))
		title.pack(pady=(0, 24))

		next_btn = tk.Button(
			content,
			text="Proxima pagina ➜",
			font=("Segoe UI", 12, "bold"),
			padx=16,
			pady=10,
			command=lambda: controller.show_frame(ActionsPage),
		)
		next_btn.pack()


class ActionsPage(tk.Frame):
	"""Tela de acoes, status e log operacional."""
	def __init__(self, parent: tk.Widget, controller: App) -> None:
		super().__init__(parent)
		self.controller = controller

		self.columnconfigure(0, weight=1)
		self.rowconfigure(3, weight=1)

		title = tk.Label(self, text="Atualizacao de dados", font=("Segoe UI", 22, "bold"))
		title.grid(row=0, column=0, sticky="w", padx=20, pady=(18, 8))

		controls_box = tk.Frame(self)
		controls_box.grid(row=1, column=0, sticky="ew", padx=20, pady=(0, 10))

		update_all_row = tk.Frame(controls_box)
		update_all_row.pack(fill="x", pady=(0, 8))

		self.update_all_btn = tk.Button(
			update_all_row,
			text="⟳  Atualizar Tudo",
			font=("Segoe UI", 12, "bold"),
			bg="#0B5ED7",
			fg="white",
			activebackground="#084298",
			activeforeground="white",
			padx=18,
			pady=6,
			command=self.start_update_all,
		)
		self.update_all_btn.pack(side="left")

		primary_btn_row = tk.Frame(controls_box)
		primary_btn_row.pack(fill="x")

		secondary_btn_row = tk.Frame(controls_box)
		secondary_btn_row.pack(fill="x", pady=(10, 0))

		self.bin_btn = tk.Button(
			primary_btn_row,
			text="Atualizar dados de Bin presentations",
			font=("Segoe UI", 11),
			command=self.start_bin_update,
		)
		self.bin_btn.pack(side="left", padx=(0, 10))

		self.uptime_btn = tk.Button(
			primary_btn_row,
			text="Atualizar dados de Uptime",
			font=("Segoe UI", 11),
			command=self.start_uptime_update,
		)
		self.uptime_btn.pack(side="left")

		self.robot_errors_btn = tk.Button(
			secondary_btn_row,
			text="Robot Errors",
			font=("Segoe UI", 11),
			command=self.start_robot_errors_update,
		)
		self.robot_errors_btn.pack(side="left", padx=(0, 10))

		self.robot_mtbf_btn = tk.Button(
			secondary_btn_row,
			text="Robot MTBF",
			font=("Segoe UI", 11),
			command=self.start_robot_mtbf_update,
		)
		self.robot_mtbf_btn.pack(side="left", padx=(0, 10))

		self.incidents_btn = tk.Button(
			secondary_btn_row,
			text="Incidents",
			font=("Segoe UI", 11),
			command=self.start_incidents_update,
		)
		self.incidents_btn.pack(side="left", padx=(0, 10))

		self.uptime_trend_btn = tk.Button(
			secondary_btn_row,
			text="Uptime Trend",
			font=("Segoe UI", 11),
			command=self.start_uptime_trend_update,
		)
		self.uptime_trend_btn.pack(side="left")

		year_box = tk.Frame(primary_btn_row)
		year_box.pack(side="right")

		year_label = tk.Label(year_box, text="Ano:", font=("Segoe UI", 10, "bold"))
		year_label.pack(side="left", padx=(0, 6))

		self.year_var = tk.StringVar(value=str(DEFAULT_YEAR))
		self.year_spin = tk.Spinbox(
			year_box,
			from_=2000,
			to=2100,
			textvariable=self.year_var,
			width=6,
			font=("Segoe UI", 10),
		)
		self.year_spin.pack(side="left")

		self.status_var = tk.StringVar(value="Status: Aguardando")
		self.status_label = tk.Label(self, textvariable=self.status_var, font=("Segoe UI", 10, "bold"), fg="#555555")
		self.status_label.grid(row=2, column=0, sticky="e", padx=20)

		self.log_box = ScrolledText(self, wrap="word", font=("Consolas", 10), state="disabled")
		self.log_box.grid(row=3, column=0, sticky="nsew", padx=20, pady=(0, 12))

		footer = tk.Frame(self)
		footer.grid(row=4, column=0, sticky="w", padx=20, pady=(0, 18))

		self.back_btn = tk.Button(
			footer,
			text="Voltar",
			font=("Segoe UI", 10),
			command=lambda: controller.show_frame(WelcomePage),
		)
		self.back_btn.pack(side="left")

	def append_log(self, message: str) -> None:
		"""Escreve uma linha no log com timestamp e autoscroll."""
		timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
		self.log_box.configure(state="normal")
		self.log_box.insert("end", f"[{timestamp}] {message}\n")
		self.log_box.see("end")
		self.log_box.configure(state="disabled")

	def set_status(self, text: str, color: str) -> None:
		"""Atualiza texto e cor de status da tela."""
		self.status_var.set(f"Status: {text}")
		self.status_label.configure(fg=color)

	def set_busy(self, busy: bool) -> None:
		"""Liga/desliga controles enquanto uma atualizacao estiver em execucao."""
		state = "disabled" if busy else "normal"
		self.update_all_btn.configure(state=state)
		self.bin_btn.configure(state=state)
		self.uptime_btn.configure(state=state)
		self.robot_errors_btn.configure(state=state)
		self.robot_mtbf_btn.configure(state=state)
		self.incidents_btn.configure(state=state)
		self.uptime_trend_btn.configure(state=state)
		self.back_btn.configure(state=state)
		self.year_spin.configure(state=state)

	def get_selected_year(self) -> int:
		"""Valida o ano informado pelo usuario."""
		raw = self.year_var.get().strip()
		if not raw.isdigit() or len(raw) != 4:
			raise ValueError("Ano invalido. Informe um ano com 4 digitos, por exemplo 2026.")

		year = int(raw)
		if year < 2000 or year > 2100:
			raise ValueError("Ano fora do intervalo permitido (2000 a 2100).")
		return year

	def start_update_all(self) -> None:
		"""Inicia atualizacao de todos os dados de uma vez usando o arquivo fixo."""
		try:
			selected_year = self.get_selected_year()
		except ValueError as exc:
			self.append_log(f"Erro: {exc}")
			messagebox.showerror("Ano invalido", str(exc))
			return

		if not os.path.exists(EXPORT_FILE):
			messagebox.showerror(
				"Arquivo nao encontrado",
				f"O arquivo de exportacao nao foi encontrado:\n{EXPORT_FILE}",
			)
			return

		self.set_busy(True)
		self.controller.log_queue.put(("status", "Executando", "#0B5ED7"))
		self.controller.log_queue.put(("log", f"Ano selecionado: {selected_year}"))
		self.controller.log_queue.put(("log", f"Arquivo: {EXPORT_FILE}"))

		self.controller.worker_thread = threading.Thread(
			target=self.run_update_all,
			args=(selected_year,),
			daemon=True,
		)
		self.controller.worker_thread.start()

	def run_update_all(self, selected_year: int) -> None:
		"""Coleta todos os dados da API e grava todas as abas no Excel de uma vez."""
		try:
			installation_id = get_installation_id(self.controller.log_queue)

			sections = [
				("BinPresentations", "dados de Bin Presentations", fetch_bin_daily_data),
				("Uptime",           "dados de Uptime",            fetch_uptime_daily_data),
				("RobotErrors",      "dados de Robot Errors",      fetch_robot_errors_data),
				("RobotMTBF",        "dados de Robot MTBF",        fetch_robot_mtbf_data),
				("Incidents",        "dados de Incidents",         fetch_incidents_data),
				("UptimeTrend",      "dados de Uptime Trend",      fetch_uptime_trend_data),
			]

			sheets: dict[str, pd.DataFrame] = {}
			for sheet_name, log_label, fetcher in sections:
				self.controller.log_queue.put(("log", f"Coletando {log_label} ({selected_year})..."))
				df = fetcher(installation_id, selected_year, self.controller.log_queue)
				if df.empty:
					self.controller.log_queue.put(("log", f"Aviso: nenhum dado para '{sheet_name}' em {selected_year}."))
				sheets[sheet_name] = df

			update_excel_all_sheets(
				file_path=EXPORT_FILE,
				sheets=sheets,
				log_queue=self.controller.log_queue,
			)

			self.controller.log_queue.put(("status", "Concluido", "#198754"))
			self.controller.log_queue.put(("log", "Processo finalizado com sucesso."))
		except Exception as exc:
			self.controller.log_queue.put(("status", "Erro", "#DC3545"))
			self.controller.log_queue.put(("log", f"Erro: {exc}"))
			self.controller.log_queue.put(("error_dialog", str(exc)))
		finally:
			self.controller.log_queue.put(("done",))

	def start_bin_update(self) -> None:
		self.start_update("bin")

	def start_uptime_update(self) -> None:
		self.start_update("uptime")

	def start_robot_errors_update(self) -> None:
		self.start_update("robot_errors")

	def start_robot_mtbf_update(self) -> None:
		self.start_update("robot_mtbf")

	def start_incidents_update(self) -> None:
		self.start_update("incidents")

	def start_uptime_trend_update(self) -> None:
		self.start_update("uptime_trend")

	def start_update(self, action: str) -> None:
		"""Inicia o fluxo da acao selecionada em thread separada."""
		try:
			selected_year = self.get_selected_year()
		except ValueError as exc:
			self.append_log(f"Erro: {exc}")
			messagebox.showerror("Ano invalido", str(exc))
			return

		excel_path = filedialog.askopenfilename(
			title="Selecione o arquivo Excel",
			filetypes=[("Excel", "*.xlsx")],
		)
		if not excel_path:
			self.append_log("Operacao cancelada: nenhum arquivo Excel selecionado.")
			return

		self.set_busy(True)
		self.controller.log_queue.put(("status", "Executando", "#0B5ED7"))
		self.controller.log_queue.put(("log", f"Ano selecionado: {selected_year}"))
		self.controller.log_queue.put(("log", f"Arquivo selecionado: {excel_path}"))

		self.controller.worker_thread = threading.Thread(
			target=self.run_update,
			args=(action, excel_path, selected_year),
			daemon=True,
		)
		self.controller.worker_thread.start()

	def run_update(self, action: str, excel_path: str, selected_year: int) -> None:
		"""Executa coleta + escrita em Excel sem bloquear a interface."""
		try:
			installation_id = get_installation_id(self.controller.log_queue)

			action_settings = {
				"bin": ("dados diarios de Bin presentations", fetch_bin_daily_data, "BinPresentations"),
				"uptime": ("dados diarios de Uptime", fetch_uptime_daily_data, "Uptime"),
				"robot_errors": ("dados de Robot Errors", fetch_robot_errors_data, "RobotErrors"),
				"robot_mtbf": ("dados de Robot MTBF", fetch_robot_mtbf_data, "RobotMTBF"),
				"incidents": ("dados de Incidents", fetch_incidents_data, "Incidents"),
				"uptime_trend": ("dados de Uptime Trend", fetch_uptime_trend_data, "UptimeTrend"),
			}

			if action not in action_settings:
				raise RuntimeError(f"Acao nao suportada: {action}")

			log_label, fetcher, sheet_name = action_settings[action]
			self.controller.log_queue.put(("log", f"Coletando {log_label} ({selected_year})..."))
			df = fetcher(installation_id, selected_year, self.controller.log_queue)

			if df.empty:
				raise RuntimeError(f"Nenhum dado encontrado para {selected_year}.")

			update_excel_safely(
				file_path=excel_path,
				sheet_name=sheet_name,
				data=df,
				log_queue=self.controller.log_queue,
			)

			self.controller.log_queue.put(("status", "Concluido", "#198754"))
			self.controller.log_queue.put(("log", "Processo finalizado com sucesso."))
		except Exception as exc:
			self.controller.log_queue.put(("status", "Erro", "#DC3545"))
			self.controller.log_queue.put(("log", f"Erro: {exc}"))
			self.controller.log_queue.put(("error_dialog", str(exc)))
		finally:
			self.controller.log_queue.put(("done",))


# ======================
# Coleta de dados da API
# ======================
def get_installation_id(log_queue: queue.Queue) -> str:
	"""Retorna o ID da primeira instalacao disponivel."""
	response = api_get(f"{BASE_URL}/installations/", headers=HEADERS, timeout=30)
	results = response.json().get("results", [])
	if not results:
		raise RuntimeError("Nenhuma instalacao encontrada.")
	installation_id = results[0]["id"]
	log_queue.put(("log", f"Installation ID: {installation_id}"))
	return installation_id


def fetch_paginated_results(url: str, timeout: int = 60) -> list[dict]:
	"""Coleta todos os resultados paginados de um endpoint da API."""
	records: list[dict] = []

	while url:
		response = api_get(url, headers=HEADERS, timeout=timeout)
		payload = response.json()
		records.extend(payload.get("results", []))
		url = payload.get("next")

	return records


def filter_records_by_year(records: list[dict], year: int) -> list[dict]:
	"""Filtra registros pelo ano usando o campo date do payload."""
	year_str = str(year)
	return [record for record in records if record.get("date", "").startswith(year_str)]


def fetch_bin_daily_data(installation_id: str, year: int, log_queue: queue.Queue) -> pd.DataFrame:
	"""Coleta e consolida Bin Presentations diario para o ano informado."""
	url = f"{BASE_URL}/installations/{installation_id}/bin-presentations/"
	records = filter_records_by_year(fetch_paginated_results(url), year)

	rows = []
	for record in records:
		date_value = record.get("date")
		if not date_value:
			continue
		
		ports = record.get("result", {}).get("bin_presentations", [])

		def wavg(field: str) -> float | None:
			# Media ponderada pelo volume de operacoes (count).
			total_weight = sum((p.get("count") or 0) for p in ports)
			if total_weight == 0:
				return None
			weighted = sum((p.get(field) or 0) * (p.get("count") or 0) for p in ports)
			return weighted / total_weight

		rows.append(
			{
				"date": date_value,
				"total_count": sum((p.get("count") or 0) for p in ports),
				"total_picks": sum((p.get("picks") or 0) for p in ports),
				"total_goods_in": sum((p.get("goods_in") or 0) for p in ports),
				"total_count_all_bins": sum((p.get("count_all_bins") or 0) for p in ports),
				"total_inspection_adhoc": sum((p.get("inspection_or_adhoc") or 0) for p in ports),
				"avg_wait_bin_seg": round(wavg("average_wait_bin"), 2) if wavg("average_wait_bin") is not None else None,
				"avg_wait_user_seg": round(wavg("average_wait_user"), 2) if wavg("average_wait_user") is not None else None,
				"avg_waste_time_seg": round(wavg("average_waste_time"), 2) if wavg("average_waste_time") is not None else None,
				"avg_wait_bin_port_seg": round(wavg("average_wait_bin_port"), 2) if wavg("average_wait_bin_port") is not None else None,
				"avg_wait_bin_robot_seg": round(wavg("average_wait_bin_robot"), 2) if wavg("average_wait_bin_robot") is not None else None,
			}
		)

	log_queue.put(("log", f"Dias de Bin presentations coletados: {len(rows)}"))
	if rows:
		return pd.DataFrame(rows).sort_values("date").reset_index(drop=True)
	else:
		return pd.DataFrame()


def fetch_uptime_daily_data(installation_id: str, year: int, log_queue: queue.Queue) -> pd.DataFrame:
	"""Coleta e consolida Uptime diario para o ano informado."""
	url = f"{BASE_URL}/installations/{installation_id}/uptime/"
	records = filter_records_by_year(fetch_paginated_results(url), year)

	rows = []
	for record in records:
		date_value = record.get("date")
		if not date_value:
			continue
		
		periods = record.get("result", {}).get("periods", [])
		up_seconds = sum((p.get("up_seconds") or 0) for p in periods)
		down_seconds = sum((p.get("down_seconds") or 0) for p in periods)
		missing_seconds = sum((p.get("missing_seconds") or 0) for p in periods)
		recovery_seconds = sum((p.get("recovery_seconds") or 0) for p in periods)
		response_seconds = sum((p.get("response_seconds") or 0) for p in periods)
		total_observed_seconds = up_seconds + down_seconds + missing_seconds
		ratio = (up_seconds / total_observed_seconds) if total_observed_seconds > 0 else None

		rows.append(
			{
				"date": date_value,
				"up_seconds": up_seconds,
				"down_seconds": down_seconds,
				"missing_seconds": missing_seconds,
				"recovery_seconds": recovery_seconds,
				"response_seconds": response_seconds,
				"downtime_events": sum(1 for p in periods if p.get("mode") == "downtime"),
				"total_observed_seconds": total_observed_seconds,
				"uptime_ratio": round(ratio, 6) if ratio is not None else None,
				"uptime_percent": round(ratio * 100, 4) if ratio is not None else None,
			}
		)

	log_queue.put(("log", f"Dias de Uptime coletados: {len(rows)}"))
	if rows:
		return pd.DataFrame(rows).sort_values("date").reset_index(drop=True)
	else:
		return pd.DataFrame()


def fetch_robot_errors_data(installation_id: str, year: int, log_queue: queue.Queue) -> pd.DataFrame:
	"""Coleta Robot Errors e expande cada erro em uma linha para o ano informado."""
	url = f"{BASE_URL}/installations/{installation_id}/robot-errors/"
	records = filter_records_by_year(fetch_paginated_results(url), year)

	rows = []
	for record in records:
		date_value = record.get("date")
		if not date_value:
			continue

		for robot_error in record.get("result", {}).get("robot_errors", []):
			row = {
				"date": date_value,
				"version": record.get("version"),
			}
			row.update(robot_error)
			rows.append(row)

	log_queue.put(("log", f"Ocorrencias de Robot Errors coletadas: {len(rows)}"))
	if not rows:
		return pd.DataFrame()

	df = pd.DataFrame(rows)
	sort_columns = [column for column in ("date", "local_installation_timestamp", "robot_id") if column in df.columns]
	if sort_columns:
		df = df.sort_values(sort_columns).reset_index(drop=True)
	return df


def fetch_robot_mtbf_data(installation_id: str, year: int, log_queue: queue.Queue) -> pd.DataFrame:
	"""Coleta Robot MTBF e expande o detalhe de tempo ativo por robo no ano informado."""
	url = f"{BASE_URL}/installations/{installation_id}/robot-mtbf/"
	records = filter_records_by_year(fetch_paginated_results(url), year)

	rows = []
	for record in records:
		date_value = record.get("date")
		if not date_value:
			continue

		result = record.get("result", {})
		summary = result.get("robot_mtbf", {})
		robot_active_times = result.get("robot_active_times", [])
		base_row = {
			"date": date_value,
			"version": record.get("version"),
			"robot_mtbf": summary.get("robot_mtbf"),
			"total_errors": summary.get("total_errors"),
			"total_time_active_s": summary.get("total_time_active_s"),
			"robot_mtbf_system_stops": summary.get("robot_mtbf_system_stops"),
			"total_errors_system_stops": summary.get("total_errors_system_stops"),
			"robot_active_times_count": len(robot_active_times),
		}

		if not robot_active_times:
			rows.append(base_row)
			continue

		for robot_active_time in robot_active_times:
			rows.append(
				{
					**base_row,
					"robot_id": robot_active_time.get("robot_id"),
					"robot_time_active_s": robot_active_time.get("total_time_active_s"),
				}
			)

	log_queue.put(("log", f"Registros de Robot MTBF coletados: {len(rows)}"))
	if not rows:
		return pd.DataFrame()

	df = pd.DataFrame(rows)
	sort_columns = [column for column in ("date", "robot_id") if column in df.columns]
	if sort_columns:
		df = df.sort_values(sort_columns).reset_index(drop=True)
	return df


def fetch_incidents_data(installation_id: str, year: int, log_queue: queue.Queue) -> pd.DataFrame:
	"""Coleta Incidents e expande cada incidente em uma linha para o ano informado."""
	url = f"{BASE_URL}/installations/{installation_id}/incidents/"
	records = filter_records_by_year(fetch_paginated_results(url), year)

	rows = []
	for record in records:
		date_value = record.get("date")
		if not date_value:
			continue

		for incident in record.get("result", {}).get("incidents", []):
			details_display_name = incident.get("details_display_name") or []
			row = {
				"date": date_value,
				"version": record.get("version"),
				**incident,
				"details_display_name": " | ".join(str(item) for item in details_display_name),
			}
			rows.append(row)

	log_queue.put(("log", f"Incidents coletados: {len(rows)}"))
	if not rows:
		return pd.DataFrame()

	df = pd.DataFrame(rows)
	sort_columns = [column for column in ("date", "start_local_timestamp", "incident_id") if column in df.columns]
	if sort_columns:
		df = df.sort_values(sort_columns).reset_index(drop=True)
	return df


def fetch_uptime_trend_data(installation_id: str, year: int, log_queue: queue.Queue) -> pd.DataFrame:
	"""Coleta Uptime Trend e retorna uma linha por data para o ano informado."""
	url = f"{BASE_URL}/installations/{installation_id}/uptime/trend/"
	records = filter_records_by_year(fetch_paginated_results(url), year)

	rows = []
	for record in records:
		date_value = record.get("date")
		if not date_value:
			continue

		trend = record.get("result", {}).get("uptime_trend", {})
		rows.append(
			{
				"date": date_value,
				"version": record.get("version"),
				"week": trend.get("week"),
				"year": trend.get("year"),
				"trend": trend.get("trend"),
				"outlier": trend.get("outlier"),
				"lower_bound": trend.get("lower_bound"),
				"upper_bound": trend.get("upper_bound"),
				"weekly_uptime": trend.get("weekly_uptime"),
				"long_term_trend": trend.get("long_term_trend"),
				"short_term_trend": trend.get("short_term_trend"),
			}
		)

	log_queue.put(("log", f"Registros de Uptime Trend coletados: {len(rows)}"))
	if not rows:
		return pd.DataFrame()

	return pd.DataFrame(rows).sort_values(["date", "week"]).reset_index(drop=True)


# ==========================
# Escrita segura no arquivo
# ==========================
def update_excel_safely(file_path: str, sheet_name: str, data: pd.DataFrame, log_queue: queue.Queue) -> None:
	"""Atualiza/cria aba no Excel com escrita atomica via arquivo temporario."""
	if not os.path.exists(file_path):
		raise RuntimeError("Arquivo Excel selecionado nao existe.")

	temp_fd, temp_path = tempfile.mkstemp(suffix=".xlsx", dir=os.path.dirname(file_path))
	os.close(temp_fd)

	try:
		existing_sheets: dict[str, pd.DataFrame] = {}
		try:
			workbook = pd.ExcelFile(file_path)
			for name in workbook.sheet_names:
				if name != sheet_name:
					existing_sheets[name] = pd.read_excel(file_path, sheet_name=name)
		except Exception:
			existing_sheets = {}

		stamp = f"Atualizado em: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}"

		with pd.ExcelWriter(temp_path, engine="openpyxl") as writer:
			for name, df_old in existing_sheets.items():
				df_old.to_excel(writer, sheet_name=name, index=False)

			data.to_excel(writer, sheet_name=sheet_name, index=False, startrow=1)

			ws = writer.book[sheet_name]
			ws["A1"] = stamp

		shutil.move(temp_path, file_path)
		log_queue.put(("log", f"Aba '{sheet_name}' criada/atualizada no Excel."))
		log_queue.put(("log", stamp))
	except Exception:
		if os.path.exists(temp_path):
			os.remove(temp_path)
		raise


def update_excel_all_sheets(file_path: str, sheets: dict[str, pd.DataFrame], log_queue: queue.Queue) -> None:
	"""Grava todas as abas de uma vez no Excel com escrita atomica via arquivo temporario."""
	if not os.path.exists(file_path):
		raise RuntimeError(f"Arquivo Excel nao encontrado: {file_path}")

	temp_fd, temp_path = tempfile.mkstemp(suffix=".xlsx", dir=os.path.dirname(file_path))
	os.close(temp_fd)

	try:
		# Preservar abas existentes que nao serao substituidas.
		existing_sheets: dict[str, pd.DataFrame] = {}
		try:
			workbook = pd.ExcelFile(file_path)
			for name in workbook.sheet_names:
				if name not in sheets:
					existing_sheets[name] = pd.read_excel(file_path, sheet_name=name)
		except Exception:
			existing_sheets = {}

		stamp = f"Atualizado em: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}"

		with pd.ExcelWriter(temp_path, engine="openpyxl") as writer:
			for name, df_old in existing_sheets.items():
				df_old.to_excel(writer, sheet_name=name, index=False)

			for sheet_name, data in sheets.items():
				if data.empty:
					continue
				data.to_excel(writer, sheet_name=sheet_name, index=False, startrow=1)
				ws = writer.book[sheet_name]
				ws["A1"] = stamp

		shutil.move(temp_path, file_path)
		updated = [name for name, df in sheets.items() if not df.empty]
		log_queue.put(("log", f"Abas atualizadas: {', '.join(updated)}"))
		log_queue.put(("log", stamp))
	except Exception:
		if os.path.exists(temp_path):
			os.remove(temp_path)
		raise


# Ponto de entrada da aplicacao.
if __name__ == "__main__":
	app = App()
	app.mainloop()
