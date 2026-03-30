# API_AUTOSTORE

Aplicacao Tkinter para consultar a API Cube Analytics do AutoStore e exportar os dados para abas de um arquivo Excel existente.

## Raio-x rapido

- Chamadas HTTP: centralizadas em `main.py` via `api_get`, `BASE_URL` e `HEADERS`.
- Autenticacao: header `API-Authorization: Token ...` aplicado em todas as rotas.
- SSL corporativo: `truststore.inject_into_ssl()` usa o store do Windows.
- Installation ID: obtido por `get_installation_id()` a partir de `GET /v1/installations/` e usa a primeira instalacao retornada.
- UI Tkinter: `App`, `WelcomePage` e `ActionsPage` em `main.py`.
- Execucao sem travar a UI: `threading.Thread` para o worker + `queue.Queue` + `after(100, process_log_queue)`.
- Persistencia de resultado: exportacao para abas do Excel por `update_excel_safely()` com escrita atomica em arquivo temporario.

## Opcoes disponiveis na UI

- Atualizar dados de Bin presentations
- Atualizar dados de Uptime
- Robot Errors
- Robot MTBF
- Incidents
- Uptime Trend

Todas as opcoes:

- usam o mesmo `installation_id` obtido automaticamente da API
- usam o ano informado na UI para filtrar o campo `date`
- executam em thread separada
- escrevem em uma aba dedicada do Excel selecionado

## Utilitarios CLI por endpoint

O projeto tambem possui scripts separados por endpoint:

- `bin_presentation_as.py`
- `uptime_as.py`
- `robot_errors_as.py`
- `robot_mtbf_as.py`
- `incidents_as.py`
- `uptime_trend_as.py`

## Novas rotas REST adicionadas

- `GET /v1/installations/{installation_id}/robot-errors/`
- `GET /v1/installations/{installation_id}/robot-mtbf/`
- `GET /v1/installations/{installation_id}/incidents/`
- `GET /v1/installations/{installation_id}/uptime/trend/`

## Abas geradas no Excel

- `BinPresentations`
- `Uptime`
- `RobotErrors`
- `RobotMTBF`
- `Incidents`
- `UptimeTrend`

## Como executar manualmente

1. Ative a virtualenv do projeto.
2. Execute a UI:

```powershell
c:/base/API_AUTOSTORE/.venv/Scripts/python.exe c:/base/API_AUTOSTORE/main.py
```

3. Na tela de acoes:

- informe o ano desejado
- clique em uma das 6 opcoes
- selecione um arquivo `.xlsx` existente
- aguarde a conclusao no log e no status

## Smoke test manual das novas APIs

Fluxo sugerido:

1. Abrir a UI.
2. Rodar `Robot Errors` para um ano com dados recentes.
3. Confirmar criacao/atualizacao da aba `RobotErrors`.
4. Repetir para `Robot MTBF`, `Incidents` e `Uptime Trend`.
5. Validar no log que o `Installation ID` foi obtido e que a contagem de registros foi exibida.

## Testes automatizados

Os testes usam `unittest` com mock do wrapper HTTP, sem chamar a API real:

```powershell
c:/base/API_AUTOSTORE/.venv/Scripts/python.exe -m unittest discover -s c:/base/API_AUTOSTORE/tests -p "test_*.py"
```

## Como gerar o executavel

O repositório nao possui `.spec` nem script de build versionado. O empacotamento pode ser gerado a partir do `main.py`.

1. Instale o PyInstaller na mesma virtualenv do projeto:

```powershell
c:/base/API_AUTOSTORE/.venv/Scripts/python.exe -m pip install pyinstaller
```

2. Gere o executavel:

```powershell
c:/base/API_AUTOSTORE/.venv/Scripts/pyinstaller.exe --noconfirm --onefile --windowed --name API_AUTOSTORE c:/base/API_AUTOSTORE/main.py
```

Observacoes:

- As novas opcoes aparecem automaticamente no executavel porque fazem parte da UI carregada por `main.py`.
- Nao ha icones, templates nem arquivos adicionais versionados que precisem ser incluidos no build atual.
