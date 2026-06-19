# Lithium Brine Evaporation Modeling Tool

Herramienta conceptual para modelizar un tren evaporitico de salmueras de litio mediante una interfaz Excel, scripts Python y PHREEQC. El flujo general es:

```text
TFGBrineControl_public.xlsm
  -> input/*.txt
  -> src/mainret.py o src/mainret_recovery_compare.py
  -> PHREEQC + pitzer.dat
  -> runs/<fecha>/
  -> results/<fecha>/
  -> Excel final + CSV + JSON + HTML viewer
```

## Estructura Del Repositorio

```text
input/
  feed_base.txt
  stages_base.txt
  months_control.txt
  scenarios_control.txt
  recovery_control.txt
src/
  runtime_config.py
  mainret.py
  mainret_recovery_compare.py
  mainret_pc1_refinado.py
viewer_assets/
  plotly-2.35.2.min.js
TFGBrineControl_public.xlsm
config.example.toml
requirements.txt
run_base.bat
run_recovery_compare.bat
```

Las carpetas `runs/` y `results/` se generan al ejecutar el modelo y no se versionan.

`TFGBrineControl_public.xlsm` es la copia saneada para publicar. Si se desea trabajar con el nombre historico del proyecto, puede copiarse localmente como `TFGBrineControl.xlsm`; ese nombre esta ignorado por Git para evitar publicar metadatos locales.

## Requisitos

- Python 3.11 o superior recomendado.
- PHREEQC instalado externamente.
- Base termodinamica `pitzer.dat` disponible en la instalacion de PHREEQC.
- Dependencias Python indicadas en `requirements.txt`.

PHREEQC no se incluye en este repositorio. Debe instalarse de forma independiente desde USGS o desde la distribucion correspondiente.
Windows 64-bit: phreeqc-3.8.6-17100-x64.msi - Executable, database files, examples, PDF documentation

## Instalacion

Desde la carpeta raiz del proyecto:

```bat
python -m venv .venv
.venv\Scripts\activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

## Configuracion De PHREEQC

El codigo no contiene rutas locales a PHREEQC. Hay dos formas de configurarlo.

Opcion 1: variables de entorno:

```bat
set PHREEQC_EXE=path\to\phreeqc.exe
set PHREEQC_DATABASE=path\to\pitzer.dat
```

Opcion 2: archivo local no versionado:

```bat
copy config.example.toml config.local.toml
```

Editar `config.local.toml`:

```toml
[phreeqc]
exe = "path/to/phreeqc.exe"
database = "path/to/pitzer.dat"
```

`config.local.toml` esta excluido por `.gitignore` para evitar publicar rutas de cada equipo.

## Ejecucion Del Modelo Base

```bat
run_base.bat
```

Equivalente:

```bat
python src\mainret.py
```

Genera resultados en:

```text
results/<fecha>_ret/
runs/<fecha>_ret/
```

## Ejecucion Del Modelo Recovery

```bat
run_recovery_compare.bat
```

Equivalente:

```bat
python src\mainret_recovery_compare.py
```

Genera resultados comparativos en:

```text
results/<fecha>_ret_recovery_compare/
runs/<fecha>_ret_recovery_compare/
```

## Archivos De Entrada

`input/feed_base.txt` define la salmuera inicial: pH, agua y molalidades principales.

`input/stages_base.txt` define el tren de etapas, tipo de etapa, superficie/evaporacion, reactivos y fases minerales habilitadas.

`input/months_control.txt` define meses, factor mensual y tasas de evaporacion por grupo de piscinas.

`input/scenarios_control.txt` define el barrido de escenarios: mes, temperatura, factor de evaporacion, remocion de Mg y retencion de salmuera.

`input/recovery_control.txt` define el caso con modulo auxiliar: agua recuperada objetivo, temperatura interna del modulo, etapa de extraccion y etapa de reinyeccion.

## Carpeta runs/

`runs/` conserva archivos intermedios de PHREEQC para trazabilidad. Incluye, por escenario y etapa:

```text
input.pqi
output.pqo
selected_output.txt
```

En el comparador recovery tambien se generan carpetas internas para `base/`, `recovery/` y `module/`.

## Carpeta results/

`results/` contiene las salidas finales:

```text
simulator_results.xlsx
results_viewer.html
summary_results.csv
evaluation_results.csv
geochemical_results.csv
precipitation_by_phase.csv
scaling_risk_results.csv
scenario_metadata.json
scaling_risk_methodology.txt
```

El comparador recovery anade:

```text
recovery_module_results.csv
module_phases.csv
comparison_results.csv
distributed_effect_results.csv
comparison_warnings.csv
recovery_control.csv
recovery_methodology.txt
```

## Visor HTML

El archivo `results_viewer.html` permite revisar tablas y graficos sin depender de internet. Usa la copia local de Plotly ubicada en `viewer_assets/plotly-2.35.2.min.js`.

## Nota Metodologica

La herramienta es un modelo conceptual y reproducible para comparar escenarios de un tren evaporitico. No constituye una validacion industrial ni experimental del proceso. Los resultados deben interpretarse como apoyo al analisis de balances, precipitacion mineral y tendencias geoquimicas.
