# Lithium Brine Evaporation Modeling Tool

Herramienta conceptual para modelizar un tren evaporítico de salmueras de litio mediante una interfaz Excel, scripts Python y PHREEQC.

Este repositorio está preparado para que una persona pueda descargar el proyecto, generar los archivos de entrada desde Excel y ejecutar los modelos Python sin depender de rutas locales del ordenador original.

---

## 1. Qué contiene este repositorio

Flujo general de la herramienta:

```text
TFGBrineControl_portable.xlsm
        |
        v
input/*.txt
        |
        v
src/mainret.py
o
src/mainret_recovery_compare.py
        |
        v
PHREEQC + pitzer.dat
        |
        v
runs/<fecha>/
        |
        v
results/<fecha>/
        |
        v
Excel final + CSV + JSON + HTML viewer
```

Estructura principal:

```text
tfg-lithium-brine-modeling/
│
├─ TFGBrineControl_portable.xlsm
├─ README.md
├─ requirements.txt
├─ config.example.toml
├─ run_base.bat
├─ run_recovery_compare.bat
│
├─ input/
│  ├─ feed_base.txt
│  ├─ stages_base.txt
│  ├─ months_control.txt
│  ├─ scenarios_control.txt
│  └─ recovery_control.txt
│
├─ src/
│  ├─ runtime_config.py
│  ├─ mainret.py
│  ├─ mainret_recovery_compare.py
│  └─ mainret_pc1_refinado.py
│
├─ viewer_assets/
│  └─ plotly-2.35.2.min.js
│
└─ documentacion_flujo_interno_python_TFG.pdf
```

Las carpetas `runs/` y `results/` no aparecen inicialmente porque se generan automáticamente al ejecutar el modelo.

---

## 2. Cómo descargar el repositorio desde GitHub

No hace falta saber usar Git para ejecutar la herramienta.

1. Entrar en la página del repositorio en GitHub.
2. Pulsar el botón verde **Code**.
3. Elegir **Download ZIP**.
4. Guardar el archivo `.zip` en el ordenador.
5. Descomprimir el `.zip`.

Después de descomprimirlo, aparecerá una carpeta parecida a:

```text
tfg-lithium-brine-modeling-main/
```

Esa carpeta es la **carpeta raíz del proyecto**. Todas las rutas internas del modelo se calculan a partir de ella.

No se recomienda ejecutar la herramienta directamente dentro del `.zip`. Primero hay que descomprimirlo.

---

## 3. Requisitos previos

Para usar la herramienta se necesita:

- Windows 64 bits.
- Microsoft Excel, para usar `TFGBrineControl_portable.xlsm`.
- Python 3.11 o superior.
- PHREEQC instalado externamente.
- Base termodinámica `pitzer.dat`.
- Dependencias Python indicadas en `requirements.txt`.

PHREEQC no se incluye en este repositorio. Debe instalarse por separado.

---

## 4. Instalación de PHREEQC

PHREEQC debe descargarse desde la página oficial de USGS:

```text
https://www.usgs.gov/software/phreeqc-version-3
```

En la sección de descargas, usar la versión de Windows 64 bits:

```text
Windows 64-bit: phreeqc-3.8.6-17100-x64.msi
Executable, database files, examples, PDF documentation
```

Instalar el archivo `.msi` siguiendo el asistente de Windows.

Después de instalar PHREEQC, hay que localizar dos archivos:

```text
phreeqc.exe
pitzer.dat
```

Normalmente estarán en una ruta parecida a:

```text
C:\Program Files\USGS\phreeqc-3.8.6-17100-x64\bin\Release\phreeqc.exe
C:\Program Files\USGS\phreeqc-3.8.6-17100-x64\database\pitzer.dat
```

En algunos equipos, `pitzer.dat` puede aparecer en otra carpeta de USGS, por ejemplo:

```text
C:\Program Files (x86)\USGS\Phreeqc Interactive 3.8.6-17100\database\pitzer.dat
```

Si no se encuentra, usar el buscador del Explorador de Windows y buscar:

```text
phreeqc.exe
pitzer.dat
```

---

## 5. Configuración de PHREEQC en el proyecto

No hay que modificar los archivos `.py` para poner rutas locales.

El proyecto usa un archivo de configuración local que no se sube a GitHub.

1. En la carpeta raíz del proyecto, localizar:

```text
config.example.toml
```

2. Hacer una copia del archivo y renombrarla como:

```text
config.local.toml
```

3. Abrir `config.local.toml` con el Bloc de notas o Visual Studio Code.

4. Escribir las rutas reales de `phreeqc.exe` y `pitzer.dat`.

Ejemplo:

```toml
[phreeqc]
exe = "C:/Program Files/USGS/phreeqc-3.8.6-17100-x64/bin/Release/phreeqc.exe"
database = "C:/Program Files/USGS/phreeqc-3.8.6-17100-x64/database/pitzer.dat"
```

Importante:

- Usar barras `/` en lugar de `\` para evitar errores de lectura.
- No subir `config.local.toml` a GitHub.
- No editar `runtime_config.py` salvo que se quiera cambiar la lógica interna del proyecto.

---

## 6. Instalación de Python y dependencias

Abrir una terminal en la carpeta raíz del proyecto.

En Windows se puede hacer así:

1. Abrir la carpeta del proyecto.
2. Hacer clic en la barra de dirección del Explorador.
3. Escribir `cmd`.
4. Pulsar Enter.

Después ejecutar:

```bat
python -m venv .venv
.venv\Scripts\activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

Esto crea un entorno virtual local e instala las librerías necesarias.

---

## 7. Uso del Excel de control

El archivo principal de entrada es:

```text
TFGBrineControl_portable.xlsm
```

Este Excel no ejecuta el cálculo geoquímico. Su función es preparar los datos y generar los archivos `.txt` que después leerán los scripts Python.

### 7.1. Hojas del Excel

El libro contiene varias hojas:

```text
Readme
ControlProcess
MainProcess
feed_base
stages_base
months_control
scenarios_control
recovery_control
AUX
```

Uso recomendado:

- `Readme`: hoja de instrucciones internas del Excel.
- `MainProcess`: hoja principal para editar parámetros.
- `ControlProcess`: hoja de control vinculada a la macro.
- `feed_base`, `stages_base`, `months_control`, `scenarios_control`, `recovery_control`: hojas auxiliares usadas para generar los TXT.
- `AUX`: hoja auxiliar interna.

El usuario debe trabajar únicamente desde `MainProcess`.

No deben modificarse manualmente las hojas auxiliares salvo que se conozca la estructura interna del modelo.

### 7.2. Qué celdas se pueden tocar

Como criterio general:

```text
Celdas azules  -> editables por el usuario.
Celdas rojas   -> no tocar.
Celdas grises  -> no tocar.
Celdas con fórmulas o referencias internas -> no tocar.
```

Las celdas rojas, grises o protegidas contienen fórmulas, referencias internas, estructuras de exportación o valores necesarios para que las macros funcionen correctamente.

### 7.3. Parámetros principales editables

#### Feed Base

Define la salmuera inicial:

- `pH`: pH inicial de la salmuera.
- `water_kg`: masa de agua considerada en la solución base.
- `Li`: concentración de litio.
- `K`: concentración de potasio.
- `Mg`: concentración de magnesio.
- `B`: concentración de boro.
- `Ca`: concentración de calcio.
- `Na`: concentración de sodio.
- `Cl`: concentración de cloruro.
- `HCO3-`: concentración de bicarbonato.
- `S6`: azufre en estado de oxidación +6, usado para representar sulfatos en PHREEQC.

#### Stages Base

Define el tren de etapas o piscinas:

- `stage_id`: identificador de etapa, por ejemplo `PC1`, `PC2`, `LIM1`, `H1`, `K1`, `C1` o `L1`.
- `stage_type`: tipo de etapa, normalmente `pond` o `reactor`.
- `surface`: superficie efectiva de evaporación.
- `reagent`: reactivo añadido, si aplica.
- `reagent_param`: parámetro asociado al reactivo.
- `phases`: fases minerales permitidas en PHREEQC.

La columna `phases` es especialmente sensible porque define qué minerales pueden precipitar en cada etapa.

#### Months Control

Controla los factores de evaporación:

- `Mes`: mes o caso de evaporación.
- `E_brine`: evaporación de referencia de la salmuera.
- `f_m`: factor mensual de corrección.
- `TP_PC`: tasa de evaporación para piscinas PC.
- `TP_H`: tasa de evaporación para piscinas H.
- `TP_K`: tasa de evaporación para piscinas K.
- `TP_C`: tasa de evaporación para piscinas C.
- `TP_L`: tasa de evaporación para piscinas L.

#### Scenarios Generator

Define los escenarios o sensibilidades:

- `variable`: variable que se desea modificar.
- `enabled`: activa o desactiva la variable. Normalmente `1` significa activo y `0` inactivo.
- `min`: valor mínimo.
- `max`: valor máximo.
- `step`: incremento del barrido.

Variables principales:

- `months`: mes o meses simulados.
- `temperature`: temperatura de cálculo.
- `evap_factor`: factor global de evaporación.
- `mg_removal`: eliminación de magnesio en la etapa de liming.
- `retention_r`: fracción de salmuera retenida en los sólidos precipitados.

#### Recovery Control

Solo se usa en el modelo comparativo `mainret_recovery_compare.py`.

- `target_freshwater_L_s`: caudal objetivo de agua dulce recuperada.
- `heating_temperature_C`: temperatura interna considerada para el módulo de recuperación.
- `extraction_stage_id`: etapa desde la que se extrae salmuera hacia el módulo auxiliar.
- `reinjection_stage_id`: etapa en la que se reinyecta la salmuera concentrada.

---

## 8. Cómo generar los archivos TXT desde Excel

Antes de ejecutar cualquier script Python, hay que generar los `.txt`.

Si no se generan de nuevo, Python usará los últimos archivos existentes en `input/`, aunque los valores visibles en Excel hayan cambiado.

### Procedimiento

1. Abrir `TFGBrineControl_portable.xlsm`.
2. Pulsar **Habilitar edición**, si Excel lo solicita.
3. Pulsar **Habilitar contenido** o **Activar macros**, si Excel lo solicita.
4. Ir a la hoja `MainProcess`.
5. Modificar solo las celdas azules.
6. Activar la pestaña **Programador** si no aparece.
7. Ejecutar la macro de exportación de TXT.

La macro de exportación se llama normalmente:

```text
ExportarInputsTXT
```

Esta macro genera o actualiza los archivos:

```text
input/feed_base.txt
input/stages_base.txt
input/months_control.txt
input/scenarios_control.txt
input/recovery_control.txt
```

### Cómo activar la pestaña Programador en Excel

1. Ir a **Archivo**.
2. Entrar en **Opciones**.
3. Entrar en **Personalizar cinta de opciones**.
4. Marcar **Programador**.
5. Aceptar.

Después:

```text
Programador -> Macros -> ExportarInputsTXT -> Ejecutar
```

También puede usarse el menú:

```text
Vista -> Macros -> Ver macros -> ExportarInputsTXT -> Ejecutar
```

---

## 9. Scripts Python incluidos

Los scripts están en la carpeta:

```text
src/
```

### `runtime_config.py`

Archivo interno de configuración.

Funciones principales:

- Detecta la carpeta raíz del proyecto.
- Define las rutas relativas a `input/`, `runs/`, `results/` y `viewer_assets/`.
- Lee `config.local.toml`.
- Localiza `phreeqc.exe`.
- Localiza `pitzer.dat`.

No se ejecuta directamente.

### `mainret.py`

Modelo principal del tren evaporítico convencional.

Lee:

```text
input/feed_base.txt
input/stages_base.txt
input/months_control.txt
input/scenarios_control.txt
```

Ejecuta el tren base y genera archivos intermedios de PHREEQC en:

```text
runs/<fecha>_ret/
```

Genera resultados finales en:

```text
results/<fecha>_ret/
```

Usar este script cuando se quiera simular el caso base sin módulo auxiliar de recuperación de agua.

### `mainret_recovery_compare.py`

Modelo comparativo base frente a recuperación de agua.

Lee los mismos archivos que `mainret.py` y además:

```text
input/recovery_control.txt
```

Ejecuta:

1. Caso base.
2. Caso con extracción hacia el módulo de recuperación.
3. Simulación del módulo auxiliar.
4. Reinyección de salmuera concentrada.
5. Comparación base/recovery.

Genera resultados en:

```text
results/<fecha>_ret_recovery_compare/
runs/<fecha>_ret_recovery_compare/
```

Usar este script cuando se quiera evaluar el efecto del módulo auxiliar de recuperación de agua.

### `mainret_pc1_refinado.py`

Script auxiliar para análisis específico o refinado de PC1.

No es el script principal de uso general. Se conserva como herramienta auxiliar para revisión o desarrollo.

---

## 10. Cómo ejecutar el modelo base

Antes de ejecutar:

1. Haber instalado PHREEQC.
2. Haber configurado `config.local.toml`.
3. Haber instalado las dependencias Python.
4. Haber generado los TXT desde Excel.

Después, ejecutar:

```bat
run_base.bat
```

O de forma equivalente:

```bat
python src\mainret.py
```

Salidas principales:

```text
results/<fecha>_ret/
runs/<fecha>_ret/
```

---

## 11. Cómo ejecutar el modelo comparativo con recuperación de agua

Antes de ejecutar:

1. Haber actualizado `recovery_control` desde Excel.
2. Haber generado los TXT.
3. Haber revisado que `input/recovery_control.txt` existe.

Después, ejecutar:

```bat
run_recovery_compare.bat
```

O de forma equivalente:

```bat
python src\mainret_recovery_compare.py
```

Salidas principales:

```text
results/<fecha>_ret_recovery_compare/
runs/<fecha>_ret_recovery_compare/
```

---

## 12. Carpeta `input/`

Contiene los archivos TXT que Python lee como entrada.

```text
input/feed_base.txt
input/stages_base.txt
input/months_control.txt
input/scenarios_control.txt
input/recovery_control.txt
```

Estos archivos se pueden revisar con el Bloc de notas, pero lo recomendable es generarlos desde Excel para evitar errores de formato.

---

## 13. Carpeta `runs/`

La carpeta `runs/` se crea automáticamente.

Conserva los archivos intermedios de PHREEQC para trazabilidad:

```text
input.pqi
output.pqo
selected_output.txt
```

En el modelo comparativo también aparecen carpetas internas para:

```text
base/
recovery/
module/
```

Esta carpeta permite revisar qué se ha enviado a PHREEQC y qué ha devuelto el programa en cada etapa.

---

## 14. Carpeta `results/`

La carpeta `results/` se crea automáticamente.

Contiene las salidas finales:

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

En el comparador recovery se añaden:

```text
recovery_module_results.csv
module_phases.csv
comparison_results.csv
distributed_effect_results.csv
comparison_warnings.csv
recovery_control.csv
recovery_methodology.txt
```

---

## 15. Visor HTML

El archivo:

```text
results_viewer.html
```

permite revisar tablas y gráficos de resultados.

Usa la copia local de Plotly situada en:

```text
viewer_assets/plotly-2.35.2.min.js
```

Por este motivo, el visor puede abrirse sin conexión a internet.

---

## 16. Orden correcto de uso

El orden recomendado es:

```text
1. Descargar y descomprimir el repositorio.
2. Instalar PHREEQC.
3. Localizar phreeqc.exe y pitzer.dat.
4. Crear config.local.toml.
5. Instalar dependencias Python.
6. Abrir TFGBrineControl_portable.xlsm.
7. Activar macros.
8. Editar solo las celdas azules de MainProcess.
9. Ejecutar ExportarInputsTXT.
10. Comprobar que input/*.txt se ha actualizado.
11. Ejecutar run_base.bat o run_recovery_compare.bat.
12. Revisar results/.
```

---

## 17. Errores frecuentes

### Python no encuentra PHREEQC

Revisar `config.local.toml`.

Comprobar que las rutas existen y que están escritas con `/`.

### Python ejecuta datos antiguos

Probablemente no se han regenerado los TXT desde Excel.

Volver a abrir el Excel, ejecutar `ExportarInputsTXT` y repetir la simulación.

### Excel no deja ejecutar macros

Activar macros al abrir el archivo.

Si Excel bloquea el archivo por haber sido descargado de internet:

1. Cerrar Excel.
2. Clic derecho sobre `TFGBrineControl_portable.xlsm`.
3. Propiedades.
4. Marcar **Desbloquear**, si aparece.
5. Aceptar.
6. Abrir de nuevo el archivo.

### No aparece la pestaña Programador

Activarla desde:

```text
Archivo -> Opciones -> Personalizar cinta de opciones -> Programador
```

### No aparecen `runs/` o `results/`

Es normal antes de ejecutar el modelo. Se crean automáticamente.

---

## 18. Nota metodológica

La herramienta es un modelo conceptual y reproducible para comparar escenarios de un tren evaporítico de salmueras de litio.

No constituye una validación industrial ni experimental del proceso. Los resultados deben interpretarse como apoyo al análisis de balances, precipitación mineral, sensibilidad operativa y tendencias geoquímicas.

El modelo no sustituye ensayos de laboratorio, pilotaje, validación industrial, ingeniería de detalle ni estudios ambientales completos.
