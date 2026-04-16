# Otts - Validación de Horímetros

Script de línea de comandos para revisar órdenes cerradas en SAP/SQL Server, validar consistencia de horímetros y generar reportes CSV/XLSX para control operativo y seguimiento de errores.

## Qué resuelve

El proyecto automatiza dos tareas:

- valida el horímetro reportado contra el historial reciente del mismo equipo;
- genera reportes de control de calidad y de servicio al cliente para una fecha dada.

La validación detecta, entre otros casos:

- horímetros en cero;
- ausencia de horímetro reciente;
- disminución de horas;
- exceso imposible de horas entre dos fechas.

## Requisitos previos

- Python 3.11 recomendado.
- Acceso de red al SQL Server origen.
- Driver ODBC de SQL Server instalado en el equipo.
  - Prioridad por defecto del script:
    - `ODBC Driver 18 for SQL Server`
    - `ODBC Driver 17 for SQL Server`
    - `SQL Server Native Client 11.0`
    - `SQL Server`

## Estructura del proyecto

```text
.
├── src/
│   └── verificar_horimetros.py
├── reportes/              # salida generada en ejecución (ignorada por Git)
├── .env.example
├── .gitignore
├── README.md
└── requirements.txt
```

## Instalación

Crear entorno virtual:

```powershell
python -m venv .venv
```

Activarlo en PowerShell:

```powershell
.venv\Scripts\Activate.ps1
```

Instalar dependencias:

```powershell
pip install -r requirements.txt
```

## Configuración

Crear `.env` a partir de `.env.example`:

```powershell
Copy-Item .env.example .env
```

Variables soportadas:

- `HORI_SERVER`: host o instancia de SQL Server.
- `HORI_DATABASE`: base de datos origen.
- `HORI_USER`: usuario de conexión.
- `HORI_PWD`: contraseña de conexión.
- `HORI_SQL_DRIVER`: driver ODBC a forzar, opcional.
- `HORI_SQL_PORT`: puerto TCP, por defecto `1433`.
- `HORI_SQL_TIMEOUT`: timeout por intento, por defecto `8`.
- `HORI_SQL_ENCRYPT`: valor de cifrado para drivers ODBC modernos.
- `HORI_SQL_TRUST_CERT`: control de `TrustServerCertificate`.
- `HORI_SQL_EXTRA`: parámetros extra para la cadena de conexión.
- `HORI_BASE_DIR`: carpeta base de salida si no se usa `--out`.

Ejemplo mínimo:

```dotenv
HORI_SERVER=mi-servidor
HORI_DATABASE=SBODemo
HORI_USER=usuario
HORI_PWD=secreto
HORI_SQL_PORT=1433
HORI_SQL_TIMEOUT=8
HORI_SQL_ENCRYPT=optional
HORI_SQL_TRUST_CERT=yes
```

## Ejecución

Comando base:

```powershell
python src\verificar_horimetros.py --fecha 2026-04-16
```

Con carpeta de salida explícita:

```powershell
python src\verificar_horimetros.py --fecha 2026-04-16 --out C:\temp\horimetros
```

Ayuda del CLI:

```powershell
python src\verificar_horimetros.py --help
```

## Salida esperada

Si no se indica `--out`, el script genera archivos en:

1. `HORI_BASE_DIR`, si está configurado.
2. `reportes/` en la raíz del proyecto, en caso contrario.

Archivos posibles por ejecución:

- `horimetros_<fecha>.csv`
- `horimetros_<fecha>.xlsx`
- `horimetros_errores_<fecha>.csv`
- `horimetros_errores_<fecha>.xlsx`
- `ots_cerradas_<fecha>.csv`
- `ots_cerradas_<fecha>.xlsx`
- `errores_servicio_<fecha>.csv`
- `errores_servicio_<fecha>.xlsx`

El script imprime al final la carpeta de salida y los nombres generados.

## Troubleshooting

### Faltan credenciales

Si ves un error indicando que faltan credenciales, revisa que `.env` tenga:

- `HORI_SERVER`
- `HORI_DATABASE`
- `HORI_USER`
- `HORI_PWD`

### No conecta a SQL Server

Revisa, en este orden:

1. acceso de red al servidor y al puerto configurado;
2. VPN, firewall o rutas internas;
3. driver ODBC instalado;
4. si el servidor requiere otro puerto o parámetros de cifrado;
5. si conviene fijar `HORI_SQL_DRIVER`.

### El script no genera `.xlsx`

El código tolera la ausencia de `xlsxwriter`, pero el `requirements.txt` ya lo incluye. Si no se genera Excel, valida que las dependencias se instalaron correctamente.

### La salida termina en `reportes/`

Es el comportamiento esperado cuando no se define ni `--out` ni `HORI_BASE_DIR`.

## Política de datos

Este repositorio no debería versionar datos operativos reales ni reportes generados. Por criterio de seguridad y mantenimiento:

- `.env` permanece ignorado;
- `reportes/` permanece ignorado;
- los insumos o exportaciones locales deben vivir fuera del árbol versionado, por ejemplo en `local-data/`;
- si en el futuro se necesitan fixtures, deben ser mínimos, anonimizados y creados específicamente para pruebas o documentación.

## Recomendación sobre los archivos de datos anteriores

Los archivos `demo_colores.xlsx`, `detalle_diario.csv`, `historico_acumulado.csv` y `top3_ok_ko.csv` no son necesarios para ejecutar el script actual y parecen corresponder a datos de trabajo o ejemplos no anonimizados. La decisión aplicada en este repo es:

- sacarlos del árbol versionado actual;
- conservarlos solo como datos locales fuera de Git;
- no usarlos como fixtures oficiales.

Si más adelante se requieren ejemplos en el repositorio, conviene crear una carpeta `samples/` con archivos sintéticos y anonimizados.

## Mejoras siguientes razonables

- separar lógica de acceso a datos, reglas de negocio y exportación;
- agregar pruebas unitarias para clasificación de horímetros y limpieza de `resolution`;
- agregar un modo de validación sin conexión a BD usando fixtures sintéticos;
- normalizar la codificación del archivo fuente para evitar texto mojibake en algunos entornos.
