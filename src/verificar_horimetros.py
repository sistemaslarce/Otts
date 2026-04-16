import argparse, pyodbc, re, os
from datetime import datetime
from collections import Counter
from pathlib import Path
from colorama import Fore, Style, init
import pandas as pd

# Cargar variables de entorno desde .env si existe
try:
    from dotenv import load_dotenv
    load_dotenv()
except Exception:
    pass

# CONFIG PORTABLE
# Credenciales: desde ENV / .env (sin fallback hardcodeado)
SERVER   = os.getenv("HORI_SERVER")   or os.getenv("SQL_SERVER",   "")
DATABASE = os.getenv("HORI_DATABASE") or os.getenv("SQL_DATABASE", "")
USER     = os.getenv("HORI_USER")     or os.getenv("SQL_USERNAME", "")
PWD      = os.getenv("HORI_PWD")      or os.getenv("SQL_PASSWORD", "")
DEFAULT_SQL_DRIVERS = (
    "ODBC Driver 18 for SQL Server",
    "ODBC Driver 17 for SQL Server",
    "SQL Server Native Client 11.0",
    "SQL Server",
)

def resolve_base_dir(cli_out: str | None) -> Path:
    """
    1) --out por CLI
    2) HORI_BASE_DIR (ENV/.env)
    3) <raiz_proyecto>/reportes  (raíz = carpeta que contiene a 'src')
    """
    if cli_out:
        base = Path(cli_out)
    elif os.getenv("HORI_BASE_DIR"):
        base = Path(os.getenv("HORI_BASE_DIR"))
    else:
        project_root = Path(__file__).resolve().parents[1]  # .../horimetros
        base = project_root / "reportes"
    base.mkdir(parents=True, exist_ok=True)
    return base


# QC parámetros
COL = {
    "Correcto":               Fore.GREEN,
    "Horímetros en 0":        Fore.YELLOW,
    "Sin horímetro reciente": Fore.MAGENTA,
    "Exceso en el horímetro": Fore.RED,
    "Horas disminuidas":      Fore.RED,
}

# Colores para Excel por estado de ERROR
XLSX_COLORS = {
    "Sin horímetro reciente": {"bg_color": "#FFEB9C", "font_color": "#000000"},  # amarillo
    "Exceso en el horímetro": {"bg_color": "#F8B88E", "font_color": "#000000"},  # naranja
    "Horas disminuidas":      {"bg_color": "#F9A4A4", "font_color": "#000000"},  # rojo
}


PRIO = {k: i for i, k in enumerate(COL.keys())}
HARD_ERRORS = {"Exceso en el horímetro", "Horas disminuidas"}

num_rx = re.compile(r"^\d+(\.\d+)?$")  # dígitos + punto opcional

# helpers
def env_int(name: str, default: int) -> int:
    try:
        return int(os.getenv(name, "").strip() or default)
    except ValueError:
        return default

def drivers_configurados() -> list[str]:
    raw = os.getenv("HORI_SQL_DRIVER", "").strip()
    if raw:
        return [drv.strip() for drv in re.split(r"[;,]", raw) if drv.strip()]
    return list(DEFAULT_SQL_DRIVERS)

def variantes_server(server: str, port: int) -> list[str]:
    server = server.strip()
    if not server:
        return []
    if server.lower().startswith("tcp:") or "\\" in server or "," in server:
        return [server]
    return [server, f"{server},{port}", f"tcp:{server},{port}"]

def build_connection_candidates() -> list[tuple[str, str]]:
    installed = set(pyodbc.drivers())
    encrypt = os.getenv("HORI_SQL_ENCRYPT", "optional").strip() or "optional"
    trust_cert = os.getenv("HORI_SQL_TRUST_CERT", "yes").strip() or "yes"
    timeout = env_int("HORI_SQL_TIMEOUT", 8)
    port = env_int("HORI_SQL_PORT", 1433)
    extra = os.getenv("HORI_SQL_EXTRA", "").strip()
    candidates: list[tuple[str, str]] = []

    for driver in drivers_configurados():
        if driver not in installed:
            continue
        for server in variantes_server(SERVER, port):
            parts = [
                f"DRIVER={{{driver}}}",
                f"SERVER={server}",
                f"DATABASE={DATABASE}",
                f"UID={USER}",
                f"PWD={PWD}",
                f"Connection Timeout={timeout}",
            ]
            if "ODBC Driver" in driver:
                parts.extend([
                    f"Encrypt={encrypt}",
                    f"TrustServerCertificate={trust_cert}",
                ])
            if extra:
                parts.append(extra.rstrip(";"))
            candidates.append((driver, ";".join(parts) + ";"))

    return candidates

def resumir_error_odbc(exc: Exception) -> str:
    return " ".join(str(exc).replace("\r", " ").replace("\n", " ").split())

def build_connection_error(attempts: list[tuple[str, str]]) -> str:
    timeout = env_int("HORI_SQL_TIMEOUT", 8)
    port = env_int("HORI_SQL_PORT", 1433)
    lines = [
        "No fue posible conectar a SQL Server para generar los reportes.",
        "",
        f"Se probaron {len(attempts)} combinaciones de driver/host con timeout de {timeout}s.",
        "Revisa, en este orden:",
        f"1. Acceso de red al SQL Server y puerto {port} (VPN, firewall o ruta).",
        "2. Que el servidor acepte el driver ODBC disponible en este equipo.",
        "3. Si hace falta otro puerto o cifrado, ajusta HORI_SQL_PORT / HORI_SQL_ENCRYPT / HORI_SQL_TRUST_CERT en .env.",
        "4. Si necesitas forzar un driver, define HORI_SQL_DRIVER.",
    ]
    if attempts:
        lines.append("")
        lines.append("Últimos errores:")
        for driver, err in attempts[-4:]:
            lines.append(f"- {driver}: {err}")
    return "\n".join(lines)

def conectar():
    if not all([SERVER, DATABASE, USER, PWD]):
        raise SystemExit("Faltan credenciales de BD. Configura SQL_SERVER, SQL_DATABASE, SQL_USERNAME, SQL_PASSWORD en .env")
    attempts: list[tuple[str, str]] = []
    for driver, conn_str in build_connection_candidates():
        try:
            return pyodbc.connect(conn_str)
        except pyodbc.Error as exc:
            attempts.append((driver, resumir_error_odbc(exc)))

    if not attempts:
        raise SystemExit(
            "No hay un driver ODBC de SQL Server compatible instalado. "
            "Instala ODBC Driver 17/18 o define HORI_SQL_DRIVER con un driver disponible."
        )
    raise SystemExit(build_connection_error(attempts))

def es_num(s):
    if pd.isna(s):
        return False
    return isinstance(s, (int, float)) or (isinstance(s, str) and num_rx.match(s))

def es_num_pos(s): return es_num(s) and float(s) > 0
def es_missing(s): return (s is None) or (str(s).strip() == '') or (es_num(s) and float(s) == 0)
def texto_vacio(s): return pd.isna(s) or str(s).strip() == ""

def clasifica(h1, h2, f1, f2):
    h1, h2 = float(h1), float(h2)
    dias, diff = (f1 - f2).days, h1 - h2
    max_h = dias * 24
    if h1 == h2 == 0:  return "Horímetros en 0", dias, diff, max_h
    if dias <= 0:      return "Fechas invertidas/iguales", dias, diff, max_h
    if diff < 0:       return "Horas disminuidas", dias, diff, max_h
    if diff > max_h:   return "Exceso en el horímetro", dias, diff, max_h
    return "Correcto", dias, diff, max_h

# SQL
SQL_HIST_BATCH = """
SELECT callID AS Call_ID, U_Tecnico AS Tecnico, manufSN AS Numero_Serie,
       U_Horimetro AS Horimetro, createDate AS Fecha
FROM OSCL
WHERE status = -1
  AND manufSN IN (
      SELECT DISTINCT manufSN FROM OSCL
      WHERE status = -1 AND CAST(createDate AS date) = ?
        AND manufSN IS NOT NULL AND manufSN <> ''
  )
ORDER BY manufSN, Fecha DESC;"""

SQL_OT_CERRADAS = """
SELECT callID AS [Call ID], createDate AS [Fecha OT], U_Horimetro AS [Horímetro],
       U_Tecnico AS [Técnico], manufSN AS [Número de Serie], resolution
FROM OSCL
WHERE status = -1 AND CAST(createDate AS date) = ?
ORDER BY createDate DESC;"""

# consultas
def historiales_dia(conn, dia):
    """Retorna dict {manufSN: [rows...]} con el historial de cada equipo activo ese día."""
    cur = conn.cursor()
    try:
        rows = cur.execute(SQL_HIST_BATCH, dia).fetchall()
    finally:
        cur.close()
    por_sn = {}
    for r in rows:
        por_sn.setdefault(r.Numero_Serie, []).append(r)
    return por_sn

def ots_cerradas_dia(conn, dia):
    cur = conn.cursor()
    try:
        rows = cur.execute(SQL_OT_CERRADAS, dia).fetchall()
        cols = [c[0] for c in cur.description]
    finally:
        cur.close()
    df = pd.DataFrame(map(tuple, rows), columns=cols)
    if not df.empty and "resolution" in df.columns:
        df["resolution"] = df.apply(
            lambda row: limpiar_resolution(row["resolution"], row.get("Técnico", ""), row.get("Fecha OT")),
            axis=1,
        )
    return df

def limpiar_resolution(texto, tecnico="", fecha=None):
    if texto is None:
        return ""

    texto = str(texto).replace("\r\n", "\n").replace("\r", "\n").strip()
    if not texto:
        return ""

    tecnico = str(tecnico or "").strip().lower()
    fecha_txt = ""
    if fecha is not None and str(fecha).strip():
        fecha_dt = pd.to_datetime(fecha, errors="coerce")
        fecha_txt = "" if pd.isna(fecha_dt) else fecha_dt.strftime("%Y-%m-%d")

    lineas_limpias = []
    vistos = set()

    for linea in texto.split("\n"):
        linea = linea.strip()
        if not linea:
            continue

        linea = re.sub(r"^\s*WORK\s+PERFORMED\s*:\s*", "", linea, flags=re.IGNORECASE).strip()
        if not linea:
            continue

        linea_norm = re.sub(r"\s+", " ", linea).strip()
        linea_key = linea_norm.lower()

        if linea_key in {"all ok", "ok", "todo ok"}:
            continue

        # Quita firmas tipo "Nombre Apellido (2026-02-03)" aunque el técnico
        # en la fila venga vacío o desplazado.
        if re.fullmatch(r"[A-Za-zÁÉÍÓÚáéíóúÑñ]+(?:\s+[A-Za-zÁÉÍÓÚáéíóúÑñ]+)+\s*\(\d{4}-\d{2}-\d{2}\)", linea_norm):
            continue

        # Quita fechas sueltas al final.
        if re.fullmatch(r"\d{4}-\d{2}-\d{2}", linea_norm):
            continue

        if tecnico:
            tecnico_con_fecha = f"{tecnico} ({fecha_txt})" if fecha_txt else tecnico
            if linea_key in {tecnico, tecnico_con_fecha}:
                continue

        if linea_key in vistos:
            continue

        vistos.add(linea_key)
        lineas_limpias.append(linea_norm)

    return "\n".join(lineas_limpias)

# QC filas
def resumen(r1, estado, faltas=""):
    return {
        "ERROR": estado,
        "Estado OT": "Único",
        "Call ID": r1.Call_ID,
        "Técnico": r1.Tecnico or "",
        "Número de Serie": r1.Numero_Serie,
        "Fecha de Cierre": r1.Fecha.date(),
        "Horímetro": r1.Horimetro,
        "Faltas Horímetro": faltas,
    }

def contar_faltas(hist, idx):
    return sum(1 for r in hist[idx:] if es_missing(r.Horimetro))

def procesar(hist, corte):
    r1 = next((r for r in hist if r.Fecha.date() == corte), None)
    if not r1: return None
    idx = hist.index(r1)

    if es_num_pos(r1.Horimetro):
        prev = next((r for r in hist[idx+1:] if es_num_pos(r.Horimetro)), None)
        if not prev:
            return resumen(r1, "Correcto")
        estado, *_ = clasifica(r1.Horimetro, prev.Horimetro, r1.Fecha, prev.Fecha)
        return resumen(r1, estado)

    faltas = contar_faltas(hist, idx)
    return resumen(r1, "Sin horímetro reciente", faltas)

def es_error_relevante(r):
    if r["ERROR"] in HARD_ERRORS: return True
    return r["ERROR"] == "Sin horímetro reciente" and str(r["Faltas Horímetro"]).strip() == "1"

# Motivos Servicio Cliente
def motivos_sc(row, dup_ids):
    motivos = []
    if row["Call ID"] in dup_ids:                    motivos.append("OT duplicada")
    if texto_vacio(row["Número de Serie"]):          motivos.append("Sin número de serie")

    if "http" in str(row["resolution"]).lower():
        motivos.append("Link en resolución")

    return "; ".join(motivos)

def reporte_sc(df):
    dup_ids = df["Call ID"][df["Call ID"].duplicated(keep=False)].unique()
    df["Motivos SC"] = df.apply(lambda r: motivos_sc(r, dup_ids), axis=1)
    return df[df["Motivos SC"] != ""]

# export util

def aplicar_formato_excel(wb, ws, df: pd.DataFrame):
    header_fmt = wb.add_format({"bold": True, "bg_color": "#F47C20", "font_color": "#FFFFFF"})
    centered_fmt = wb.add_format({"align": "center", "valign": "vcenter"})
    centered_date_fmt = wb.add_format({"align": "center", "valign": "vcenter", "num_format": "yyyy-mm-dd"})
    text_fmt = wb.add_format({"valign": "vcenter"})
    date_cols = {"Fecha OT", "Fecha de Cierre"}

    ws.write_row(0, 0, [c if (str(c).strip()) else "" for c in df.columns], header_fmt)

    nrows, ncols = df.shape
    ws.autofilter(0, 0, nrows, ncols - 1)
    ws.freeze_panes(1, 0)

    centered_cols = {"Call ID", "Fecha OT", "Fecha de Cierre", "Horímetro", "Técnico", "Número de Serie", "Motivos SC"}
    width_overrides = {
        "Call ID": 12,
        "Fecha OT": 14,
        "Fecha de Cierre": 14,
        "Horímetro": 14,
        "Técnico": 20,
        "Número de Serie": 18,
        "resolution": 48,
        "Motivos SC": 28,
    }

    for i, col in enumerate(df.columns):
        try:
            max_len = max(len(str(col)), int(df[col].astype(str).str.len().max()))
        except Exception:
            max_len = len(str(col))

        width = min(60, max(12, max_len + 2))
        width = max(width, width_overrides.get(col, 0))
        fmt = centered_date_fmt if col in date_cols else centered_fmt if col in centered_cols else text_fmt
        ws.set_column(i, i, width, fmt)

        if col in date_cols:
            for row_idx, value in enumerate(df[col], start=1):
                if pd.isna(value) or value == "":
                    ws.write_blank(row_idx, i, None, centered_date_fmt)
                    continue
                date_value = pd.to_datetime(value, errors="coerce")
                if pd.isna(date_value):
                    ws.write(row_idx, i, value, centered_fmt)
                else:
                    ws.write_datetime(row_idx, i, date_value.to_pydatetime(), centered_date_fmt)

def export(df, nombre, tag, base_dir: Path):
    csv_path = base_dir / f"{nombre}_{tag}.csv"
    xls_path = base_dir / f"{nombre}_{tag}.xlsx"
    df.to_csv(csv_path, index=False, encoding="utf-8")

    try:
        import xlsxwriter  # asegura el engine
        with pd.ExcelWriter(xls_path, engine="xlsxwriter", datetime_format="yyyy-mm-dd") as writer:
            df.to_excel(writer, index=False, header=False, startrow=1, sheet_name="Hoja1")
            wb = writer.book
            ws = writer.sheets["Hoja1"]
            aplicar_formato_excel(wb, ws, df)
    except ModuleNotFoundError:
        pass

    return csv_path.name, xls_path.name if xls_path.exists() else "(sin .xlsx)"

def export_xlsx_por_error(df: pd.DataFrame, xlsx_path: Path, pintar_fila_completa: bool = False):
    """
    - Encabezado (fila 0) naranja #F47C20 y texto blanco (solo si hay texto).
    - Datos empiezan en fila 1.
    - Si existe 'ERROR':
        * False -> colorea solo la celda 'ERROR' por fila.
        * True  -> colorea toda la fila.
    """
    import xlsxwriter  # asegura el engine
    with pd.ExcelWriter(xlsx_path, engine="xlsxwriter", datetime_format="yyyy-mm-dd") as writer:
        df.to_excel(writer, index=False, header=False, startrow=1, sheet_name="Hoja1")
        wb = writer.book
        ws = writer.sheets["Hoja1"]
        aplicar_formato_excel(wb, ws, df)

        if "ERROR" in df.columns:
            state_fmt = {k: wb.add_format({**v, "bold": True}) for k, v in XLSX_COLORS.items()}
            nrows, ncols = df.shape
            col_error_idx = df.columns.get_loc("ERROR")

            for i in range(nrows):
                val = str(df.iloc[i, col_error_idx])
                fmt = state_fmt.get(val)
                if not fmt:
                    continue
                excel_row = i + 1
                if pintar_fila_completa:
                    for j in range(ncols):
                        ws.write(excel_row, j, df.iloc[i, j], fmt)
                else:
                    ws.write(excel_row, col_error_idx, df.iloc[i, col_error_idx], fmt)


def export_coloreado_por_error(df, nombre, tag, base_dir: Path):
    csv_path = base_dir / f"{nombre}_{tag}.csv"
    xls_path = base_dir / f"{nombre}_{tag}.xlsx"
    df.to_csv(csv_path, index=False, encoding="utf-8")
    try:
        export_xlsx_por_error(df, xls_path, pintar_fila_completa=False)
    except ModuleNotFoundError:
        pass
    return csv_path.name, xls_path.name if xls_path.exists() else "(sin .xlsx)"


QC_COLS = ["ERROR", "Call ID", "Técnico", "Número de Serie",
           "Fecha de Cierre", "Horímetro", "Faltas Horímetro", "Estado OT"]

def preparar_df_qc(rows):
    df = pd.DataFrame(rows)[QC_COLS]
    df["Estado OT"] = df["Estado OT"].replace({"Único": ""})
    df["ERROR"] = df["ERROR"].replace({"Correcto": "", "Fechas invertidas/iguales": ""})
    return df

# main
def main():
    init(autoreset=True)
    pa = argparse.ArgumentParser()
    pa.add_argument("--fecha", required=True, help="AAAA-MM-DD")
    pa.add_argument("--out", help="Carpeta de salida (opcional). "
                                  "Si no se indica, usa HORI_BASE_DIR o reportes/ en la raíz del proyecto.")
    args = pa.parse_args()

    dia = datetime.strptime(args.fecha, "%Y-%m-%d").date()
    tag = args.fecha
    base_dir = resolve_base_dir(args.out)
    conn = conectar()

    try:
        # QC horímetro
        qc_rows = []
        for sn, hist in historiales_dia(conn, dia).items():
            if (row := procesar(hist, dia)):
                qc_rows.append(row)

        dup_counts = Counter(r["Call ID"] for r in qc_rows)
        for r in qc_rows:
            if dup_counts[r["Call ID"]] > 1:
                r["Estado OT"] = "Duplicada"

        qc_rows.sort(key=lambda r: (PRIO.get(r["ERROR"], 99), r["Número de Serie"]))
        qc_err = [r for r in qc_rows if es_error_relevante(r)]

        out = []

        if qc_rows:
            out.append(export_coloreado_por_error(
                preparar_df_qc(qc_rows), "horimetros", tag, base_dir))

        if qc_err:
            out.append(export_coloreado_por_error(
                preparar_df_qc(qc_err), "horimetros_errores", tag, base_dir))

        # Servicio Cliente
        df_c = ots_cerradas_dia(conn, dia)
        if not df_c.empty:
            out.append(export(df_c, "ots_cerradas", tag, base_dir))
            df_sc = reporte_sc(df_c.copy())
            if not df_sc.empty:
                out.append(export(df_sc, "errores_servicio", tag, base_dir))
    finally:
        conn.close()

    print(Style.BRIGHT + f"\nArchivos generados en: {base_dir}" + Style.RESET_ALL)
    for c, x in out:
        print(f"  {c}\n  {x}")

if __name__ == "__main__":
    main()
