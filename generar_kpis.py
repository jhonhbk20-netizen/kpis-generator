# -*- coding: utf-8 -*-
"""
============================================================================
 GENERADOR DE KPIs — PAGOS MASIVOS Y CTS
============================================================================
Lee automaticamente los reportes crudos de Excel/CSV que dejes en ESTA
carpeta y genera un unico archivo Excel con todos los KPIs calculados,
listos para armar las diapositivas del analisis mensual.

COMO USARLO
-----------
1) Deja en esta carpeta los reportes crudos (pueden tener sufijos como
   "(1)", "(2)", copia, etc. — el script los detecta por su nombre):
       - Detalle de Pago Lote y CTS ....xlsx      (Pago Lote + Deposito Lote CTS)
       - Detalle CTS Individual ....xlsx          (CTS individual por cliente)
       - Extornos ....xlsx
       - Abono de Sueldos y CTS ....xlsx
       - REQUERIMIENTOS BACKOFFICE ....csv
   Si hay varios que coinciden, usa el MAS RECIENTE (fecha de modificacion).
   (No tienes estos reportes? Corre antes "generar_datos_demo.py" para crear
   una version ficticia de cada uno y probar el flujo completo.)
2) Ejecuta:  python generar_kpis.py
3) Se crea:  KPIs_Pagos_CTS_<MES>.xlsx  (con hoja de Explicacion).

No necesitas editar rutas. Si quieres fijar el mes destacado, cambia
MES_OBJETIVO abajo; si lo dejas en None se destaca el ultimo mes con datos.
============================================================================
"""

import os
import re
import glob
import unicodedata
import datetime as dt
import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.marker import Marker
from openpyxl.drawing.text import CharacterProperties

# ---------------------------------------------------------------------------
# CONFIGURACION (ajustable) --------------------------------------------------
# ---------------------------------------------------------------------------
CARPETA        = os.path.dirname(os.path.abspath(__file__))   # esta misma carpeta
MES_OBJETIVO   = None          # ej. "2026-07"; None = ultimo mes con datos
TOTAL_TIENDAS_RED = 62         # denominador de la cobertura: cantidad de agencias de tu red

# Roles de operadores (usuario del core -> rol). Ajusta a tu propio equipo.
ROLES = {
    "ANALISTA1": "Analista",
    "ASISTENTE1": "Asistente 1",
    "ASISTENTE2": "Asistente 2",
}
ANALISTAS = {u for u, r in ROLES.items() if r.lower().startswith("analista")}
ASISTENTES = {u for u, r in ROLES.items() if r.lower().startswith("asistente")}

# Tienda que representa a nuestro equipo (backoffice) dentro del reporte de
# Abono Sueldos y CTS, que junta TODAS las agencias de la red.
TIENDA_PROPIA = "OFICINA PRINCIPAL"

# Supuestos de tiempo de procesamiento (ESTIMADO, no cronometrado)
MIN_POR_PAGO_LOTE = 1.3        # minutos por transaccion de Pago Lote
MIN_POR_CTS       = 4.0        # minutos por transaccion de CTS
HORAS_JORNADA     = 8.0        # horas por dia laboral (para pasar a dias-persona)

# ---------------------------------------------------------------------------
# UTILIDADES -----------------------------------------------------------------
# ---------------------------------------------------------------------------
def norm(texto):
    """minusculas + sin tildes + espacios colapsados (para comparar textos)."""
    if texto is None:
        return ""
    s = unicodedata.normalize("NFKD", str(texto))
    s = "".join(c for c in s if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", s).strip().lower()


def buscar_archivo(*claves, ext=("xlsx", "xls", "csv")):
    """Devuelve el archivo mas reciente cuyo nombre contenga TODAS las claves."""
    candidatos = []
    for e in ext:
        candidatos += glob.glob(os.path.join(CARPETA, f"*.{e}"))
    def coincide(nombre):
        base = os.path.basename(nombre)
        if base.startswith("~$"):   # archivo temporal de Office (esta abierto)
            return False
        low = base.lower()
        return all(k.lower() in low for k in claves)
    hits = [c for c in candidatos if coincide(c)]
    if not hits:
        return None
    return max(hits, key=os.path.getmtime)   # el mas reciente


def mes_de(valor):
    """Convierte una fecha/valor a 'YYYY-MM' o None."""
    if valor is None or valor == "":
        return None
    if isinstance(valor, (dt.datetime, dt.date)):
        return f"{valor.year:04d}-{valor.month:02d}"
    ts = pd.to_datetime(valor, errors="coerce", dayfirst=True)
    return None if pd.isna(ts) else ts.strftime("%Y-%m")


def fecha_de(valor):
    if isinstance(valor, dt.datetime):
        return valor.date()
    if isinstance(valor, dt.date):
        return valor
    ts = pd.to_datetime(valor, errors="coerce", dayfirst=True)
    return None if pd.isna(ts) else ts.date()


def a_numero(v):
    if v is None or v == "":
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(",", "").strip())
    except ValueError:
        return 0.0


def leer_filas(ruta):
    """Lee todas las filas de la primera hoja como listas (valores)."""
    wb = openpyxl.load_workbook(ruta, read_only=True, data_only=True)
    filas = list(wb.active.iter_rows(values_only=True))
    wb.close()
    return filas


def mapa_columnas(fila_encabezado):
    """De una fila de encabezado construye {nombre_normalizado: indice}."""
    m = {}
    for i, c in enumerate(fila_encabezado):
        if c is None:
            continue
        nombre = norm(c)
        if nombre and nombre not in m:
            m[nombre] = i
    return m


def col(m, *nombres):
    """Devuelve el indice de la primera columna que coincida (por 'contiene')."""
    for n in nombres:
        n = norm(n)
        if n in m:
            return m[n]
    for n in nombres:
        n = norm(n)
        for nombre, idx in m.items():
            if n in nombre:
                return idx
    return None


# ---------------------------------------------------------------------------
# LECTORES DE CADA FUENTE ----------------------------------------------------
# ---------------------------------------------------------------------------
def leer_detalle_pago_lote(ruta):
    """
    Archivo con DOS sub-tablas apiladas:
      Seccion 1 = PAGO LOTE   -> segmento 'PAGO'
      Seccion 2 = DEPOSITO LOTE CTS -> segmento 'CTS'
    Se detecta cada encabezado por la columna 'Numero de movimiento' y se
    mapean columnas por NOMBRE (las dos secciones estan desplazadas).
    """
    filas = leer_filas(ruta)
    encabezados = [i for i, f in enumerate(filas)
                   if any("numero de movimiento" in norm(c) for c in f)]
    encabezados.append(len(filas))  # centinela final
    recs = []
    for k in range(len(encabezados) - 1):
        h = encabezados[k]
        fin = encabezados[k + 1]
        m = mapa_columnas(filas[h])
        i_mov  = col(m, "numero de movimiento")
        i_tipo = col(m, "tipo de operacion", "tipo de  operacion", "tipo")
        i_mon  = col(m, "monto")
        i_fec  = col(m, "fecha y hora", "fecha")
        i_usu  = col(m, "usuario")
        i_tie  = col(m, "tienda")
        i_emp  = col(m, "empresa", "empleador")
        for f in filas[h + 1:fin]:
            if i_mov is None or i_mov >= len(f):
                continue
            mov = f[i_mov]
            if mov in (None, "") or "numero de movimiento" in norm(mov):
                continue
            tipo = str(f[i_tipo]).strip().upper() if i_tipo is not None and i_tipo < len(f) and f[i_tipo] else ""
            if not tipo:
                continue
            segmento = "PAGO" if "PAGO LOTE" in tipo else ("CTS" if "CTS" in tipo else "OTRO")
            fval = f[i_fec] if (i_fec is not None and i_fec < len(f)) else None
            recs.append({
                "segmento": segmento,
                "tipo": tipo,
                "mov": mov,
                "monto": a_numero(f[i_mon]) if (i_mon is not None and i_mon < len(f)) else 0.0,
                "mes": mes_de(fval),
                "dia": fecha_de(fval),
                "usuario": (str(f[i_usu]).strip() if (i_usu is not None and i_usu < len(f) and f[i_usu]) else None),
                "tienda": (str(f[i_tie]).strip() if (i_tie is not None and i_tie < len(f) and f[i_tie]) else None),
                "empleador": (str(f[i_emp]).strip() if (i_emp is not None and i_emp < len(f) and f[i_emp]) else None),
            })
    df = pd.DataFrame(recs)
    return df[df["mes"].notna()].copy()


def leer_detalle_cts(ruta):
    """CTS individual por cliente (APERTURA / DEPOSITO / CANCELACION)."""
    filas = leer_filas(ruta)
    hi = next((i for i, f in enumerate(filas)
               if any("de movimiento" in norm(c) for c in f)), None)
    if hi is None:
        return pd.DataFrame()
    m = mapa_columnas(filas[hi])
    i_mov = col(m, "de movimiento", "n° de movimiento")
    i_fec = col(m, "fecha")
    i_tie = col(m, "tienda")
    i_emp = col(m, "empresa", "empleador")
    i_tipo = col(m, "tipo operacion", "tipo operación", "tipo")
    i_mon = col(m, "monto")
    recs = []
    for f in filas[hi + 1:]:
        if i_mov is None or i_mov >= len(f) or f[i_mov] in (None, ""):
            continue
        fval = f[i_fec] if (i_fec is not None and i_fec < len(f)) else None
        recs.append({
            "mes": mes_de(fval),
            "tienda": str(f[i_tie]).strip() if (i_tie is not None and i_tie < len(f) and f[i_tie]) else None,
            "empleador": str(f[i_emp]).strip() if (i_emp is not None and i_emp < len(f) and f[i_emp]) else None,
            "tipo": str(f[i_tipo]).strip().upper() if (i_tipo is not None and i_tipo < len(f) and f[i_tipo]) else "",
            "monto": a_numero(f[i_mon]) if (i_mon is not None and i_mon < len(f)) else 0.0,
        })
    df = pd.DataFrame(recs)
    return df[df["mes"].notna()].copy() if not df.empty else df


def leer_extornos(ruta):
    """Un extorno = un 'N° Mov Extorno' unico (varias filas = un evento)."""
    filas = leer_filas(ruta)
    hi = next((i for i, f in enumerate(filas)
               if any("mov extorno" in norm(c) for c in f)), None)
    if hi is None:
        return pd.DataFrame()
    m = mapa_columnas(filas[hi])
    i_ext = col(m, "mov extorno", "n° mov extorno")
    i_fec = col(m, "fecha extorno", "fecha")
    i_mon = col(m, "monto")
    i_mot = col(m, "motivo")
    recs = []
    for f in filas[hi + 1:]:
        if i_ext is None or i_ext >= len(f) or f[i_ext] in (None, ""):
            continue
        fval = f[i_fec] if (i_fec is not None and i_fec < len(f)) else None
        recs.append({
            "extorno": f[i_ext],
            "mes": mes_de(fval),
            "monto": a_numero(f[i_mon]) if (i_mon is not None and i_mon < len(f)) else 0.0,
            "motivo": str(f[i_mot]).strip().upper() if (i_mot is not None and i_mot < len(f) and f[i_mot]) else "SIN MOTIVO",
        })
    df = pd.DataFrame(recs)
    return df[df["mes"].notna()].copy() if not df.empty else df


def leer_abono(ruta):
    """Abono de Sueldos/CTS. No tiene Monto; tiene columna Estado."""
    filas = leer_filas(ruta)
    hi = next((i for i, f in enumerate(filas)
               if any("fecha operaci" in norm(c) for c in f)), None)
    if hi is None:
        return pd.DataFrame()
    m = mapa_columnas(filas[hi])
    i_fec = col(m, "fecha operacion", "fecha operación", "fecha")
    i_tipo = col(m, "tipo operacion", "tipo operación", "tipo")
    i_est = col(m, "estado")
    i_emp = col(m, "nombre empleador", "empleador")
    i_can = col(m, "canal")
    i_tie = col(m, "tienda operacion de abono", "tienda")
    recs = []
    for f in filas[hi + 1:]:
        if i_fec is None or i_fec >= len(f) or f[i_fec] in (None, ""):
            continue
        recs.append({
            "mes": mes_de(f[i_fec]),
            "tipo": str(f[i_tipo]).strip().upper() if (i_tipo is not None and i_tipo < len(f) and f[i_tipo]) else "",
            "estado": str(f[i_est]).strip().upper() if (i_est is not None and i_est < len(f) and f[i_est]) else "",
            "empleador": str(f[i_emp]).strip() if (i_emp is not None and i_emp < len(f) and f[i_emp]) else None,
            "canal": str(f[i_can]).strip().upper() if (i_can is not None and i_can < len(f) and f[i_can]) else "",
            "tienda": str(f[i_tie]).strip().upper() if (i_tie is not None and i_tie < len(f) and f[i_tie]) else "",
        })
    df = pd.DataFrame(recs)
    return df[df["mes"].notna()].copy() if not df.empty else df


def leer_requerimientos(ruta):
    """Buzon de requerimientos (CSV separado por ';', encabezado en fila 3)."""
    for enc in ("utf-8-sig", "latin-1", "cp1252"):
        try:
            raw = pd.read_csv(ruta, sep=";", encoding=enc, header=None,
                              dtype=str, keep_default_na=False, engine="python")
            break
        except Exception:
            raw = None
    if raw is None or raw.empty:
        return pd.DataFrame()
    # La fila de encabezado real tiene MUCHAS columnas no vacias (a diferencia
    # de la fila de titulo, que solo tiene una celda con texto).
    hi = next((i for i in range(min(8, len(raw)))
               if raw.iloc[i].astype(str).str.contains("SOLICITUD", case=False).any()
               and (raw.iloc[i].astype(str).str.strip() != "").sum() >= 5), 2)
    header = [re.sub(r"\s+", " ", str(c)).strip() for c in raw.iloc[hi]]
    data = raw.iloc[hi + 1:].copy()
    data.columns = header
    data = data.dropna(how="all")
    return data


# ---------------------------------------------------------------------------
# CALCULO DE KPIs ------------------------------------------------------------
# ---------------------------------------------------------------------------
def calcular(det, cts, ext, abo, req):
    # Los meses del analisis salen del reporte principal (Detalle Pago Lote y CTS),
    # que es el que registra la actividad operativa de Pago Lote y CTS lote.
    # El archivo CTS individual solo se usa para cobertura/empleadores del
    # mismo mes.
    meses = sorted(det["mes"].dropna().unique().tolist()) if not det.empty else []

    resumen, operadores, extornos_k8, tiempo_k9, eficiencia_k7 = [], [], [], [], []

    for mes in meses:
        d = det[det["mes"] == mes]
        pago = d[d["segmento"] == "PAGO"]
        ctsl = d[d["segmento"] == "CTS"]

        k1 = len(pago)
        k2 = pago["monto"].sum()
        k3 = len(ctsl)
        k4 = ctsl["monto"].sum()
        total_ops = k1 + k3
        total_monto = k2 + k4

        # --- por operador ---
        ops_usu = d.groupby("usuario").size().to_dict()
        # K5 compara SOLO a los asistentes nombrados en ROLES (ignora apoyos
        # esporadicos de otras tiendas que a veces aparecen en el crudo).
        asist = {u: n for u, n in ops_usu.items() if u in ASISTENTES}
        analista_ops = sum(n for u, n in ops_usu.items() if u in ANALISTAS)
        if len(asist) >= 2:
            k5 = round(max(asist.values()) / max(min(asist.values()), 1), 2)
        else:
            k5 = np.nan
        k5a = analista_ops / total_ops if total_ops else 0.0

        # cobertura y empleadores
        tiendas_pago = pago["tienda"].dropna().nunique()
        cts_mes = cts[cts["mes"] == mes] if not cts.empty else pd.DataFrame()
        if not cts_mes.empty:
            tiendas_cts = cts_mes["tienda"].dropna().nunique()
            empleadores = cts_mes["empleador"].dropna().nunique()
        else:  # respaldo: usar la seccion 2 del detalle
            tiendas_cts = ctsl["tienda"].dropna().nunique()
            empleadores = ctsl["empleador"].dropna().nunique()
        cob_pago = tiendas_pago / TOTAL_TIENDAS_RED if TOTAL_TIENDAS_RED else 0
        cob_cts = tiendas_cts / TOTAL_TIENDAS_RED if TOTAL_TIENDAS_RED else 0

        # dias / eficiencia: la eficiencia del EQUIPO (persona-dias, operadores
        # activos) se mide solo sobre el personal fijo definido en ROLES.
        # Los apoyos esporadicos de otras tiendas (no listados en ROLES) SI
        # suman al volumen total (arriba), pero no inflan el denominador de
        # productividad del equipo — asi lo hacia el KPI original.
        dd = d[d["usuario"].isin(ROLES)].dropna(subset=["dia", "usuario"])
        persona_dias = dd.groupby(["usuario", "dia"]).ngroups
        dias_calendario = dd["dia"].nunique()
        operadores_activos = dd["usuario"].nunique()
        ops_dia_persona = total_ops / persona_dias if persona_dias else 0
        ops_dia_calend = total_ops / (operadores_activos * dias_calendario) if operadores_activos and dias_calendario else 0

        resumen.append({
            "Mes": mes, "K1 Vol Pago": k1, "K2 Monto Pago": round(k2, 2),
            "K3 Vol CTS": k3, "K4 Monto CTS": round(k4, 2),
            "Total Ops": total_ops, "Total Monto": round(total_monto, 2),
            "K5 Ratio Asist.": "N/A" if pd.isna(k5) else k5, "K5a % Analista": round(k5a, 4),
            "K7 % Cobert. Pago": round(cob_pago, 4), "K8 % Cobert. CTS": round(cob_cts, 4),
            "K9 Empleadores": empleadores, "Operadores": operadores_activos,
            "Persona-dias": persona_dias, "Dias calendario": dias_calendario,
            "Ops/dia (persona)": round(ops_dia_persona, 1),
            "Ops/dia (calend.)": round(ops_dia_calend, 1),
        })

        # detalle por operador
        for u in sorted([x for x in ops_usu if x]):
            du = d[d["usuario"] == u]
            dpago = du[du["segmento"] == "PAGO"]
            dcts = du[du["segmento"] == "CTS"]
            dias_u = du["dia"].dropna().nunique()
            tot_u = len(du)
            operadores.append({
                "Mes": mes, "Usuario": u, "Rol": ROLES.get(u, "Apoyo/Otro"),
                "Ops Pago": len(dpago), "Monto Pago": round(dpago["monto"].sum(), 2),
                "Ops CTS": len(dcts), "Monto CTS": round(dcts["monto"].sum(), 2),
                "Total Ops": tot_u, "Total Monto": round(du["monto"].sum(), 2),
                "Dias Activos": dias_u, "Ops/Dia": round(tot_u / dias_u, 1) if dias_u else 0,
            })

        # eficiencia K7
        eficiencia_k7.append({
            "Mes": mes, "Total Ops": total_ops, "Operadores Activos": operadores_activos,
            "Persona-dias Reales": persona_dias, "Dias Calendario": dias_calendario,
            "Ops/Dia (persona-dia)": round(ops_dia_persona, 1),
            "Ops/Dia (calendario)": round(ops_dia_calend, 1),
            "Brecha": round(ops_dia_persona - ops_dia_calend, 1),
        })

        # tiempo K9
        horas_pago = k1 * MIN_POR_PAGO_LOTE / 60
        horas_cts = k3 * MIN_POR_CTS / 60
        horas_tot = horas_pago + horas_cts
        tiempo_k9.append({
            "Mes": mes, "K1 Pago Lote": k1, "K3 CTS": k3,
            "Horas Pago Lote": round(horas_pago, 1), "Horas CTS": round(horas_cts, 1),
            "Horas Totales Estim.": round(horas_tot, 1),
            "Equiv. Dias (8h)": round(horas_tot / HORAS_JORNADA, 1),
        })

        # extornos K8
        e = ext[ext["mes"] == mes] if not ext.empty else pd.DataFrame()
        if not e.empty:
            tot_ext = e["extorno"].nunique()
            def ev(motivo):
                return e[e["motivo"].str.contains(motivo)]["extorno"].nunique()
            imp = ev("USUARIO")
            extornos_k8.append({
                "Mes": mes, "K1 Pago Lote": k1, "Extornos Totales": tot_ext,
                "Imputables (Error Usuario)": imp,
                "Error Sistema": ev("SISTEMA"), "Error Cliente": ev("CLIENTE"),
                "Tasa Imputable": round(imp / k1, 5) if k1 else 0,
                "Tasa Total": round(tot_ext / k1, 5) if k1 else 0,
                "Monto Extornado (S/)": round(e["monto"].sum(), 2),
            })

    # --- Buzon de requerimientos (todo el periodo del CSV) ---
    buzon_tipo, buzon_asig, buzon_alerta, buzon_total = [], [], None, 0
    if not req.empty:
        def col_req(*nombres):
            """Columna exacta si existe; si no, primera que la contenga."""
            cols = {norm(c): c for c in req.columns}
            for n in nombres:
                if norm(n) in cols:
                    return cols[norm(n)]
            for n in nombres:
                for k, v in cols.items():
                    if norm(n) in k:
                        return v
            return None
        c_sol = col_req("solicitud")
        c_asig = col_req("asignado a", "asignado")
        c_fat = col_req("fecha de atencion")
        total = len(req)
        buzon_total = total
        # Se colapsan espacios multiples ("APERTURA  CUENTA" -> "APERTURA CUENTA")
        # para no duplicar categorias por errores de tipeo en la fuente.
        def limpio(serie):
            return (serie.fillna("").astype(str).str.strip()
                    .str.replace(r"\s+", " ", regex=True)
                    .str.upper().replace("", np.nan).dropna())
        if c_sol:
            vc = limpio(req[c_sol]).value_counts()
            buzon_tipo = [{"Tipo de Solicitud": k, "Cantidad": int(v),
                           "% del Total": round(v / total, 4)} for k, v in vc.items()]
        if c_asig:
            vc = limpio(req[c_asig]).value_counts()
            buzon_asig = [{"Asignado a": k, "Cantidad": int(v),
                           "% del Total": round(v / total, 4)} for k, v in vc.items()]
        if c_fat is not None:
            sin_fecha = req[c_fat].fillna("").str.strip().eq("").sum()
            buzon_alerta = (f"BRECHA: {sin_fecha} de {total} solicitudes SIN fecha de "
                            f"atencion -> el tiempo de respuesta del buzon hoy es invisible.")

    # --- Abono de sueldos/CTS: SOLO Oficina Principal (somos nosotros) ---
    abono = []
    centralizacion = []
    if not abo.empty:
        abo_op = abo[abo["tienda"] == TIENDA_PROPIA]
        for mes in sorted(abo_op["mes"].dropna().unique()):
            a = abo_op[abo_op["mes"] == mes]
            exito = a["estado"].str.contains("EXITOSO").sum()
            abono.append({
                "Mes": mes, "Total Abonos": len(a),
                "Sueldos": int((a["tipo"].str.contains("SUELDO")).sum()),
                "CTS": int((a["tipo"].str.contains("CTS")).sum()),
                "Exitosos": int(exito),
                "% Exitoso": round(exito / len(a), 4) if len(a) else 0,
                "Empleadores": a["empleador"].dropna().nunique(),
            })

        # K6 % Centralizacion CTS = CTS atendido en Oficina Principal (Centrales)
        # / CTS total de TODA la red (todas las agencias).
        cts_all = abo[abo["tipo"].str.contains("CTS")]
        for mes in sorted(cts_all["mes"].dropna().unique()):
            c = cts_all[cts_all["mes"] == mes]
            total_red = len(c)
            propio = (c["tienda"] == TIENDA_PROPIA).sum()
            centralizacion.append({
                "Mes": mes, "CTS Oficina Principal": int(propio),
                "CTS Total Red (todas agencias)": total_red,
                "% Centralizacion CTS": round(propio / total_red, 4) if total_red else 0,
            })

    return dict(resumen=resumen, operadores=operadores, extornos=extornos_k8,
                tiempo=tiempo_k9, eficiencia=eficiencia_k7,
                buzon_tipo=buzon_tipo, buzon_asig=buzon_asig, buzon_alerta=buzon_alerta,
                abono=abono, centralizacion=centralizacion, meses=meses)


# ---------------------------------------------------------------------------
# ESCRITURA DEL EXCEL --------------------------------------------------------
# ---------------------------------------------------------------------------
AZUL   = "1F4E78"
CLARO  = "D9E1F2"
AMBAR  = "FFF2CC"
VERDE  = "E2EFDA"
GRIS   = "F2F2F2"
BORDE  = Border(*(Side(style="thin", color="BFBFBF"),) * 4)

def _titulo(ws, texto, ncols):
    ws.append([texto])
    r = ws.max_row
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=max(ncols, 1))
    c = ws.cell(r, 1)
    c.font = Font(bold=True, size=13, color="FFFFFF")
    c.fill = PatternFill("solid", fgColor=AZUL)
    c.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[r].height = 22


def _tabla(ws, filas, mes_obj=None, pct_cols=(), money_cols=()):
    if not filas:
        ws.append(["(sin datos)"]); return
    cols = list(filas[0].keys())
    ws.append(cols)
    hr = ws.max_row
    for j in range(len(cols)):
        c = ws.cell(hr, j + 1)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=AZUL)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDE
    for fila in filas:
        ws.append([fila.get(c) for c in cols])
        r = ws.max_row
        es_obj = mes_obj and str(fila.get("Mes")) == mes_obj
        for j, cname in enumerate(cols):
            c = ws.cell(r, j + 1)
            c.border = BORDE
            if es_obj:
                c.fill = PatternFill("solid", fgColor=AMBAR)
            elif r % 2 == 0:
                c.fill = PatternFill("solid", fgColor=GRIS)
            if cname in pct_cols and isinstance(fila.get(cname), (int, float)):
                c.number_format = "0.0%"
            elif cname in money_cols and isinstance(fila.get(cname), (int, float)):
                c.number_format = "#,##0.00"
    # ancho de columnas
    for j, cname in enumerate(cols):
        largo = max([len(str(cname))] + [len(str(f.get(cname, ""))) for f in filas])
        ws.column_dimensions[get_column_letter(j + 1)].width = min(max(largo + 2, 10), 40)


def _nota(ws, texto):
    ws.append([texto])
    ws.cell(ws.max_row, 1).font = Font(italic=True, size=9, color="595959")
    ws.append([])


# Paleta para series de graficos (distinta de los colores de tabla: necesita
# buen contraste en barras/lineas). Pago=azul, CTS=verde, alerta=rojo,
# acento/meta=naranja.
COLOR_PAGO   = "1F4E78"
COLOR_CTS    = "2E8B57"
COLOR_ALERTA = "C00000"
COLOR_ACENTO = "ED7D31"
COLOR_GRIS2  = "808080"


def _estilo_serie(serie, color, linea=False, marcador=True):
    serie.graphicalProperties.solidFill = color
    serie.graphicalProperties.line.solidFill = color
    if linea:
        serie.graphicalProperties.line.width = 25000  # EMU ~ 2pt
        serie.smooth = False
        if marcador:
            serie.marker = Marker(symbol="circle", size=6)
            serie.marker.graphicalProperties.solidFill = color
            serie.marker.graphicalProperties.line.solidFill = color


def _eje(axis, texto):
    """Titulo de eje en negrita/legible (el default de openpyxl es muy chico)."""
    axis.title = texto
    try:
        axis.title.tx.rich.p[0].r[0].rPr = CharacterProperties(sz=1100, b=True)
    except (AttributeError, IndexError):
        pass


def _hoja_graficos(wb, K):
    """Hoja 'Graficos': una tabla fuente compacta + graficos nativos de Excel."""
    ws = wb.create_sheet("Graficos")
    ws.sheet_view.showGridLines = False
    _titulo(ws, "GRAFICOS — PAGOS MASIVOS Y CTS", 12)

    resumen = K["resumen"]
    cent_por_mes = {c["Mes"]: c["% Centralizacion CTS"] for c in K["centralizacion"]}

    # ---- Tabla fuente A: una fila por mes ----
    fa_header = ["Mes", "K1 Vol Pago", "K3 Vol CTS", "K2 Monto Pago", "K4 Monto CTS",
                 "K5a %Analista", "K7 %Cobert Pago", "K8 %Cobert CTS", "K5 Ratio Asist.",
                 "Ops/dia persona", "Ops/dia calendario", "K6 %Centralizacion CTS"]
    fa_start = ws.max_row + 1
    ws.append(fa_header)
    for j in range(len(fa_header)):
        ws.cell(fa_start, j + 1).font = Font(bold=True)
    for fila in resumen:
        k5 = fila["K5 Ratio Asist."]
        ws.append([
            fila["Mes"], fila["K1 Vol Pago"], fila["K3 Vol CTS"],
            fila["K2 Monto Pago"], fila["K4 Monto CTS"], fila["K5a % Analista"],
            fila["K7 % Cobert. Pago"], fila["K8 % Cobert. CTS"],
            None if k5 == "N/A" else k5,
            fila["Ops/dia (persona)"], fila["Ops/dia (calend.)"],
            cent_por_mes.get(fila["Mes"]),
        ])
    fa_end = ws.max_row
    for col_pct in (6, 7, 8, 12):
        for r in range(fa_start + 1, fa_end + 1):
            ws.cell(r, col_pct).number_format = "0.0%"
    ws.append([])

    # ---- Tabla fuente B: K8 Extornos ----
    fb_start = ws.max_row + 1
    ws.append(["Mes", "Tasa Imputable", "Tasa Total"])
    for j in range(3):
        ws.cell(fb_start, j + 1).font = Font(bold=True)
    for fila in K["extornos"]:
        ws.append([fila["Mes"], fila["Tasa Imputable"], fila["Tasa Total"]])
    fb_end = ws.max_row
    for r in range(fb_start + 1, fb_end + 1):
        ws.cell(r, 2).number_format = "0.0%"
        ws.cell(r, 3).number_format = "0.0%"
    ws.append([])

    # ---- Tabla fuente C: K9 Tiempo estimado ----
    fc_start = ws.max_row + 1
    ws.append(["Mes", "Horas Pago Lote", "Horas CTS"])
    for j in range(3):
        ws.cell(fc_start, j + 1).font = Font(bold=True)
    for fila in K["tiempo"]:
        ws.append([fila["Mes"], fila["Horas Pago Lote"], fila["Horas CTS"]])
    fc_end = ws.max_row

    for col in range(1, len(fa_header) + 1):
        ws.column_dimensions[get_column_letter(col)].width = 16

    def refs(min_col, max_col, start, end):
        return (Reference(ws, min_col=min_col, max_col=max_col, min_row=start, max_row=end),
                Reference(ws, min_col=1, min_row=start + 1, max_row=end))

    # Tamaño fisico del grafico (cm) y separacion en filas/columnas de grid.
    CHART_H, CHART_W = 8.0, 15.0          # cm
    ALTO_FILA_CM, ANCHO_COL_CM = 0.53, 1.59  # cm por fila / columna default
    ALTO = int((CHART_H + 2.5) / ALTO_FILA_CM) + 1     # filas entre graficos apilados
    ANCHO = int((CHART_W + 2.0) / ANCHO_COL_CM) + 1    # columnas entre graficos en fila

    ANCLA_COL = len(fa_header) + 2  # graficos a la derecha de las tablas
    fila_ancla = 3
    for col in range(ANCLA_COL, ANCLA_COL + 2 * ANCHO + 2):
        ws.column_dimensions[get_column_letter(col)].width = 8.43

    def ubicar(idx):
        fila = fila_ancla + (idx // 2) * ALTO
        col = ANCLA_COL + (idx % 2) * ANCHO
        return f"{get_column_letter(col)}{fila}"

    charts = []

    # 1) Volumen de transacciones por proceso (K1 vs K3)
    c1 = BarChart(); c1.type = "col"; c1.grouping = "clustered"; c1.style = 10
    c1.title = "Volumen de Transacciones por Proceso"
    _eje(c1.y_axis, "N° de transacciones"); _eje(c1.x_axis, "Mes")
    data, cats = refs(2, 3, fa_start, fa_end)
    c1.add_data(data, titles_from_data=True); c1.set_categories(cats)
    _estilo_serie(c1.series[0], COLOR_PAGO); _estilo_serie(c1.series[1], COLOR_CTS)
    charts.append(c1)

    # 2) Monto operado por proceso (K2 vs K4)
    c2 = BarChart(); c2.type = "col"; c2.grouping = "clustered"; c2.style = 10
    c2.title = "Monto Operado por Proceso (S/)"
    _eje(c2.y_axis, "Monto (S/)"); _eje(c2.x_axis, "Mes")
    data, cats = refs(4, 5, fa_start, fa_end)
    c2.add_data(data, titles_from_data=True); c2.set_categories(cats)
    _estilo_serie(c2.series[0], COLOR_PAGO); _estilo_serie(c2.series[1], COLOR_CTS)
    charts.append(c2)

    # 3) K5a % Analista operando
    c3 = LineChart(); c3.style = 12
    c3.title = "K5a: % Transacciones Ejecutadas por el Analista"
    _eje(c3.y_axis, "% del total de ops"); _eje(c3.x_axis, "Mes")
    data, cats = refs(6, 6, fa_start, fa_end)
    c3.add_data(data, titles_from_data=True); c3.set_categories(cats)
    c3.y_axis.numFmt = "0%"
    _estilo_serie(c3.series[0], COLOR_ALERTA, linea=True)
    charts.append(c3)

    # 4) Cobertura % agencias atendidas
    c4 = LineChart(); c4.style = 12
    c4.title = f"Cobertura: % Agencias Atendidas (de {TOTAL_TIENDAS_RED})"
    _eje(c4.y_axis, "% de agencias cubiertas"); _eje(c4.x_axis, "Mes")
    data, cats = refs(7, 8, fa_start, fa_end)
    c4.add_data(data, titles_from_data=True); c4.set_categories(cats)
    c4.y_axis.numFmt = "0%"
    _estilo_serie(c4.series[0], COLOR_PAGO, linea=True)
    _estilo_serie(c4.series[1], COLOR_CTS, linea=True)
    charts.append(c4)

    # 5) Eficiencia Ops/dia (persona vs calendario)
    c5 = LineChart(); c5.style = 12
    c5.title = "Eficiencia: Ops/Operador/Dia (persona-dia vs calendario)"
    _eje(c5.y_axis, "Ops por operador y dia"); _eje(c5.x_axis, "Mes")
    data, cats = refs(10, 11, fa_start, fa_end)
    c5.add_data(data, titles_from_data=True); c5.set_categories(cats)
    _estilo_serie(c5.series[0], COLOR_PAGO, linea=True)
    _estilo_serie(c5.series[1], COLOR_GRIS2, linea=True)
    charts.append(c5)

    # 6) K5 Ratio de carga entre asistentes
    c6 = BarChart(); c6.type = "col"; c6.style = 10
    c6.title = "K5: Ratio de Carga entre Asistentes"
    _eje(c6.y_axis, "Ratio MAX/MIN (meta <= 2.0x)"); _eje(c6.x_axis, "Mes")
    data, cats = refs(9, 9, fa_start, fa_end)
    c6.add_data(data, titles_from_data=True); c6.set_categories(cats)
    _estilo_serie(c6.series[0], COLOR_ACENTO)
    charts.append(c6)

    # 7) K6 % Centralizacion CTS
    c7 = LineChart(); c7.style = 12
    c7.title = "K6: % Centralizacion CTS (Oficina Principal / Red)"
    _eje(c7.y_axis, "% del CTS de la red"); _eje(c7.x_axis, "Mes")
    data, cats = refs(12, 12, fa_start, fa_end)
    c7.add_data(data, titles_from_data=True); c7.set_categories(cats)
    c7.y_axis.numFmt = "0%"
    _estilo_serie(c7.series[0], COLOR_CTS, linea=True)
    charts.append(c7)

    # 8) K8 Tasa de extornos (imputable vs total)
    if fb_end > fb_start:
        c8 = LineChart(); c8.style = 12
        c8.title = "K8: Tasa de Extornos (Imputable vs Total)"
        _eje(c8.y_axis, "% de K1 Pago Lote"); _eje(c8.x_axis, "Mes")
        data, cats = refs(2, 3, fb_start, fb_end)
        c8.add_data(data, titles_from_data=True); c8.set_categories(cats)
        c8.y_axis.numFmt = "0.0%"
        _estilo_serie(c8.series[0], COLOR_ALERTA, linea=True)
        _estilo_serie(c8.series[1], COLOR_GRIS2, linea=True)
        charts.append(c8)

    # 9) K9 Horas estimadas por proceso
    if fc_end > fc_start:
        c9 = BarChart(); c9.type = "col"; c9.grouping = "stacked"; c9.style = 10
        c9.title = "K9: Horas Estimadas por Proceso"
        _eje(c9.y_axis, "Horas estimadas"); _eje(c9.x_axis, "Mes")
        data, cats = refs(2, 3, fc_start, fc_end)
        c9.add_data(data, titles_from_data=True); c9.set_categories(cats)
        _estilo_serie(c9.series[0], COLOR_PAGO); _estilo_serie(c9.series[1], COLOR_CTS)
        charts.append(c9)

    for idx, ch in enumerate(charts):
        ch.height, ch.width = CHART_H, CHART_W
        ch.legend.position = "b"
        ws.add_chart(ch, ubicar(idx))

    _nota(ws, "Graficos nativos de Excel a partir de las tablas de la izquierda. "
              "Se pueden copiar/pegar como imagen o editar el estilo directo en Excel.")
    return ws


def escribir_excel(K, salida, mes_obj):
    wb = openpyxl.Workbook()

    # ---- Hoja 1: Resumen KPIs ----
    ws = wb.active
    ws.title = "Resumen KPIs"
    _titulo(ws, "TABLERO DE KPIs — PAGOS MASIVOS Y CTS", 17)
    _nota(ws, f"Mes destacado: {mes_obj or '—'}  ·  Generado: {dt.datetime.now():%Y-%m-%d %H:%M}")
    _tabla(ws, K["resumen"], mes_obj,
           pct_cols=("K5a % Analista", "K7 % Cobert. Pago", "K8 % Cobert. CTS"),
           money_cols=("K2 Monto Pago", "K4 Monto CTS", "Total Monto"))
    ws.append([])
    _nota(ws, "K1/K2: Pago Lote (Detalle Pago Lote y CTS, Seccion 1). K3/K4: Deposito Lote CTS (Seccion 2).")
    _nota(ws, "K5 = MAX/MIN ops de asistentes. K5a = ops del analista / total. Cobertura = tiendas unicas / tiendas de la red.")

    # ---- Hoja 2: Detalle por Operador ----
    ws = wb.create_sheet("Detalle por Operador")
    _titulo(ws, "DETALLE POR OPERADOR", 11)
    _tabla(ws, K["operadores"], mes_obj,
           money_cols=("Monto Pago", "Monto CTS", "Total Monto"))

    # ---- Hoja 3: Eficiencia K7 ----
    ws = wb.create_sheet("K7 Eficiencia")
    _titulo(ws, "K7 · EFICIENCIA OPERATIVA", 8)
    _nota(ws, "Persona-dia = Total ops / (dias efectivamente trabajados sumados por operador).")
    _nota(ws, "Calendario = Total ops / (operadores activos x dias calendario con actividad).")
    _tabla(ws, K["eficiencia"], mes_obj)

    # ---- Hoja 4: Extornos K8 ----
    ws = wb.create_sheet("K8 Extornos")
    _titulo(ws, "K8 · TASA DE EXTORNOS", 9)
    _nota(ws, "Un extorno = un N° Mov Extorno unico. Imputable = motivo 'ERROR USUARIO'. Tasa = imputables / K1 Pago Lote.")
    _tabla(ws, K["extornos"], mes_obj,
           pct_cols=("Tasa Imputable", "Tasa Total"),
           money_cols=("Monto Extornado (S/)",))

    # ---- Hoja 5: Tiempo K9 ----
    ws = wb.create_sheet("K9 Tiempo Proc")
    _titulo(ws, "K9 · TIEMPO DE PROCESAMIENTO (ESTIMADO)", 7)
    _nota(ws, f"ADVERTENCIA: estimado, NO cronometrado. Supuesto: {MIN_POR_PAGO_LOTE} min/Pago Lote y {MIN_POR_CTS} min/CTS.")
    _tabla(ws, K["tiempo"], mes_obj)

    # ---- Hoja 6: Abono Sueldos (Oficina Principal) ----
    if K["abono"]:
        ws = wb.create_sheet("Abono Oficina Principal")
        _titulo(ws, "ABONO DE SUELDOS Y CTS · OFICINA PRINCIPAL (NOSOTROS)", 7)
        _nota(ws, f"Filtrado a Tienda = '{TIENDA_PROPIA}' — el resto de agencias hacen su propio abono local.")
        _nota(ws, "Unico reporte con columna 'Estado' -> permite medir % exitoso, brecha ausente en los demas reportes.")
        _tabla(ws, K["abono"], mes_obj, pct_cols=("% Exitoso",))

    # ---- Hoja 6b: K6 Centralizacion CTS ----
    if K["centralizacion"]:
        ws = wb.create_sheet("K6 Centralizacion CTS")
        _titulo(ws, "K6 · % CENTRALIZACION CTS", 3)
        _nota(ws, f"Formula: CTS atendido en {TIENDA_PROPIA} (Centrales) / CTS total de TODA la red.")
        _nota(ws, "Mide cuanto del CTS de la red ya se resuelve de forma centralizada en vez de en cada agencia.")
        _tabla(ws, K["centralizacion"], mes_obj, pct_cols=("% Centralizacion CTS",))

    # ---- Hoja 7: Buzon ----
    ws = wb.create_sheet("Buzon Requerimientos")
    _titulo(ws, "BUZON DE REQUERIMIENTOS", 3)
    _nota(ws, "Categorias tal como se escriben en el CSV: variantes de tipeo (ej. 'ABONO CTS' vs "
              "'ABONO C TS') pueden aparecer separadas. Revisar/agrupar manualmente si se requiere.")
    if K["buzon_tipo"]:
        _tabla(ws, K["buzon_tipo"], None, pct_cols=("% del Total",))
        ws.append([])
    if K["buzon_asig"]:
        _tabla(ws, K["buzon_asig"], None, pct_cols=("% del Total",))
        ws.append([])
    if K["buzon_alerta"]:
        ws.append([K["buzon_alerta"]])
        c = ws.cell(ws.max_row, 1)
        c.font = Font(bold=True, color="9C0006")
        c.fill = PatternFill("solid", fgColor="FFC7CE")

    # ---- Hoja 7b: Graficos ----
    _hoja_graficos(wb, K)

    # ---- Hoja 8: Explicacion / Metodologia ----
    ws = wb.create_sheet("Explicacion")
    _titulo(ws, "METODOLOGIA Y EXPLICACION DE CADA KPI", 4)
    explicacion = [
        ("KPI", "Que mide", "Como se calcula", "Fuente"),
        ("K1 Volumen Pago Lote", "Cantidad de transacciones de pago masivo (planillas)",
         "Conteo de filas 'PAGO LOTE' por mes (fecha y hora)", "Detalle Pago Lote y CTS - Seccion 1"),
        ("K2 Monto Pago Lote", "Soles movilizados en pago lote",
         "Suma de 'Monto' de las filas Pago Lote", "Detalle Pago Lote y CTS - Seccion 1"),
        ("K3 Volumen CTS", "Cantidad de depositos lote de CTS",
         "Conteo de filas 'DEPOSITO LOTE CTS' por mes", "Detalle Pago Lote y CTS - Seccion 2"),
        ("K4 Monto CTS", "Soles depositados por CTS",
         "Suma de 'Monto' de las filas CTS", "Detalle Pago Lote y CTS - Seccion 2"),
        ("K5 Ratio carga Asistentes", "Desbalance de carga entre asistentes",
         "MAX(ops asistente) / MIN(ops asistente). Meta <= 2.0x", "Detalle Pago Lote y CTS - Usuario"),
        ("K5a % Analista operando", "Cuanto del volumen lo hace el analista (deberia delegar)",
         "Ops del analista / Total ops. Meta <= 10%", "Detalle Pago Lote y CTS - Usuario"),
        ("K6 % Centralizacion CTS", "Cuanto del CTS de la red ya se resuelve de forma centralizada",
         f"CTS en {TIENDA_PROPIA} / CTS total de toda la red", "Abono de Sueldos y CTS - Tienda"),
        ("K7 Cobertura Pago Lote", "Alcance geografico del pago lote",
         f"Tiendas unicas Pago Lote / {TOTAL_TIENDAS_RED} tiendas de la red", "Detalle Pago Lote y CTS - Tienda"),
        ("K8 Cobertura CTS", "Alcance geografico de CTS",
         f"Tiendas unicas CTS / {TOTAL_TIENDAS_RED} tiendas de la red", "Detalle CTS - Tienda"),
        ("K9 Empleadores CTS", "Cantidad de empresas atendidas en CTS",
         "Conteo de empleadores unicos del mes", "Detalle CTS - Empleador"),
        ("K7 Eficiencia (hoja)", "Productividad por operador",
         "Total ops / persona-dias y / (operadores x dias calendario)", "Detalle Pago Lote y CTS"),
        ("K8 Tasa Extornos (hoja)", "Calidad: errores que se revierten",
         "Extornos imputables (ERROR USUARIO) / K1 Pago Lote", "Extornos"),
        ("K9 Tiempo Proc. (hoja)", "Carga horaria estimada del area",
         f"K1x{MIN_POR_PAGO_LOTE}min + K3x{MIN_POR_CTS}min (ESTIMADO)", "Detalle Pago Lote y CTS"),
        ("Abono Sueldos", "Volumen y exito de abonos",
         "Conteo por mes y % con Estado='EXITOSO'", "Abono de Sueldos y CTS"),
        ("Buzon", "Demanda de requerimientos y su reparto",
         "Conteo por tipo de solicitud y por persona asignada", "REQUERIMIENTOS ...csv"),
    ]
    for i, fila in enumerate(explicacion):
        ws.append(list(fila))
        r = ws.max_row
        for j in range(4):
            c = ws.cell(r, j + 1)
            c.border = BORDE
            c.alignment = Alignment(wrap_text=True, vertical="top")
            if i == 0:
                c.font = Font(bold=True, color="FFFFFF")
                c.fill = PatternFill("solid", fgColor=AZUL)
            elif r % 2 == 0:
                c.fill = PatternFill("solid", fgColor=GRIS)
    for j, w in enumerate((26, 42, 50, 34)):
        ws.column_dimensions[get_column_letter(j + 1)].width = w
    ws.append([])
    _nota(ws, "SUPUESTOS: operadores y roles definidos en ROLES; tiempos de procesamiento son estimados; "
              f"denominador de cobertura = {TOTAL_TIENDAS_RED} tiendas.")
    _nota(ws, "LIMITACIONES CONOCIDAS: no hay motivo de extorno estandarizado; el buzon no siempre "
              "registra fecha de atencion; el tiempo de procesamiento no esta cronometrado.")

    for hoja in wb.worksheets:
        hoja.sheet_view.showGridLines = False
    wb.save(salida)


# ---------------------------------------------------------------------------
# MAIN -----------------------------------------------------------------------
# ---------------------------------------------------------------------------
def main():
    print("=" * 64)
    print("  GENERADOR DE KPIs — PAGOS MASIVOS Y CTS")
    print("=" * 64)

    f_det = buscar_archivo("detalle", "lote")
    f_cts = buscar_archivo("cts", "individual")
    f_ext = buscar_archivo("extorno")
    f_abo = buscar_archivo("abono", "sueldo")
    f_req = buscar_archivo("requerimiento", ext=("csv",))

    for etq, ruta in [("Detalle Pago Lote y CTS", f_det), ("Detalle CTS", f_cts),
                      ("Extornos", f_ext), ("Abono Sueldos", f_abo),
                      ("Requerimientos", f_req)]:
        print(f"  [{'OK ' if ruta else '-- '}] {etq:20s}: {os.path.basename(ruta) if ruta else 'NO ENCONTRADO'}")

    if not f_det:
        print("\n  ERROR: no encuentro 'Detalle de Pago Lote y CTS'. Es obligatorio.")
        return

    print("\n  Leyendo y calculando...")
    det = leer_detalle_pago_lote(f_det)
    cts = leer_detalle_cts(f_cts) if f_cts else pd.DataFrame()
    ext = leer_extornos(f_ext) if f_ext else pd.DataFrame()
    abo = leer_abono(f_abo) if f_abo else pd.DataFrame()
    req = leer_requerimientos(f_req) if f_req else pd.DataFrame()

    K = calcular(det, cts, ext, abo, req)
    if not K["meses"]:
        print("  ERROR: no se detectaron meses con datos."); return

    mes_obj = MES_OBJETIVO or K["meses"][-1]
    salida = os.path.join(CARPETA, f"KPIs_Pagos_CTS_{mes_obj}.xlsx")
    escribir_excel(K, salida, mes_obj)

    print(f"  Meses detectados: {', '.join(K['meses'])}")
    print(f"  Mes destacado   : {mes_obj}")
    fila = next((r for r in K["resumen"] if r["Mes"] == mes_obj), None)
    if fila:
        print("\n  Resumen del mes destacado:")
        print(f"    K1 Vol Pago Lote : {fila['K1 Vol Pago']:>8}   (S/ {fila['K2 Monto Pago']:,.2f})")
        print(f"    K3 Vol CTS       : {fila['K3 Vol CTS']:>8}   (S/ {fila['K4 Monto CTS']:,.2f})")
        print(f"    Total Ops        : {fila['Total Ops']:>8}   (S/ {fila['Total Monto']:,.2f})")
        print(f"    Empleadores CTS  : {fila['K9 Empleadores']:>8}")
    ek = next((e for e in K["extornos"] if e["Mes"] == mes_obj), None)
    if ek:
        print(f"    Extornos totales : {ek['Extornos Totales']:>8}   (imputables: {ek['Imputables (Error Usuario)']})")
    print(f"\n  Excel generado: {os.path.basename(salida)}")
    print("=" * 64)


if __name__ == "__main__":
    main()
