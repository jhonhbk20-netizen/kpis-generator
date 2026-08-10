# -*- coding: utf-8 -*-
"""
============================================================================
 GENERADOR DE DATOS DE DEMOSTRACION (ficticios)
============================================================================
Crea, en esta misma carpeta, una version FICTICIA de cada reporte crudo que
espera "generar_kpis.py":
    - Detalle de Pago Lote y CTS DEMO.xlsx
    - Detalle CTS Individual DEMO.xlsx
    - Extornos DEMO.xlsx
    - Abono de Sueldos y CTS DEMO.xlsx
    - REQUERIMIENTOS BACKOFFICE DEMO.csv

Los nombres de columnas y la estructura (dos secciones apiladas, encabezado
en fila 3 del CSV, etc.) imitan el formato real de un core bancario, pero
TODOS los valores (usuarios, empresas, montos, fechas) son inventados con un
generador de numeros aleatorios de semilla fija — sirven solo para probar
"generar_kpis.py" y "generar_diapositivas.py" de punta a punta.

Uso:  python generar_datos_demo.py
============================================================================
"""
import os
import csv
import random
import datetime as dt

import openpyxl
from openpyxl.styles import Font

CARPETA = os.path.dirname(os.path.abspath(__file__))
random.seed(42)

MESES = ["2026-04", "2026-05", "2026-06", "2026-07"]
DIAS_POR_MES = {"2026-04": 30, "2026-05": 31, "2026-06": 30, "2026-07": 31}

USUARIOS_EQUIPO = ["ANALISTA1", "ASISTENTE1", "ASISTENTE2"]
USUARIOS_APOYO = ["APOYO01", "APOYO02"]
AGENCIAS = [f"AGENCIA {i:02d}" for i in range(1, 39)] + ["OFICINA PRINCIPAL"]
EMPRESAS = [f"EMPRESA DEMO {n} SAC" for n in
            ["ALFA", "BETA", "GAMMA", "DELTA", "OMEGA", "ZETA", "NORTE", "SUR"]]


def fecha_random(mes):
    anio, m = map(int, mes.split("-"))
    dia = random.randint(1, DIAS_POR_MES[mes])
    hora = random.randint(8, 17)
    minuto = random.randint(0, 59)
    return dt.datetime(anio, m, dia, hora, minuto)


def escribir_hoja(ruta, secciones):
    """secciones: lista de (encabezados, filas) que se escriben una tras otra."""
    wb = openpyxl.Workbook()
    ws = wb.active
    for encabezados, filas in secciones:
        ws.append(encabezados)
        for c in ws[ws.max_row]:
            c.font = Font(bold=True)
        for fila in filas:
            ws.append(fila)
        ws.append([])  # fila en blanco entre secciones
    wb.save(ruta)


# ---------------------------------------------------------------------------
# 1) Detalle de Pago Lote y CTS (Pago Lote + Deposito Lote CTS)
# ---------------------------------------------------------------------------
def generar_detalle_pago_lote():
    encabezados = ["Numero de Movimiento", "Tipo de Operacion", "Monto",
                   "Fecha y Hora", "Usuario", "Tienda", "Empresa"]
    mov = 100000

    def filas_pago():
        nonlocal mov
        filas = []
        base = {"2026-04": 180, "2026-05": 210, "2026-06": 260, "2026-07": 300}
        for mes in MESES:
            n = base[mes] + random.randint(-15, 15)
            for _ in range(n):
                mov += 1
                usuario = random.choices(USUARIOS_EQUIPO + USUARIOS_APOYO,
                                          weights=[10, 45, 40, 3, 2])[0]
                filas.append([mov, "PAGO LOTE", round(random.uniform(800, 15000), 2),
                              fecha_random(mes), usuario,
                              random.choice(AGENCIAS[:25]), random.choice(EMPRESAS)])
        return filas

    def filas_cts():
        nonlocal mov
        filas = []
        base = {"2026-04": 90, "2026-05": 140, "2026-06": 100, "2026-07": 95}
        for mes in MESES:
            n = base[mes] + random.randint(-10, 10)
            for _ in range(n):
                mov += 1
                usuario = random.choices(USUARIOS_EQUIPO + USUARIOS_APOYO,
                                          weights=[8, 46, 44, 1, 1])[0]
                filas.append([mov, "DEPOSITO LOTE CTS", round(random.uniform(1500, 25000), 2),
                              fecha_random(mes), usuario,
                              random.choice(AGENCIAS[:18]), random.choice(EMPRESAS)])
        return filas

    ruta = os.path.join(CARPETA, "Detalle de Pago Lote y CTS DEMO.xlsx")
    escribir_hoja(ruta, [(encabezados, filas_pago()), (encabezados, filas_cts())])
    return ruta


# ---------------------------------------------------------------------------
# 2) Detalle CTS Individual (por cliente)
# ---------------------------------------------------------------------------
def generar_detalle_cts():
    encabezados = ["N° de Movimiento", "Fecha", "Tienda", "Empresa",
                   "Tipo Operacion", "Monto"]
    filas = []
    mov = 500000
    tipos = ["APERTURA", "DEPOSITO", "CANCELACION"]
    for mes in MESES:
        n = random.randint(60, 100)
        for _ in range(n):
            mov += 1
            filas.append([mov, fecha_random(mes), random.choice(AGENCIAS[:20]),
                          random.choice(EMPRESAS),
                          random.choices(tipos, weights=[2, 7, 1])[0],
                          round(random.uniform(1000, 20000), 2)])
    ruta = os.path.join(CARPETA, "Detalle CTS Individual DEMO.xlsx")
    escribir_hoja(ruta, [(encabezados, filas)])
    return ruta


# ---------------------------------------------------------------------------
# 3) Extornos
# ---------------------------------------------------------------------------
def generar_extornos():
    encabezados = ["N° Mov Extorno", "Fecha Extorno", "Monto", "Motivo"]
    filas = []
    ext = 900000
    motivos = ["ERROR USUARIO", "ERROR SISTEMA", "ERROR CLIENTE"]
    for mes in MESES:
        n = random.randint(2, 6)
        for _ in range(n):
            ext += 1
            filas.append([ext, fecha_random(mes), round(random.uniform(500, 8000), 2),
                          random.choices(motivos, weights=[5, 3, 2])[0]])
    ruta = os.path.join(CARPETA, "Extornos DEMO.xlsx")
    escribir_hoja(ruta, [(encabezados, filas)])
    return ruta


# ---------------------------------------------------------------------------
# 4) Abono de Sueldos y CTS (toda la red)
# ---------------------------------------------------------------------------
def generar_abono():
    encabezados = ["Fecha Operacion", "Tipo Operacion", "Estado",
                   "Nombre Empleador", "Canal", "Tienda Operacion de Abono"]
    filas = []
    estados = ["EXITOSO", "OBSERVADO", "RECHAZADO"]
    for mes in MESES:
        # Sueldos y CTS resueltos en Oficina Principal (equipo centralizado)
        for _ in range(random.randint(120, 180)):
            filas.append([fecha_random(mes), random.choice(["ABONO SUELDO", "ABONO CTS"]),
                          random.choices(estados, weights=[9, 1, 0.3])[0],
                          random.choice(EMPRESAS), "LOTE", "OFICINA PRINCIPAL"])
        # CTS resuelto de forma local en el resto de agencias de la red
        for _ in range(random.randint(300, 420)):
            filas.append([fecha_random(mes), "ABONO CTS",
                          random.choices(estados, weights=[9, 1, 0.3])[0],
                          random.choice(EMPRESAS), "LOTE", random.choice(AGENCIAS[:38])])
    ruta = os.path.join(CARPETA, "Abono de Sueldos y CTS DEMO.xlsx")
    escribir_hoja(ruta, [(encabezados, filas)])
    return ruta


# ---------------------------------------------------------------------------
# 5) Buzon de Requerimientos (CSV ';' con encabezado en fila 3)
# ---------------------------------------------------------------------------
def generar_requerimientos():
    ruta = os.path.join(CARPETA, "REQUERIMIENTOS BACKOFFICE DEMO.csv")
    tipos_solicitud = ["APERTURA CUENTA", "ABONO CTS", "CONSULTA SALDO",
                       "RECLAMO", "ACTUALIZACION DATOS", "OTROS"]
    asignados = USUARIOS_EQUIPO
    with open(ruta, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh, delimiter=";")
        w.writerow(["REQUERIMIENTOS BACKOFFICE (DATOS DE EJEMPLO)"])
        w.writerow([])
        w.writerow(["N°", "Fecha de Solicitud", "Tipo de Solicitud", "Asignado a",
                    "Fecha de Atencion", "Estado"])
        n = 1
        for mes in MESES:
            for _ in range(random.randint(30, 45)):
                f_sol = fecha_random(mes).date()
                tiene_fecha_atencion = random.random() > 0.35
                f_at = (f_sol + dt.timedelta(days=random.randint(0, 5))) if tiene_fecha_atencion else ""
                w.writerow([n, f_sol.strftime("%d/%m/%Y"),
                           random.choice(tipos_solicitud), random.choice(asignados),
                           f_at.strftime("%d/%m/%Y") if f_at else "",
                           "ATENDIDO" if tiene_fecha_atencion else "PENDIENTE"])
                n += 1
    return ruta


def main():
    print("Generando datos de demostracion (ficticios) en:", CARPETA)
    rutas = [generar_detalle_pago_lote(), generar_detalle_cts(), generar_extornos(),
             generar_abono(), generar_requerimientos()]
    for r in rutas:
        print("  ->", os.path.basename(r))
    print("\nListo. Ahora puedes correr:  python generar_kpis.py")


if __name__ == "__main__":
    main()
