# Generador de KPIs y Diapositivas — Pagos Masivos y CTS

Automatización en Python que toma los reportes crudos de un equipo de
backoffice (pagos masivos, depósitos de CTS, extornos, abonos y un buzón de
requerimientos) y en segundos genera:

1. **Un Excel de KPIs** con 8 hojas: resumen ejecutivo, detalle por operador,
   eficiencia, tasa de extornos, tiempo estimado de procesamiento,
   centralización, buzón de requerimientos y gráficos nativos de Excel.
2. **Una presentación de PowerPoint** con las diapositivas de KPIs y gestión
   operativa listas para el informe mensual, con gráficos, tarjetas de
   "qué significa" en lenguaje simple y tabla de mejoras solicitadas.

Este proyecto nació para automatizar un proceso que antes tomaba varias horas
manuales cada mes (cruzar 5 reportes distintos, calcular ~15 indicadores y
armar las diapositivas a mano). Aquí se publica **sin los datos reales** de
la empresa donde se usa: nombres de usuarios, agencias, cifras y hasta el
nombre del equipo/área se reemplazaron por datos ficticios, y la plantilla
corporativa de PowerPoint se sustituyó por una presentación generada desde
cero.

## Cómo probarlo

```bash
pip install -r requirements.txt

# 1) Genera datos ficticios de ejemplo (o usa los tuyos, ver formato abajo)
python generar_datos_demo.py

# 2) Calcula los KPIs a partir de los archivos que haya en la carpeta
python generar_kpis.py

# 3) Arma la presentación con esos mismos KPIs
python generar_diapositivas.py
```

Salidas: `KPIs_Pagos_CTS_<MES>.xlsx` e `Informe_Demo_KPIs.pptx`. En
[`ejemplos/`](ejemplos/) hay una muestra ya generada (con datos ficticios)
para ver el resultado sin ejecutar nada.

## Cómo usarlo con tus propios datos

`generar_kpis.py` detecta automáticamente, por el NOMBRE del archivo, los 5
reportes que necesita (acepta sufijos tipo "(1)", "(2)", copias, etc.):

| Archivo esperado (el nombre debe contener...) | Contenido |
|---|---|
| `detalle` + `lote` (.xlsx) | Pago Lote + Depósito Lote CTS, dos secciones apiladas |
| `cts` + `individual` (.xlsx) | CTS individual por cliente |
| `extorno` (.xlsx) | Reversiones de transacciones |
| `abono` + `sueldo` (.xlsx) | Abono de sueldos y CTS de toda la red de agencias |
| `requerimiento` (.csv, separado por `;`) | Buzón de solicitudes de clientes |

Edita al inicio de `generar_kpis.py`:
- `ROLES`: usuarios de tu equipo y su rol (analista/asistentes).
- `TIENDA_PROPIA`: el nombre de tu oficina/agencia centralizadora.
- `TOTAL_TIENDAS_RED`: cantidad de agencias de tu red (denominador de cobertura).
- `MIN_POR_PAGO_LOTE` / `MIN_POR_CTS`: supuestos de tiempo por transacción.

Y en `generar_diapositivas.py`:
- La paleta de colores (`ROJO`, `NARANJA`, etc.) por la de tu marca.
- `LOGO_PNG` (opcional): ruta a tu logo en PNG.
- La tabla de ejemplo `filas_hc` en `construir_slide_gestion` por tus propias
  solicitudes/mejoras del mes.

## KPIs calculados

| KPI | Qué mide |
|---|---|
| K1 / K2 | Volumen y monto de Pago Lote |
| K3 / K4 | Volumen y monto de depósitos CTS |
| K5 / K5a | Balance de carga entre asistentes / % ejecutado por el analista |
| K6 | % de CTS de la red resuelto de forma centralizada |
| K7 / K8 | Cobertura geográfica (agencias atendidas) de Pago Lote y CTS |
| K9 | Empleadores distintos atendidos en CTS |
| Eficiencia | Transacciones por persona-día vs. por día calendario |
| Extornos | Tasa de reversiones, separando error de usuario/sistema/cliente |
| Tiempo estimado | Horas de trabajo estimadas a partir del volumen |
| Buzón | Demanda de solicitudes por tipo y por persona asignada |

La hoja **Explicación** del Excel generado documenta la fórmula y la fuente
exacta de cada indicador.

## Qué NO incluye este repo

Por confidencialidad, no se publican los reportes reales, la plantilla
corporativa de PowerPoint original, el nombre real del equipo/área, ni los
nombres reales de sus integrantes. Lo que ves aquí es la misma lógica de
cálculo, corriendo sobre datos 100% ficticios generados por
`generar_datos_demo.py`.

## Stack

Python · pandas · numpy · openpyxl (Excel + gráficos nativos) · python-pptx
(PowerPoint)
