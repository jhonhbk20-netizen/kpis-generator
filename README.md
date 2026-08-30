# Generador de KPIs y Diapositivas para Backoffice

Automatización en Python que toma los reportes crudos de un equipo de
backoffice (dos procesos operativos en lote, reversiones, confirmaciones y un
buzón de requerimientos) y en segundos genera:

1. **Un Excel de KPIs**: una hoja de **Cálculos** con todas las tablas
   (resumen ejecutivo con fórmulas de Excel, detalle por operador,
   eficiencia, tasa de reversiones, altas y confirmaciones, centralización,
   buzón de requerimientos), una hoja de **Gráficos** nativos y una de
   **Explicación** que documenta la fórmula y la fuente de cada KPI.
2. **Una presentación de PowerPoint** con las diapositivas de KPIs y gestión
   operativa listas para el informe mensual, con gráficos, tarjetas de
   "qué significa" en lenguaje simple y tabla de mejoras solicitadas.

Este proyecto nació para automatizar un proceso que antes tomaba varias horas
manuales cada mes (cruzar reportes distintos, calcular indicadores y armar
las diapositivas a mano), y después siguió creciendo a medida que aparecían
casos reales que un cálculo ingenuo no resolvía bien: meses piloto o
atípicos que no debían compararse contra el resto, personal de apoyo
rotativo que distorsionaba la productividad del equipo fijo, o la diferencia
entre "cuántas sucursales ya se sumaron" y "cuánto de lo que ya está dentro
del alcance realmente lo ejecuta el equipo". Ese historial de decisiones de
diseño es la parte más interesante del repo — se explica en detalle en
[**Metodología**](#metodología) más abajo.

Aquí se publica **sin los datos reales** de la empresa donde se usa: nombres
de usuarios, sucursales, cifras y hasta el nombre del equipo/área se
reemplazaron por datos ficticios (con una historia coherente, ver
[`generar_datos_demo.py`](generar_datos_demo.py)), y la plantilla
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

`generar_kpis.py` detecta automáticamente, por el NOMBRE del archivo, los
reportes que necesita (acepta sufijos tipo "(1)", "(2)", copias, etc.). Los
nombres de archivo esperados vienen del sistema origen de este proyecto;
`buscar_archivo()` y `buscar_detalle()` en el código son el punto de partida
si tu propio sistema exporta con otros nombres — la lógica de detección
(por palabras clave, tolerante a sufijos) es reutilizable tal cual.

Edita al inicio de `generar_kpis.py`:
- `ROLES`: usuarios de tu equipo y su rol (analista/asistentes).
- `TIENDA_PROPIA`: el nombre de tu oficina/sucursal centralizadora.
- `TOTAL_TIENDAS_RED`: cantidad de sucursales de tu red (denominador de una de las coberturas).
- `TIENDAS_CONVENIO`: sucursales habilitadas para el primer proceso operativo (el que requiere un acuerdo previo); `None` para auto-detectar de la data.
- `FUERA_UNIVERSO_POR_FAMILIA`: sucursales que operan una familia sin pertenecer a su universo (anomalías a excluir).
- Los supuestos de tiempo por transacción, uno por cada proceso.

Y en `generar_diapositivas.py`:
- La paleta de colores (`ROJO`, `NARANJA`, etc.) por la de tu marca.
- `LOGO_PNG` (opcional): ruta a tu logo en PNG.
- La tabla de ejemplo `filas_hc` en `construir_slide_gestion` por tus propias
  solicitudes/mejoras del mes.

## KPIs calculados

| KPI | Qué mide |
|---|---|
| K1 / K2 | Volumen y monto del primer proceso en lote |
| K3 / K4 | Volumen y monto del segundo proceso en lote |
| K5 / K5a | Balance de carga entre asistentes / % ejecutado por el analista |
| K6a Eficacia | De lo que ya está dentro del alcance, cuánto lo ejecuta el equipo (no cobertura) |
| K6b Cobertura | Cuántas sucursales del universo real ya se incorporaron al proceso |
| K6c Fuga | Operaciones que una sucursal ya incorporada sigue haciendo por su cuenta |
| K7 Eficiencia | Ops/persona-día (incluye apoyo rotativo) vs. ritmo del mes pico de referencia |
| K8 Reversiones | Tasa de reversiones sobre el TOTAL de operaciones, separando error de usuario/sistema/cliente |
| K9 Altas y Confirmaciones | Volumen de altas y confirmaciones que ejecuta el equipo |
| Contrapartes | Entidades distintas atendidas en el segundo proceso |
| Tiempo estimado | Horas de trabajo estimadas a partir del volumen (auxiliar, no numerado) |
| Buzón | Demanda de solicitudes por tipo y por persona asignada |

La hoja **Explicación** del Excel generado documenta la fórmula y la fuente
exacta de cada indicador. Los nombres que aparecen en esa hoja (y en las
columnas del Excel) siguen la nomenclatura del sistema origen; el propósito
de cada uno está explicado en lenguaje simple ahí mismo.

## Metodología

Este proyecto pasó por dos versiones. La primera calculaba los indicadores
de forma directa; la segunda corrige varios supuestos que sonaban razonables
pero no sobrevivían al contacto con datos reales de varios meses seguidos.
Los datos ficticios de `generar_datos_demo.py` reproducen a propósito estos
mismos casos, para que se puedan ver funcionando sin exponer información real:

- **No todos los meses son comparables.** Un mes piloto (pocos días de
  actividad) o un mes con sobretiempo (más días trabajados que días hábiles,
  ej. una avalancha estacional) no pueden ser la vara con la que se mide la
  capacidad del equipo — `clasificar_meses()` los marca como PARCIAL/ATIPICO
  y los excluye del "mes pico de referencia" de K7. Un mes con el reporte
  cortado a mitad de camino (EN CURSO) tampoco se congela en el histórico.
- **El personal de apoyo rotativo distorsiona la productividad si no se
  cuenta bien.** Si el numerador de un indicador de productividad cuenta las
  operaciones de todos los que trabajaron, pero el denominador (persona-días)
  solo cuenta a la plantilla fija, el trabajo del apoyo se le atribuye a la
  plantilla — inflando su ritmo de forma artificial. La corrección es incluir
  a todos los que efectivamente trabajaron en ambos lados del cálculo.
- **"Cuántas sucursales ya se sumaron" y "cuánto de eso lo hace el equipo" son
  dos preguntas distintas.** Medirlas con el mismo número (K6 original)
  esconde la que realmente importa: con cobertura al 100%, lo que queda vivo
  como indicador es la fuga residual (K6c) — quién sigue procesando por su
  cuenta pese a estar ya incorporado.
- **El universo de un proceso no siempre es "toda la red".** Un proceso que
  requiere un acuerdo previo solo aplica a las sucursales que lo tienen;
  medir su cobertura contra el total de sucursales da un número
  artificialmente bajo y esconde que, dentro de su universo real, el
  proceso puede estar completo.
- **Los totales quedan como fórmulas de Excel, no como valores pegados.**
  Así se puede verificar en el propio libro que un total efectivamente
  cuadra con la suma de sus componentes.

## Qué NO incluye este repo

Por confidencialidad, no se publican los reportes reales, la plantilla
corporativa de PowerPoint original, el nombre real del equipo/área, ni los
nombres reales de sus integrantes. Lo que ves aquí es la misma lógica de
cálculo, corriendo sobre datos 100% ficticios generados por
`generar_datos_demo.py`. Algunos nombres de archivo, variables y columnas
que aparecen en el código y en el Excel generado siguen la nomenclatura
específica del sistema de origen (un core bancario) — no fue posible
generalizarlos sin reescribir la lógica de lectura, así que se mantuvieron
tal cual para no arriesgar el comportamiento ya verificado.

## Stack

Python · pandas · numpy · openpyxl (Excel + gráficos nativos) · python-pptx
(PowerPoint)

## Licencia

MIT — ver [LICENSE](LICENSE).
