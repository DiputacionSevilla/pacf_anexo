# Contexto de desarrollo - PACF 2026

## Proposito de la aplicacion

Esta aplicacion Streamlit se usa para preparar el Anexo I del Plan Anual de Control Financiero 2026. Su objetivo es cargar los Excels de expedientes de varios ejercicios, normalizar columnas, depurar registros no validos, calcular indicadores de probabilidad e impacto por seccion y generar una matriz final de riesgo.

La aplicacion esta pensada para trabajar normalmente con tres ejercicios historicos: 2023, 2024 y 2025, aunque permite cargar uno, dos o tres ejercicios y configurar sus pesos.

## Archivo principal

- `app_pacf_2026.py`: contiene toda la aplicacion Streamlit, los calculos, la visualizacion y la exportacion a Excel y Word.

Actualmente el proyecto no esta inicializado como repositorio Git en esta carpeta.

## Flujo funcional

1. El usuario carga opcionalmente un Excel de mapeo de secciones.
2. El usuario carga uno o varios Excels anuales de expedientes.
3. La app detecta columnas candidatas y permite mapear manualmente:
   - Numero de expediente FLP.
   - Importe.
   - Informes desfavorables.
   - Informes favorables.
   - Estado.
   - Fase del gasto.
4. Se normaliza cada ejercicio.
5. Se excluyen registros con estados no validos.
6. Se calculan tablas anuales de probabilidad e impacto.
7. Se consolida la probabilidad con media ponderada.
8. Se consolida el impacto con media simple.
9. Se genera una matriz final de riesgo.
10. Se exportan resultados a Excel y el Anexo I a Word.

## Depuracion de registros

Los estados excluidos por defecto son:

- `Borrador`
- `Pendiente visto bueno del area`
- `Anulado`

La comparacion se hace con normalizacion de texto para tolerar acentos y diferencias menores.

## Calculo de probabilidad

La probabilidad se calcula por ejercicio y seccion.

Columnas principales:

- `Id*s`: suma de informes desfavorables de la seccion.
- `Its`: suma de informes totales de la seccion, entendidos como favorables + desfavorables.
- `It`: suma total de informes del ejercicio.
- `P1 (Id*s / Its)`: proporcion de desfavorables dentro de la seccion.
- `P2 (Its / It)`: peso de los informes de la seccion sobre el total del ejercicio.
- `Ps (%)`: probabilidad ponderada.

Formula:

```text
Ps = [(Id*s / Its * 65) + (Its / It * 35)] / 100
```

Tramos de nivel de probabilidad:

```text
Raro:        0 <= Ps < 10
Improbable: 10 <= Ps < 20
Posible:    20 <= Ps < 40
Probable:   40 <= Ps < 80
Esperado:   Ps >= 80
```

## Consolidacion de probabilidad

La app permite configurar tres pesos desde la barra lateral:

- Peso del ejercicio mas antiguo.
- Peso del ejercicio intermedio.
- Peso del ejercicio mas reciente.

Valores por defecto:

```text
20 / 30 / 50
```

Si solo hay un ejercicio cargado, se usa peso 100.

Si hay dos ejercicios, se usan los pesos correspondientes al ejercicio intermedio y al mas reciente, normalizados.

Si todos los pesos suman cero, se restauran pesos razonables por defecto.

## Calculo de impacto

El impacto anual se calcula por seccion con la formula:

```text
Is (%) = Ms / M * 100
```

Donde:

- `M`: importe total de todos los expedientes validos del ejercicio.
- `Ms`: importe de la seccion.

El calculo de `Ms` es configurable desde la barra lateral:

- `Todos los informes validos`: suma todos los importes validos de la seccion.
- `Solo peticiones con 1 o mas desfavorables`: suma solo importes de expedientes con al menos un informe desfavorable.

En ambos modos, `M` se mantiene como el importe global de todos los expedientes validos del ejercicio.

Tramos de nivel de impacto:

```text
Muy bajo:  0 <= Is < 0,1
Bajo:      0,1 <= Is < 2
Medio:     2 <= Is < 10
Alto:      10 <= Is < 25
Muy alto:  Is >= 25
```

Por tanto, una seccion tiene nivel de impacto `Muy alto` cuando su `Is (%)` es igual o superior al 25% del importe total valido del ejercicio.

## Consolidacion de impacto

El impacto final se calcula como media simple de los `Is (%)` de los ejercicios cargados.

La severidad esta preparada en la estructura, pero no se incorpora al calculo final porque no existe un dato historico homogeneo en las hojas de entrada.

## Matriz de riesgo

La matriz cruza:

- `Nivel de probabilidad`
- `Nivel de impacto`

Resultado posible:

- `Bajo`
- `Medio`
- `Alto`

La matriz de equivalencias esta definida en `RISK_MATRIX`.

## Exportaciones

La app genera:

- Excel con tablas intermedias y finales.
- Word con el Anexo I.

Los nombres de hojas Excel se limpian para evitar caracteres invalidos.

## Estado reciente del desarrollo

Ultimos ajustes realizados:

- Se anadio `import re`, necesario para convertir importes en texto.
- Se anadio `unicodedata` para normalizar nombres de columnas de forma mas robusta.
- Se mejoro `leer_excel()` para elegir motor segun extension:
  - `.xls`: `xlrd`.
  - `.xlsx`: `openpyxl`.
  - fallback automatico de pandas.
- Se reforzo la generacion de nombres de hojas Excel, sustituyendo caracteres invalidos.
- Se rellena `Descripcion` y `Orden_mapeo` en la matriz final cuando una seccion entra solo por uno de los lados del cruce.
- Se verifico que `python -m py_compile app_pacf_2026.py` pasa correctamente.
- Se verifico que las dependencias principales estan instaladas: `streamlit`, `docx`, `openpyxl`, `pandas`, `numpy`.

## Criterios de desarrollo

Mantener cambios acotados y trazables. La aplicacion es de uso administrativo y el objetivo principal es que los calculos sean claros, revisables y exportables.

Prioridades:

1. Evitar errores con datos reales de Excel.
2. Mantener formulas y tramos visibles en codigo y documentacion.
3. No cambiar criterios metodologicos sin confirmacion expresa.
4. Mejorar robustez de lectura, normalizacion y exportacion.
5. Conservar salidas Word y Excel como entregables principales.

## Posibles mejoras pendientes

- Crear datos de prueba anonimizados para validar el flujo completo.
- Separar calculos puros de la interfaz Streamlit para poder testear mejor.
- Anadir tests unitarios para:
  - conversion de importes.
  - clasificacion de probabilidad.
  - clasificacion de impacto.
  - consolidaciones.
  - nombres de hojas Excel.
- Revisar la codificacion visual de algunos textos si aparecen caracteres raros en entornos concretos.
- Documentar formalmente la matriz `RISK_MATRIX` en una tabla dentro del Word generado.
