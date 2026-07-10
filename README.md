# Validación de Consumos CEPP · Almacén

Aplicativo **HTML autónomo de un solo archivo** (`index.html`) para el área de almacén.
Lee un Excel de manufactura, filtra el material **CEPP**, lo cruza contra el inventario
y calcula la validación de existencia. Todo el procesamiento ocurre **en el navegador**:
no hay servidor, no se sube ningún dato y **no requiere Microsoft Excel** para funcionar.

## Uso

1. Abre `index.html` en cualquier navegador moderno (Chrome, Edge, Firefox).
2. Arrastra o selecciona el archivo `.xlsx`.
3. Revisa las tarjetas de resumen y la tabla de validación.
4. Exporta el resultado a **Excel (.xlsx)** o **CSV** con los cálculos ya resueltos.

### Contenido del Excel exportado

El `.xlsx` descargado incluye, en este orden, las siguientes hojas:

| Hoja | Contenido |
|------|-----------|
| `MANUFACTURA` | Hoja original cargada, tal cual. |
| `INV001` | Hoja de inventario original. |
| `INV043` | Hoja de inventario original — **solo si existía** en el archivo de entrada. |
| `CEPP` | Filas de `MANUFACTURA` con `TIPO CONSUMO = CEPP`, con todas las columnas originales. |
| `ANALISIS` | Tabla de análisis (OBSERVACION, Item, ..., columnas por bodega, Total general, INV, VAL). |

> **Formato y colores originales:** la exportación a `.xlsx` usa **ExcelJS**, que **preserva el
> formato de celda original** (colores de relleno, fuentes, bordes y anchos de columna) de las
> hojas `MANUFACTURA`, `INV001` e `INV043`. La hoja `CEPP` se arma copiando las filas CEPP desde
> `MANUFACTURA` **conservando su formato** (incluido el marcado en amarillo usado como control
> humano). La hoja `ANALISIS` resalta en rojo los ítems con `VAL < 0`, igual que en pantalla.
>
> Si ExcelJS no llega a cargar (por ejemplo, sin acceso a su CDN), el aplicativo **recurre
> automáticamente** a un export de respaldo con SheetJS: genera el mismo libro con todos los
> datos pero **sin formato**, y avisa en pantalla. La exportación a **CSV** no cambió.

> Requiere conexión a internet la primera vez para cargar la librería **SheetJS** desde su CDN.

## Estructura esperada del archivo

| Hoja | Obligatoria | Notas |
|------|-------------|-------|
| `MANUFACTURA` | Sí | Encabezados en la fila 1. Columnas localizadas **por nombre**, no por posición. |
| `INV001` | Sí | Inventario bodega 001. Cruce por `Item`; existencia en `Cant. disponible` (8.ª columna). |
| `INV043` | Opcional | Solo si los datos filtrados incluyen bodega 043. Misma estructura que `INV001`. |

Columnas usadas de `MANUFACTURA`: `TIPO CONSUMO`, `Bodega`, `Item`, `Desc. item`, `U.M.`, `SOLICITUD`, `UE_FC`.

## Lógica

1. **Filtrar**: solo filas con `TIPO CONSUMO = CEPP` (sin distinguir mayúsculas, con `trim()`).
2. **Agrupar**: por `Item + Desc. item + U.M. + UE-FC + Bodega`, sumando `SOLICITUD`.
   Las bodegas detectadas se vuelven columnas dinámicas. Se calcula el **Total general** por ítem.
3. **INV** (tipo VLOOKUP): primera coincidencia del ítem en el inventario correspondiente. Si no existe, `INV = 0`.
4. **Validación**: `VAL = INV − Total general`, con **OBSERVACION de tres estados**:
   - `INV ≥ Total general` → OBSERVACION **vacío** (sale completo).
   - `0 < INV < Total general` → **"Sale parcial por existencias"** (resaltado **ámbar**).
   - `INV ≤ 0` → **"NO SALE POR EXISTENCIA"** (resaltado **rojo**).

### Separación de la bodega 043

Cuando los datos CEPP incluyen la bodega **043** (y por tanto la hoja `INV043`), el análisis se
divide en **dos tablas independientes**, lado a lado y separadas por **dos columnas vacías**, tanto
en pantalla como en la hoja `ANALISIS` del Excel:

- **Tabla principal** (izquierda): todas las bodegas **excepto** 043; su `INV` se busca en `INV001`.
- **Tabla 043** (derecha): únicamente la bodega 043; su `INV` se busca en `INV043` (0 si no existe la hoja).

Cada tabla tiene su propio **bloque de tarjetas de resumen** ("Resumen — Bodegas principales" y
"Resumen — Bodega 043") con sus cifras (ítems, SOLICITUD, no salen, salen parcial, bodegas). Cuando
**no** existe la 043, se muestra una sola tabla y un solo bloque de resumen, como antes.

### Correcciones de datos aplicadas

- `Bodega` y `U.M.` se normalizan con `.trim()` (vienen con espacios al final).
- `Item` se normaliza con `String(item).trim()` para el cruce (es numérico en ambas hojas).
- Los códigos de bodega se **descubren dinámicamente**; no están hardcodeados.

## Ajuste a múltiplo de UE_FC

Módulo adicional (sección propia bajo el análisis) que ajusta las cantidades para que la **suma
por grupo** sea un **múltiplo exacto de `UE_FC`** (factor numérico: 1, 2, 2.12, 2.3, 3, 4, 6…),
sin tocar la SOLICITUD original ni la hoja MANUFACTURA. La `U.M.` (m, kg, und) es solo texto y no
se usa como base. Con `UE_FC` fraccionario el total ajustado puede quedar fraccionario (ej.
`38.16 = 18 × 2.12`): es correcto, debe ser múltiplo exacto de `UE_FC`.

- **Columnas** (localizadas por nombre): `CEPP` (clase: `M2`/`Forrado`/`Estructura`/`Elect M2`
  — **distinta** de `TIPO CONSUMO`), `Desc. C.trabajo`, `UE_FC` y `O.P. Número`. **Aviso:** MANUFACTURA
  trae `Desc. C.trabajo` **dos veces** (columnas P y Z); se usa siempre la **primera (P)**.
- **UE_FC del grupo**: el valor numérico más frecuente del grupo; si es inválido/0/vacío se usa `1`.
- **Agrupación**: por `Item` + clase `CEPP`; para las cuatro clases especiales (`m2`, `forrado`,
  `estructura`, `electm2`, comparadas **sin espacios y en minúsculas**, así `ELECT M2` = `ElectM2`)
  se agrega también `Desc. C.trabajo`. Cada fila es una O.P. ajustable (un mismo `O.P. Número` puede repetirse).
- **Algoritmo determinístico** por grupo con `k = UE_FC`: si el total ya es múltiplo de `k`, no se
  ajusta. Si no, se **prefiere subir** al múltiplo completo (`ceilMult`) siempre que sea factible con un
  **tope duro del 130 %** por O.P. (`ceilMult ≤ Σ 1.30·SOLICITUD`); en caso contrario se baja a `floorMult`.
  Subir reparte proporcional a la SOLICITUD y topa cada O.P. a `1.30×` (redistribuyendo el excedente entre
  las que tengan margen); bajar elimina a 0 las O.P. más pequeñas y reparte el residuo. Ninguna O.P. queda
  negativa ni supera el 130 %, y la suma iguala exactamente el múltiplo.
- **Salida**: tarjetas de resumen y un panel por grupo con `UE-FC`, `Desc. C.trabajo` y el detalle de cada
  O.P. (original, ajustada, delta, % impacto); O.P. eliminadas en **rojo** y modificadas en **ámbar**.
- **Excel**: hoja **`AJUSTE_UM`** (con columnas `UE-FC` y `Desc. C.trabajo`) añadida por los dos caminos de
  exportación (ExcelJS con formato y SheetJS de respaldo), después de `ANALISIS`.

## Manejo de errores

Se avisa en pantalla, sin fallar en silencio, cuando: falta la hoja `MANUFACTURA` o `INV001`,
no existe la columna `TIPO CONSUMO`, faltan columnas obligatorias (incluidas `CEPP` de clase,
`Desc. C.trabajo`, `UE_FC` y `U.M.` que requiere el módulo de ajuste), o ninguna fila tiene `CEPP`.
