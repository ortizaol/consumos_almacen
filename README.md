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

> **Nota sobre el formato:** SheetJS en su versión gratuita exporta **solo valores, no formato**.
> Las hojas originales conservan todos los datos, pero pierden colores de relleno, anchos de
> columna y estilos al re-exportarse. Si el área de almacén necesita preservar el formato visual
> original (por ejemplo, el marcado en amarillo como control humano), se debe evaluar una
> librería alternativa como **ExcelJS**. La exportación a **CSV** no cambió.

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
3. **INV** (tipo VLOOKUP): primera coincidencia del ítem en el inventario correspondiente
   (`INV001`, o `INV043` si la bodega es 043). Si no existe, `INV = 0`.
4. **Validación**: `VAL = INV − Total general`.
   `OBSERVACION = "NO SALE POR EXISTENCIA"` cuando `VAL < 0`; vacío cuando `VAL ≥ 0`.

### Correcciones de datos aplicadas

- `Bodega` y `U.M.` se normalizan con `.trim()` (vienen con espacios al final).
- `Item` se normaliza con `String(item).trim()` para el cruce (es numérico en ambas hojas).
- Los códigos de bodega se **descubren dinámicamente**; no están hardcodeados.

## Manejo de errores

Se avisa en pantalla, sin fallar en silencio, cuando: falta la hoja `MANUFACTURA` o `INV001`,
no existe la columna `TIPO CONSUMO`, faltan columnas obligatorias, o ninguna fila tiene `CEPP`.
