# Dashboard Logística Inversa Genfar — DHL

BI web interactivo que analiza el archivo `SEGUIMIENTO LOGISTICA INVERSA GENFAR 2026.xlsx` y muestra el estado del ciclo de devoluciones por etapa (SAC → Transporte → CEDI → Cliente) junto con el indicador LOPI consolidado.

## Cómo usarlo

1. Abra `index.html` en un navegador moderno (Chrome, Edge, Firefox o Safari). No requiere servidor — basta con doble clic.
2. Haga clic en **Cargar Excel** (botón amarillo arriba a la derecha).
3. Seleccione el archivo `SEGUIMIENTO LOGISTICA INVERSA GENFAR 2026.xlsx`.
4. El tablero se construye automáticamente leyendo la hoja **BASE**.
5. Para actualizar los datos, vuelva a hacer clic en **Recargar** y cargue el archivo nuevamente — todos los visuales se recalculan.

## Estructura de archivos

```
dashboard/
├── index.html    ← página principal
├── styles.css    ← estilos corporativos DHL
├── app.js        ← lectura, consolidación, cálculos y render
└── README.md     ← este archivo
```

## Reglas de negocio implementadas

- **Unidad de análisis:** una devolución = un `N.GUIA` único. El Excel tiene varias filas por guía (una por producto/lote) y el dashboard las consolida.
- **Ciclo cerrado:** existe `FECHA ENTREGA DOCUMENTOS AL CLIENTE`.
- **SLA del ciclo:** 12 días.
- **Estados LOPI:**
  - `CUMPLE` → ciclo cerrado y `ORDER CYCLE TIME ≤ 12`
  - `NO CUMPLE` → ciclo cerrado y `ORDER CYCLE TIME > 12`
  - `PENDIENTE CIERRE DE CICLO` → sin fecha de entrega de documentos
- **LOPI %:** se calcula **solo sobre ciclos cerrados** = cumple / (cumple + no cumple).

## Bloques del dashboard

1. **Encabezado** — branding DHL, contador de filas/devoluciones, timestamp.
2. **Filtros dinámicos** — destinatario, ciudad, zona, estado del ciclo, rango de fechas (se aplican en cascada a todos los visuales).
3. **Ciclo de la devolución** — visual fijo SAC → Transporte → CEDI → Cliente.
4. **Tarjetas KPI por etapa** — días promedio en SAC, Transporte y CEDI; % entrega a tiempo al cliente.
5. **LOPI consolidado** — velocímetro + desglose por estado + totales + cumplimiento por zona de transporte.
6. **Tabla de detalle** — devoluciones consolidadas ordenadas por fecha.

## Procesamiento

- El archivo se lee **localmente en el navegador** usando [SheetJS](https://sheetjs.com/). No se sube a ningún servidor.
- El archivo original **no se modifica**.
- Todos los cálculos se recalculan en tiempo real al cambiar filtros.

## Stack técnico

- HTML5 + CSS3 puros (sin frameworks)
- JavaScript vanilla
- [SheetJS 0.18.5](https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js) desde CDN
- Canvas API para el velocímetro LOPI
- Google Fonts: Archivo + JetBrains Mono
