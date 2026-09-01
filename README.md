# Lector de Control de Operaciones — GitHub Pages

Aplicación web 100% del lado del navegador. El usuario carga un Excel y la app detecta automáticamente la tabla, procesa indicadores y genera un dashboard visual.

## Qué hace

- Acepta `.xlsx`, `.xls`, `.xlsm` y `.csv`.
- El nombre del archivo puede cambiar: la lectura no depende del nombre.
- Detecta automáticamente la fila de encabezados entre las primeras filas del archivo.
- Detecta columnas de indicadores tipo Sí/No.
- En columnas detectadas como indicadores:
  - `Sí`, `SI`, `X`, `1`, `true`, `yes`, `ok`, etc. → **Sí**.
  - celda vacía, `No`, `0`, `false`, etc. → **No**.
- En columnas descriptivas, una celda vacía se muestra como **Sin dato**, para no convertir nombres, fechas u observaciones vacías en “No”.
- Si el libro tiene varias hojas con datos, permite cambiar de hoja.
- Genera:
  - KPIs generales.
  - gráfico Sí vs. No.
  - cumplimiento por indicador.
  - comparación por cualquier categoría detectada.
  - filtros y búsqueda global.
  - tabla navegable.
  - lectura rápida automática de mejor y menor indicador.
- Exporta cada gráfico como PNG, JPG o SVG.
- Exporta el dashboard como PNG/JPG.
- Descarga el detalle filtrado como CSV.
- Permite imprimir o guardar como PDF desde el navegador.

## Publicar en GitHub Pages

1. Crea un repositorio nuevo en GitHub.
2. Sube `index.html`, `styles.css`, `app.js` y este `README.md` en la raíz.
3. En el repositorio ve a **Settings → Pages**.
4. En **Build and deployment**, elige **Deploy from a branch**.
5. Selecciona la rama `main` y carpeta `/ (root)`.
6. Guarda. GitHub mostrará la URL pública de la app.

## Uso

1. Abre la página.
2. Arrastra o selecciona el Excel actualizado.
3. La app procesa el archivo automáticamente.
4. Usa filtros y gráficos.
5. Descarga imágenes o el CSV filtrado.

## Privacidad

El archivo se procesa localmente en el navegador. La aplicación no envía el Excel a un servidor propio.

## Librerías

Se cargan por CDN: SheetJS, Plotly.js y html2canvas. Por ello, el navegador necesita conexión a Internet al abrir la app.
