# Tablero Excelencia SSOMA · Salfa Montajes

Página de GitHub Pages: https://gestionssoma141-a11y.github.io/excelencia-ssoma-salfa/
Se publica desde la rama `main`, carpeta raíz. `index.html` es la página completa, con los datos incluidos.

## Cuando pidan "actualizar tablero y página web"

1. **Descargar la matriz desde Google Drive**: "Matriz OKR-KPI 2026 - Centros Salfa Montajes.xlsx"
   (id de Drive `1D9tuykf79ocWX-MzrH0ys3UwsxzM2bR3`). Usar siempre la versión más reciente: revisar
   la hora de modificación, porque a veces se guarda dos veces seguidas.
2. **Regenerar la página**:
   `python3 herramientas/actualizar.py "<ruta a la matriz>.xlsx"`
   El script imprime la nota del año y de cada mes por centro. Si dice que falta la nota calculada,
   repetir con `--recalcular` (usa LibreOffice).
3. **Comparar con la versión anterior** de la matriz, si se tiene, y contarle al usuario qué datos
   cambiaron y cómo se movieron las notas (mes y año, por centro).
4. **Publicar**: commit de `index.html` en `main` y push. La página se actualiza en 1–2 minutos.

No editar `index.html` a mano: los cambios de diseño van en `herramientas/plantilla.html`, que
contiene el marcador `__DATA__` donde el script inserta los datos.

## Cómo se calcula

- Criterios, metas y ponderaciones: los de la matriz corporativa SalfaCorp. El tablero no recalcula
  nada; toma los resultados que calcula la matriz en cada pestaña.
- Pestañas por centro: `CC-141`, `CC-146`, `Chimeneas`. Una fila por mes (filas 5 a 16). La columna A
  vale 1 cuando el mes tiene datos; solo esos meses aparecen en el tablero.
- `Global Salfa Montajes` = promedio simple de los centros con datos en cada mes (no ponderado por dotación).
- Nota mensual (columna EL) = promedio de los 4 focos: Alertas Tempranas (AH), Gestión Preventiva (CW),
  Aprendizaje Continuo (DR) y Excelencia en la Gestión (EK). Meta: 95 %.
- La nota del año en el tablero es el promedio de los meses con datos.
- Un indicador "no se evalúa" un mes cuando su resultado no entra en la fórmula del total del foco de
  esa fila (por ejemplo, el simulacro documental en agosto–octubre).
- En el mapa de calor cada celda muestra el valor real; el color indica el logro contra la meta
  (verde ≥ 95 %, ámbar 60–95 %, rojo < 60 %).

## Revisar al actualizar

- **Noviembre no se publica hasta que el usuario confirme la dotación** (CC-141, celda D15). Mientras
  tanto, correr el script con `--excluir-mes=11` e informar lo que trae noviembre en la matriz sin
  publicarlo. Quitar esta regla solo cuando el usuario confirme la dotación.

- En la pestaña CC-141, las celdas AD15 y AQ15 (reportes esperados de LMG Operativa, noviembre) tenían
  números fijos en vez de la fórmula `=F15*4`. Avisar si siguen así cuando se cargue noviembre.
- Las fórmulas de IAP (columnas DH y DM) dan puntaje completo cuando hay reportes sin cerrar o sin
  aprendizaje; es el criterio de la matriz corporativa y no se ha cambiado.
