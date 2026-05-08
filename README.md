# Dashboard de Carteras PH - EMDECOB

Sistema de visualización premium para el análisis de capital, abonos y saldos de carteras de Propiedad Horizontal.

## 🚀 Funcionalidades
- **Procesamiento Automático**: Lee múltiples archivos Excel de carteras.
- **Visualización Interactiva**: Gráficas de evolución, ranking de conjuntos y detalle por unidad.
- **Standalone**: Genera un archivo HTML estático con todos los datos integrados.

## 🛠️ Estructura del Proyecto
- `index.html`: Dashboard final generado (listo para GitHub Pages).
- `assets/`: Recursos visuales del dashboard.
- `scripts/`:
  - `update_dashboard.py`: Script de Python que procesa los Excels y regenera el dashboard.
  - `template.html`: Plantilla base para la regeneración.

## 📦 Uso Local
1. Coloca tus archivos Excel en una carpeta accesible.
2. Configura la ruta `BASE_DIR` en `scripts/update_dashboard.py`.
3. Ejecuta el script de Python para actualizar los datos.
4. Abre `index.html` en tu navegador.

---
© 2026 EMDECOB | Dirección Analítica
