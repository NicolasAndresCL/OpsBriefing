# 📊 Reporte Operativo LiveOps — PedidosYa

Aplicación Streamlit para generación de reportes operacionales a partir de
exportaciones Excel de Tableau. Persistencia SQLite y exportación de BD.

---

## 🚀 Inicio rápido (Windows)

1. Copia todos los archivos en una misma carpeta.
2. Haz doble clic en **`iniciar.bat`**.  
   El script crea el entorno virtual e instala dependencias automáticamente.
3. Abre `http://localhost:8501` en tu navegador.

### Manual (cualquier SO)
```bash
pip install -r requirements.txt
streamlit run reporte_operativo.py
```

---

## 📂 Estructura del Excel de Tableau

| Parámetro       | Valor           |
|-----------------|-----------------|
| Rango de datos  | `C2:AH38`       |
| Fila encabezados| Fila 2          |
| Sheet           | Primera hoja    |

### Columnas esperadas (nombre flexible, se detecta por coincidencia):

| Columna en Excel          | Clave interna        |
|---------------------------|----------------------|
| Region / City / Zona      | `region`             |
| DT / Delivery Time        | `dt`                 |
| % Late Orders             | `pct_late`           |
| % FR / FR                 | `pct_fr`             |
| UTR                       | `utr`                |
| At Vendor Time            | `at_vendor`          |
| Rider Accepting Time      | `rider_accepting`    |
| Hold Back Time            | `hold_back`          |
| Responsable / Responsible | `responsable`        |

> Si algún nombre de columna no coincide, edita `COL_MAP` al inicio de `reporte_operativo.py`.

---

## 🗺️ Lógica del reporte

1. **Nacional**: primera fila cuya columna `region` contenga TOTAL / CL / NACIONAL / CHILE.
2. **RM**: fila con RM / SANTIAGO / METROPOLITANA.
3. **Top 3 regiones**: resto de filas ordenadas por score de desviación  
   `score = (DT / 33) + (% Late Orders)` — las 3 con mayor score aparecen.

### Diagnóstico automático

El texto de diagnóstico se genera según reglas:

| Condición                   | Mensaje                              |
|-----------------------------|--------------------------------------|
| DT ≥ 33 min                 | Supera umbral de gestión             |
| DT ≥ 30 min                 | Alerta preventiva                    |
| Hold Back ≥ 15 min          | Retención pre-despacho elevada       |
| At Vendor ≥ 10 min          | Latencia en tienda                   |
| Rider Accepting ≥ 10 min    | Flota lenta                          |
| UTR < 1.2                   | Posible exceso de flota              |
| % Late Orders ≥ 25%         | Alta proporción de tardías           |

---

## 💾 Base de datos (SQLite)

Archivo: `reportes_operativos.db` — se crea automáticamente en la misma carpeta.

### Tabla `reportes`
| Campo       | Tipo    | Descripción                        |
|-------------|---------|------------------------------------|
| id          | INTEGER | PK autoincremental                 |
| fecha       | TEXT    | Fecha del reporte (YYYY-MM-DD)     |
| horario     | TEXT    | Breakfast / Lunch / Afternoon / Dinner |
| mensaje     | TEXT    | Mensaje completo generado          |
| datos_json  | TEXT    | KPIs en JSON para reprocesamiento  |
| creado_en   | TEXT    | Timestamp de creación              |

Puedes descargar la BD desde:
- Sidebar → **⬇️ Descargar BD SQLite**
- Tab Historial → **⬇️ Exportar BD**

---

## ✏️ Personalización

| Variable         | Ubicación              | Descripción                     |
|------------------|------------------------|---------------------------------|
| `THRESHOLD_DT`   | Línea ~18              | Umbral de gestión (default 33)  |
| `COL_MAP`        | Línea ~25              | Mapeo de columnas del Excel     |
| `HORARIOS`       | Línea ~14              | Tipos de horario disponibles    |
