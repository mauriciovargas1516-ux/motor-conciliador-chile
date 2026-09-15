# Motor Conciliador Bancario e Inteligencia de Cuentas (Python)

Pipeline automatizado diseñado para agilizar la conciliación de flujos de caja corporativos, homologando cartolas bancarias chilenas (BancoEstado, Santander, etc.) con listados de ERPs mediante análisis de datos y lógica difusa.

## Funcionalidades Clave

* **Lectura Inteligente (Anti-BancoEstado):** Sistema de escaneo dinámico que detecta automáticamente las cabeceras reales de los archivos `.xlsx`, ignorando metadatos y formato corporativo basura.
* **Adaptador Universal:** Homologación automática de columnas de diferentes instituciones bancarias y plataformas a un esquema estandarizado mediante un buscador estadístico.
* **Limpieza Forense:** Extracción automática de RUTs chilenos desde glosas complejas y eliminación de *stopwords* transaccionales.
* **Motor de Conciliación en 3 Fases:**
  1. *El Cajero:* Agrupación de pagos consolidados (varios a uno).
  2. *El Francotirador:* Cruce exacto por RUT y monto.
  3. *El Detective (Fuzzy Matching):* Emparejamiento por coincidencia de texto difuso (`thefuzz`) para descripciones imperfectas.
* **Auditoría y Reportes:** Registro en tiempo real mediante archivos `.log` y generación autónoma de reportes en Excel (`Reporte_Excepciones_Gerencia.xlsx`).

## Tecnologías Utilizadas

* **Lenguaje:** Python 3
* **Procesamiento de Datos:** Pandas
* **Interfaz Gráfica:** Tkinter
* **Inteligencia de Texto:** TheFuzz
* **Generación de Reportes:** OpenPyXL
