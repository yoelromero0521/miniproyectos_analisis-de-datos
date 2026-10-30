# ✈️ Dashboard Analítico de Operaciones y Finanzas - Aerolínea KODIGO

## 📌 Descripción del Proyecto
Diseño e implementación de una solución integral de Business Intelligence en **Power BI** para evaluar el desempeño financiero y operativo de una aerolínea. El proyecto consolida datos de vuelos, tripulación, rutas y costos operativos mediante un **modelo dimensional en estrella**, permitiendo analizar rentabilidad, puntualidad y eficiencia de la flota.

---

## 🎯 Objetivo y Problemática
* **Problemática:** La dispersión de registros operativos (vuelos, estados, tripulación) y métricas financieras (costos de combustible, ingresos por carga/pasajeros) dificulta la evaluación consolidada del margen de ganancia y el rendimiento operacional.
* **Objetivo:** Construir un modelo de datos analítico centralizado y un dashboard interactivo que permita monitorear los ingresos totales, la distribución de costos mes a mes y la productividad operacional por aeronave y capitán.

---

## 🏗️ Arquitectura y Modelado de Datos (Star Schema)
El proyecto fue construido bajo una arquitectura de **Modelo Estrella (Star Schema)** conectando tablas de dimensiones relacionales a una tabla de hechos centralizada:

* 📊 **Tabla de Hechos (`Fact_Vuelos`):**
  * Registra métricas clave a nivel de vuelo: `IngresoPasaje`, `IngresoCarga`, `CostoOperativo_USD`, `CombustibleConsumido`, `MinutosRetraso`, `PasajerosPagados` y `CargaKg`.
* 🧩 **Tablas de Dimensiones (`Dim_*`):**
  * `Dim_Aeronave`: Especificaciones de modelo y flota (ej. Airbus A320neo, Boeing 737-800).
  * `Dim_Tripulacion`: Asignación de capitanes y personal de vuelo.
  * `Dim_Ruta`: Orígenes, destinos y distancias.
  * `Dim_EstadoVuelo`: Clasificación operativa (*A tiempo, Cancelado, Desviado, Retrasado*).

---

## 🛠️ Tecnologías y Herramientas
* **Power BI Desktop:** Modelado de datos, creación de relaciones, medidas en DAX y maquetación del tablero.
* **DAX (Data Analysis Expressions):** Cálculo de indicadores financieros, comparativos intermensuales y agregaciones dinámicas.
* **Power Query:** Transformación, limpieza y estandarización del tipo de datos de los archivos fuente CSV.
* **CSV Datasets:** Arquitectura de datos estructurada.

---

## 📊 KPIs e Indicadores Monitoreados
El panel de control visualiza los siguientes indicadores estratégicos:

* 💰 **Métricas Financieras:**
  * **Ingresos Totales:** Consolidación global de venta de pasajes e ingresos por carga.
  * **Análisis de Costos Operativos:** Comparativa visual mensual (*Mes Actual vs. Mes Anterior*).
* ✈️ **Métricas Operativas y Flota:**
  * **Capacidad Operativa:** Cantidad total de aeronaves activas en flota.
  * **Volumen de Vuelos:** Distribución de vuelos por capitán y por modelo de avión.
  * **Tendencia Temporal:** Evolución del volumen de vuelos a lo largo del año.
  * **Segmentación:** Filtro dinámico por estado del vuelo (*A tiempo, Retrasado, Cancelado, Desviado*).

---

## 📂 Contenido del Repositorio
* `ModeloTemaLibre_RomeroElias.pbix`: Archivo interactivo de Power BI con el modelo de datos, DAX y dashboard.
* `Fact_Vuelos-v2.csv`: Tabla de hechos con el histórico de transacciones operativas y financieras.
* `Dim_Aeronave.csv`: Dimensión de modelos y unidades de vuelo.
* `Dim_EstadoVuelo.csv`: Dimensión del estatus de operaciones.
* `Dim_Ruta.csv`: Dimensión de conexiones y trayectos.
* `Dim_Tripulacion.csv`: Dimensión de capitanes y personal a bordo.
* `pbi dashboard.png`: Captura de pantalla del dashboard interactivo.

---

## 📸 Evidencia Visual

![Dashboard de Finanzas y Rendimiento](pbi%20dashboard.png)
*Figura 1: Tablero interactivo en Power
