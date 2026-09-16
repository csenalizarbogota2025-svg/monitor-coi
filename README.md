# Monitor COI – SDM Bogotá v8

Aplicación Streamlit para cargar, analizar y consultar el COI completo de la Secretaría Distrital de Movilidad.

## Flujo
1. Cargar PDF.
2. Pulsar **ANALIZAR COI**.
3. Seleccionar una o varias empresas, contratos, estados, localidades o secciones.
4. Pulsar **GENERAR LISTA SELECCIONADA**.
5. Consultar dashboard, trazabilidad y descargar Excel/PDF.

La v8 obtiene empresas y contratos directamente del COI cargado y no obliga a seleccionar una lista fija de empresas.

La extracción usa los límites reales de las celdas del PDF y valida que CIV INICIO/CIV FIN sean numéricos, evitando que nombres de ingenieros o direcciones se mezclen con los CIV.
