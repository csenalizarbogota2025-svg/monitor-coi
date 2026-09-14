# Monitor COI – SDM Bogotá

Aplicación Streamlit para analizar el Reporte Consolidado de Obras de Infraestructura (COI) de la Secretaría Distrital de Movilidad.

## Empresas configuradas

- CONSORCIO SEGURVIAL BOGOTÁ — contrato 2024-3651
- CONSORCIO SEÑALIZAR BOGOTÁ 2025 — contrato 2024-3652
- SOCINTER S.A.S. — identificación por contratista

## Esta versión

- Conserva las 16 columnas oficiales del COI.
- Extrae cada fila individual del COI.
- No usa menciones de empresas en OBSERVACIONES para identificar el contratista.
- Usa el No. CONTRATO como validación para Segurvial y Señalizar.
- Permite filtrar por empresa, contratista/nombre, contrato, estado y localidad.
- Incluye dashboard gráfico.
- Exporta Excel.
- Genera PDF por empresa con las páginas originales correspondientes.

## Streamlit

Main file: `app.py`
Requirements: `requirements.txt`
Branch: `main`
