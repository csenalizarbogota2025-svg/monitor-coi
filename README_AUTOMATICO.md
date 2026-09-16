# Monitor COI 12.0 · Automatización semanal

## Qué hace
Cada viernes el workflow de GitHub Actions:
1. Consulta la página oficial de PMT de la Secretaría Distrital de Movilidad.
2. Localiza el COI más reciente y verifica número y fecha dentro del PDF.
3. Descarga el PDF original.
4. Extrae todos los registros con el mismo motor del Monitor COI.
5. Genera un Excel completo.
6. Genera un PDF resumen.
7. Envía por Gmail el PDF original + Excel + PDF resumen.
8. Guarda `automatic/status.json` con el resultado para mostrarlo en la pestaña Automático de la app.

No compara con el COI anterior.

## Secretos de GitHub
En el repositorio: Settings → Secrets and variables → Actions → New repository secret.

Crear:
- `GMAIL_USER`: correo Gmail/Google Workspace que enviará los informes.
- `GMAIL_APP_PASSWORD`: contraseña de aplicación de Google.
- `MAIL_TO`: uno o varios correos separados por coma.
- `MAIL_CC`: opcional, uno o varios correos separados por coma.

Google indica que las contraseñas de aplicación requieren Verificación en dos pasos. Para SMTP, Gmail usa `smtp.gmail.com`; el script usa SSL en el puerto 465.

## Prueba inmediata
En GitHub: Actions → Monitor COI semanal → Run workflow.

La ejecución manual permite probar el flujo sin esperar al viernes.

## Programación
Viernes a las 18:00 UTC (13:00 hora de Bogotá). GitHub puede presentar algo de retraso en la ejecución programada.
