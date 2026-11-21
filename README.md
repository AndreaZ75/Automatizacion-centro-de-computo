# Automatizacion-centro-de-computo
# Reservación y uso de laboratorios de cómputo

Aplicación web para celular que automatiza la reservación y el uso de los laboratorios de cómputo del Instituto Tecnológico de Matehuala. Proyecto "Automatiza tu Tec" (SCC-1007), empresa BEP · Be Easy Process.

## Link de App Web
https://script.google.com/macros/s/AKfycbzHrQ5itITbaOo7FFCgU3D3JesUoRFy0aYJJfT__PvZuWEel7hkYIDMiBEkMd9sfzIp/exec 

## Qué hace
- Valida al usuario con correo y contraseña contra la tabla de usuarios.
- Muestra horarios libres y permite reservar de 1 a 3 horas, sin empalmes.
- Registra entrada y salida, y libera la reserva si nadie llega en 15 minutos.
- Permite reportar fallas de equipo y darles seguimiento.
- El responsable consulta reservas y envía el reporte al jefe del Centro de Cómputo.

## Archivos
- `Codigo.gs`: servidor (Google Apps Script). Reglas de negocio, correos y acceso a la hoja de datos.
- `Index.html`: interfaz (HTML, CSS y JavaScript).
- `appsscript.json`: configuración del proyecto.

## Tecnología
HTML, CSS y JavaScript (interfaz), Google Apps Script (servidor) y Google Sheets (base de datos).

## Captura de inicio de sesión
<img width="720" height="1600" alt="WhatsApp Image 2026-11-21 at 2 20 00 AM" src="https://github.com/user-attachments/assets/044c39e3-2b9b-4374-8754-19312a9c0871" />


## Cómo se realizó - implementó
1. Crea una hoja de Google y abre Extensiones → Apps Script.
2. Pega `Codigo.gs` e `Index.html` y copia el contenido de `appsscript.json`.
3. Ejecuta la función `configurarHoja` una vez para crear las pestañas y los datos de ejemplo.
4. Implementar → Nueva implementación → Aplicación web (ejecutar como: yo; acceso: cualquier persona).

Las cuentas incluidas en el código son de ejemplo. La base de datos real no se incluye en este repositorio.

Autora: Andrea Dalith Zavala Barbosa.
