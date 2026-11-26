# Automatiza tu Tec: reservación y uso de laboratorios de cómputo

Aplicación web para reservar y usar los laboratorios de cómputo

https://script.google.com/macros/s/AKfycbzHrQ5itITbaOo7FFCgU3D3JesUoRFy0aYJJfT__PvZuWEel7hkYIDMiBEkMd9sfzIp/exec

Proyecto individual de la materia Fundamentos de Ingeniería de Software

- **Autora:** Andrea Dalith Zavala Barbosa
- **Carrera:** Ingeniería en Sistemas Computacionales

## Captura de inicio de sesión
<img width="658" height="1280" alt="WhatsApp Image 2026-11-25 at 8 31 50 PM" src="https://github.com/user-attachments/assets/a57818cc-ecf5-4272-b116-7614ca80208c" />


## ¿Qué problema resuelve?

Antes, apartar una computadora implicaba preguntar en persona, anotarse en papel y esperar. Con esta aplicación cada persona reserva desde su celular o computadora, recibe una computadora asignada y entra al laboratorio con un código QR.

## Funciones principales

**Alumnos y docentes**
- Inicio de sesión con correo y contraseña.
- Consulta de horarios libres por laboratorio y fecha.
- Reservación individual: se elige laboratorio, fecha, hora y duración (1, 2 o 3 horas, siempre en hora completa).
- Se elige la materia o tarea de una lista según el tipo de laboratorio (no se escribe libremente).
- **Computadora asignada automáticamente.** Si la persona ya usó ese laboratorio y su computadora está libre, se le da la misma; si no, se le da otra libre.
- **Código QR de acceso**, único para cada reservación, laboratorio y computadora.
- Modificar o cancelar una reservación con al menos 2 horas de anticipación.
- Reporte de fallas eligiendo el tipo de problema de una lista (con detalle opcional).
- Los docentes reservan el laboratorio completo para su clase; sus alumnos obtienen su propio QR y computadora en la sección "Clase de tu docente".

**Lector de acceso**
- Pantalla para una tablet o computadora en la puerta del laboratorio o junto a una computadora.
- Funciona con un lector USB de códigos QR o con la cámara.
- Muestra "Acceso permitido" o "Acceso denegado" con el motivo.
- Permite salir y volver a entrar dentro del horario reservado.

**Responsable de laboratorio**
- Panel con las computadoras libres por horario y las reservas del día.
- Atención de fallas (Abierto, En revisión, Resuelto) con aviso por correo.
- Envío de reportes de uso al jefe del Centro de Cómputo.

**Jefe del Centro de Cómputo**
- Reporte de uso: reservas, horas usadas, reservas no usadas y fallas.

## Reglas del sistema

| Regla | Valor |
|---|---|
| Horario | 7:00 a 20:30 |
| Inicio de reservación | Solo en hora completa |
| Duración | 1, 2 o 3 horas |
| Anticipación para modificar o cancelar | 2 horas |
| Tolerancia para usar el QR | 15 minutos (si no, el horario se libera) |
| Estados de reservación | Reservada, Usada, No usada, Cancelada |

## Roles

Alumno, Docente, Responsable de laboratorio y Jefe del Centro de Cómputo. Cada rol solo puede hacer lo suyo; los permisos se revisan en el servidor.

## Tecnologías

- **Google Apps Script**: servidor (`Codigo.gs`) y publicación de la aplicación web.
- **Google Sheets**: base de datos (hojas Usuarios, Laboratorios, Equipos, Materias, Reservas, Accesos, Usos, Fallas, Parametros y Correos).
- **HTML, CSS y JavaScript**: pantalla de la aplicación (`Index.html`).
- **Gmail (MailApp)**: confirmaciones, liberaciones y avisos de fallas.
- Generador de códigos QR integrado en la página (no necesita internet externo).

## Seguridad

- Contraseñas guardadas con SHA-256 y sal, nunca en texto plano.
- Bloqueo de 10 minutos tras 5 intentos fallidos.
- Sesión con token temporal de 6 horas.
- Permisos por rol revisados en el servidor en cada acción.
- Control de concurrencia (LockService) para que dos personas no tomen la misma computadora.
- Los códigos QR son aleatorios, de un solo uso por reservación, y no se muestran a otros usuarios.

## Archivos del repositorio

| Archivo | Para qué sirve |
|---|---|
| `Codigo.gs` | Lógica del servidor: reglas, reservaciones, QR, fallas, reportes y correos |
| `Index.html` | Pantallas de la aplicación |
| `appsscript.json` | Configuración del proyecto de Apps Script |


## Limitaciones

- La cuota de correos de Gmail es limitada (100 al día en cuentas personales).
- El lector USB debe ser de códigos 2D para leer un QR desde una pantalla.
- La cámara puede estar bloqueada dentro de la página de Apps Script, según el navegador; en ese caso se usa el lector USB o se escribe el código.
- El acceso físico depende de que el laboratorio tenga un lector instalado.

## Trabajo futuro

- Conectar el lector QR a una cerradura eléctrica en cada laboratorio.
- Usar la credencial del Tec como acceso en lugar de un QR.
- Lectores integrados en cada computadora.
- Notificaciones por mensaje además de correo.
