# Formulario de Ingresos — Creamos Guatemala

Script de Google Apps Script que conecta el formulario de KoboToolbox con Google Sheets y envía notificaciones por correo cuando alguien llena el formulario.

---

## Cómo configurarlo (paso a paso)

### 1. Crear el Google Sheet

1. Ve a [sheets.google.com](https://sheets.google.com) y crea una nueva hoja
2. Ponle un nombre descriptivo, por ejemplo: `Ingresos Creamos`

### 2. Abrir el editor de Apps Script

1. En Google Sheets, ve a **Extensiones → Apps Script**
2. Borra el código que aparece por defecto
3. Copia y pega todo el contenido del archivo `Code.gs` de este repositorio

### 3. Obtener tu Token de API de KoboToolbox

1. Entra a [kf.kobotoolbox.org](https://kf.kobotoolbox.org)
2. Haz clic en tu nombre de usuario (arriba a la derecha) → **Account Settings**
3. En la sección **API**, copia tu token

### 4. Editar los valores de configuración

Al inicio del script, edita estas 2 líneas:

```javascript
KOBO_API_TOKEN:     'PON_AQUI_TU_TOKEN_DE_KOBO',
NOTIFICATION_EMAIL: 'correo@creamos.org',
```

El `KOBO_ASSET_UID` ya está configurado con el UID correcto del formulario.

### 5. Guardar y probar la conexión

1. Guarda el script con **Ctrl+S**
2. En el menú de funciones (arriba), selecciona `testConnection`
3. Haz clic en **Ejecutar**
4. Acepta los permisos que Google solicita (es normal la primera vez)
5. Revisa el **Log** (Ver → Registros) — debe decir "Conexión exitosa"

### 6. Activar el trigger automático

1. Selecciona la función `setupTrigger`
2. Haz clic en **Ejecutar**
3. Listo — el script revisará KoboToolbox **cada hora** y enviará correos de las respuestas nuevas

---

## Qué hace el script

| Función | Descripción |
|---|---|
| `checkNewSubmissions` | Función principal: revisa nuevas respuestas y envía correos |
| `testConnection` | Prueba que la conexión con KoboToolbox funcione |
| `setupTrigger` | Activa el trigger automático cada hora |

## Datos que guarda en el Sheet

| Columna | Campo KoboToolbox |
|---|---|
| ID | `_id` |
| Nombre del Visitante | `visitor_info/visitor_name` |
| Fecha de Visita | `visitor_info/visit_date` |
| Fecha de Envío | `_submission_time` |
| Boletín | `visitor_info/newsletter` |
| Contacto de Emergencia | `emergency_contact/emergency_name` |
| Teléfono de Emergencia | `emergency_contact/emergency_phone` |
| Permiso de Fotos | `guidelines_section/photo_permission` |
| Exención de Responsabilidad | `liability_section/liability_waiver` |

---

## Formulario KoboToolbox

- **Asset UID:** `aJHSDMnJqjhZ6YPUsiyEze`
- **API URL:** `https://kf.kobotoolbox.org/api/v2/assets/aJHSDMnJqjhZ6YPUsiyEze/data/?format=json`
