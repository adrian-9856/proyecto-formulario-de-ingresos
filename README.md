# Formulario de Ingresos — Creamos Guatemala

Script de Google Apps Script que conecta el formulario de KoboToolbox con Google Sheets y envía un resumen semanal cada lunes con los nuevos ingresos.

---

## Cómo funciona

```
KoboToolbox (formulario)
        ↓  (cada hora, silencioso)
  Google Sheets  ←──── nuevas filas marcadas como "Nuevo"
        ↓  (cada lunes 8 AM, solo si hay nuevos)
  Correo resumen con tabla de ingresos de la semana
        ↓
  Filas marcadas como "Enviado" (para no repetir)
```

---

## Configuración paso a paso

### 1. Crear el Google Sheet

1. Ve a [sheets.google.com](https://sheets.google.com) y crea una hoja nueva
2. Nombre sugerido: `Ingresos Creamos`

### 2. Abrir el editor de Apps Script

1. En Google Sheets → **Extensiones → Apps Script**
2. Borra el código por defecto
3. Pega todo el contenido del archivo `Code.gs`

### 3. Obtener tu Token de API de KoboToolbox

1. Entra a [kf.kobotoolbox.org](https://kf.kobotoolbox.org)
2. Clic en tu nombre → **Account Settings**
3. Sección **API** → copia el token

### 4. Editar la configuración

Al inicio del script, edita estas 2 líneas:

```javascript
KOBO_API_TOKEN:     'tu-token-aqui',
NOTIFICATION_EMAIL: 'correo@creamos.org',
```

### 5. Probar la conexión

1. Selecciona la función `testConnection` → **Ejecutar**
2. Acepta los permisos de Google (solo la primera vez)
3. Verifica en el log: `Conexión exitosa. Total de respuestas: XX`

### 6. Probar el correo semanal

1. Asegúrate de tener al menos una fila con estado **"Nuevo"** en el Sheet
2. Selecciona `testWeeklySummary` → **Ejecutar**
3. Revisa que llegue el correo con la tabla de ingresos

### 7. Activar los triggers automáticos

1. Selecciona `setupTrigger` → **Ejecutar**
2. El log confirmará:
   - `checkNewSubmissions`: corre cada hora (sincroniza datos)
   - `sendWeeklySummary`: corre cada lunes a las 8 AM (envía resumen)

---

## Funciones disponibles

| Función | Cuándo ejecutar |
|---|---|
| `testConnection` | Una vez para verificar que la API funciona |
| `testWeeklySummary` | Para probar el correo sin esperar el lunes |
| `setupTrigger` | Una sola vez para activar todo |
| `checkNewSubmissions` | Automático cada hora (no ejecutar manualmente) |
| `sendWeeklySummary` | Automático cada lunes 8 AM (no ejecutar manualmente) |

---

## Columnas del Google Sheet

| Col | Campo | Fuente |
|---|---|---|
| A | ID | `_id` |
| B | Nombre del Visitante | `visitor_info/visitor_name` |
| C | **Correo del Visitante** | `visitor_info/visitor_email` ← campo nuevo en el formulario |
| D | Fecha de Visita | `visitor_info/visit_date` |
| E | Boletín (newsletter) | `visitor_info/newsletter` |
| F | Fecha de Envío | `_submission_time` |
| G | Contacto de Emergencia | `emergency_contact/emergency_name` |
| H | Teléfono de Emergencia | `emergency_contact/emergency_phone` |
| I | Permiso de Fotos | `guidelines_section/photo_permission` |
| J | Exención de Responsabilidad | `liability_section/liability_waiver` |
| K | **Estado** | `Nuevo` → `Enviado` (automático) |

---

## Agregar correo del visitante al formulario XLSForm

En la hoja **survey** del XLSForm, agrega esta fila dentro del grupo `visitor_info`:

| type | name | label::English (en) | label::Español (es) | required |
|---|---|---|---|---|
| `text` | `visitor_email` | `Email address` | `Correo electrónico` | `false` |

El campo en KoboToolbox quedará como `visitor_info/visitor_email`, que es exactamente lo que lee el script.

---

## Datos técnicos

- **Asset UID:** `aJHSDMnJqjhZ6YPUsiyEze`
- **API URL:** `https://kf.kobotoolbox.org/api/v2/assets/aJHSDMnJqjhZ6YPUsiyEze/data/?format=json`
