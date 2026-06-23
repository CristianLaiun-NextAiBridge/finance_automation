# Finance Automation — NextAI Bridge

Automatización contable que descarga transacciones desde la API de Mercury, las organiza en un Google Sheet de conciliación y las matchea con comprobantes físicos usando Gemini AI.

---

## Arquitectura

```
API Mercury
    │
    ▼
pestaña "mercury"          ← raw data sin transformar
    │
    ▼
pestaña "Ledger"           ← vista limpia para el equipo contable
    │
    ├── Categoría asignada por IA  (pestaña "setup" como fuente)
    └── Receipt vinculado por IA   (Drive: FOLDER_COMPROBANTES_ID)
```

### Spreadsheet de conciliación
`https://docs.google.com/spreadsheets/d/19-SS1TaCDNcSZvZsTHpIdwudZSzrssbagIHmNl4hj5U`

| Pestaña | Descripción |
|---|---|
| `mercury` | Raw data de la API de Mercury. Se sobreescribe la ventana de los últimos 30 días en cada ejecución. |
| `Ledger` | Hoja de trabajo del equipo. Columnas transformadas, categorías, comentarios y links a comprobantes. |
| `setup` | Lista de categorías contables usadas para la asignación automática por IA. |

---

## Estructura de archivos

| Archivo | Responsabilidad |
|---|---|
| `Code.js` | Constantes, credenciales, helpers compartidos y orquestadores |
| `mercury_sync.js` | Descarga transacciones desde la API de Mercury y calcula el saldo corriente |
| `formatting.js` | Transforma y copia datos de `mercury` al `Ledger`; aplica formato visual |
| `invoice_matching.js` | Matchea cada transacción con su comprobante físico usando Gemini AI |
| `setup_credentials.js` | Carga las credenciales en Script Properties (ejecutar una sola vez) |
| `setup_triggers.js` | Configura los triggers automáticos (ejecutar una sola vez) |

---

## Columnas del Ledger

| Columna | Descripción |
|---|---|
| `Checked` | Checkbox manual. Fila tildada = verificada, ningún proceso la toca. |
| `Date` | Fecha de la transacción (de `Date (UTC)` en Mercury). |
| `Description` | Unificación de Description + Bank Description + Reference + Note de Mercury. |
| `Category` | Categoría asignada por IA desde la lista de la pestaña `setup`. Dropdown editable. |
| `Comments` | Campo libre para anotaciones del equipo. Única columna editable. Fondo celeste. |
| `Amount In (+)` | Créditos (montos positivos). |
| `Amount Out (-)` | Débitos (montos negativos, mostrados como positivos). |
| `Account balance` | Saldo de la cuenta después de cada transacción, calculado hacia atrás desde el saldo real de la API. |
| `Receipt` | Hipervínculo al PDF del comprobante en Drive, asignado por IA. |
| `JSON Data` | Hipervínculo al JSON extraído del comprobante, asignado por IA. |

**Reglas de la hoja:**
- Toda la hoja está protegida excepto la columna `Comments`.
- Filas con `Checked = TRUE` son inmutables: ningún proceso automatizado las modifica.
- Al tildar `Checked`, la fila entera se pinta de verde claro.

---

## Schedule automático

| Hora | Función | Qué hace |
|---|---|---|
| 5:00 AM | `actualizarTablaMercury` | Descarga transacciones de Mercury → pestaña `mercury` |
| 6:00 AM | `procesarLedger` | Formateo + Categorías IA + Matching de receipts |
| 3:00 PM | `actualizarTablaMercury` | Segunda bajada diaria |
| 4:00 PM | `procesarLedger` | Segundo procesamiento completo |

> Apps Script ejecuta cada trigger dentro de una ventana de ~1 hora a partir de la hora indicada.

---

## Setup inicial (primera vez)

### 1. Clonar el repo y linkear con Apps Script

```bash
git clone https://github.com/CristianLaiun-NextAiBridge/finance_automation
cd finance_automation
clasp login
clasp clone 1xDJO2VfIOG9vJ3Ea-hQoWuP_dSD9SmGsYrlbZymAUR457cZnW5V28Jms
```

El hook `pre-push` sincroniza automáticamente con Apps Script en cada `git push`.

### 2. Configurar credenciales

Abrí el proyecto en [Apps Script](https://script.google.com/home/projects/1xDJO2VfIOG9vJ3Ea-hQoWuP_dSD9SmGsYrlbZymAUR457cZnW5V28Jms/edit), completá los valores en `setup_credentials.js` y ejecutá:

```
setupCredentials()
```

Las credenciales se guardan cifradas en Script Properties. Los valores reales nunca van al repo.

| Script Property | Descripción |
|---|---|
| `SA_PROJECT_ID` | `project_id` del JSON de la Service Account (Vertex AI) |
| `SA_CLIENT_EMAIL` | `client_email` del JSON de la Service Account |
| `SA_PRIVATE_KEY` | `private_key` completa del JSON de la Service Account |
| `MERCURY_API_TOKEN` | Token Read-Only de Mercury (Settings → API Tokens) |

Verificá con `verifyCredentials()`.

### 3. Configurar triggers

```
setupTriggers()
```

Crea los 4 triggers automáticos. Verificá con `listTriggers()`.

### 4. Primera carga de datos

```
actualizarTablaMercury()   → carga histórico completo desde 2020-01-01
formatearHoja()            → copia y transforma al Ledger
asignarCategorias()        → categoriza con IA
matchearFacturasConDrive()  → matchea receipts
```

---

## Funciones disponibles

| Función | Cuándo usarla |
|---|---|
| `actualizarTablaMercury()` | Actualizar la raw data desde Mercury |
| `formatearHoja()` | Sincronizar Ledger y aplicar formato |
| `asignarCategorias()` | Asignar/reasignar categorías con IA |
| `matchearFacturasConDrive()` | Buscar y linkear receipts |
| `procesarLedger()` | Ejecutar los tres pasos anteriores en secuencia |
| `setupCredentials()` | Cargar credenciales en Script Properties (una vez) |
| `verifyCredentials()` | Verificar que las credenciales están cargadas |
| `clearCredentials()` | Borrar todas las Script Properties |
| `setupTriggers()` | Configurar los triggers automáticos (una vez) |
| `listTriggers()` | Ver los triggers activos |

---

## Flujo de desarrollo

```bash
# Editar código localmente
git add .
git commit -m "descripción"
git push   # → pre-push hook ejecuta clasp push → Apps Script actualizado
```

> **Importante:** editar siempre localmente. Nunca modificar directamente en el editor web de Apps Script o los cambios no quedan en el repo.
