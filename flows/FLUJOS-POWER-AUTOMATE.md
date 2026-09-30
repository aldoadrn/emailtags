# Flujos de Power Automate — EmailTags HITSS

El nuevo Outlook **no permite que un add-in mueva correos** ni que ponga categorías mientras redactas. Eso lo hacen estos 2 flujos, usando solo el conector estándar **Office 365 Outlook → "Enviar una solicitud HTTP" (Send an HTTP request)**. No hay Excel, SharePoint ni bases externas: la decisión viaja dentro del propio correo (propiedad invisible del add-in) y el aprendizaje vive en tu buzón.

| Flujo | Qué hace | Tiempo |
|---|---|---|
| **EmailTags · Archivar** | Busca correos con la categoría `ET-Pendiente`, aplica categorías/bandera, mueve el correo (y todo el hilo si así lo elegiste) y quita la marca. | cada 2 min |
| **EmailTags · Enviados** | Cada correo que envías: si lo clasificaste en el panel, lo manda a archivar; si no, **hereda** carpeta y categorías del hilo ya clasificado. | al enviar |

## Valores que vas a copiar

| Nombre | Valor |
|---|---|
| `PROP` | `String {00020329-0000-0000-C000-000000000046} Name cecp-7c1e6f0a-4d2b-4b8e-9a3f-2e5d8c1b6a47` |
| `ELIMINADOS` (ID de Elementos eliminados) | `AAMkAGYwODM0MDA3LWI1MjUtNDc4ZS05MThlLWYzMzM3NWU1NmVlNwAuAAAAAAB3VnRzaEwBQYKqdBldNTSOAQDO-xdreMjfQ6cEllJkstX3AAAAAAEKAAA=` |
| `BANDEJA` (ID de Bandeja de entrada) | `AAMkAGYwODM0MDA3LWI1MjUtNDc4ZS05MThlLWYzMzM3NWU1NmVlNwAuAAAAAAB3VnRzaEwBQYKqdBldNTSOAQDO-xdreMjfQ6cEllJkstX3AAAAAAEMAAA=` |

> **Muy importante:** renombra cada acción **exactamente** como se indica (el nombre entre `comillas`), porque las expresiones las referencian por nombre. Donde dice *expresión*, pégala en la pestaña **fx / Expresión**, no como texto.
> En todas las acciones "Enviar una solicitud HTTP", el campo **Content-Type** se deja en `application/json`.

---

## Flujo 1 · "EmailTags · Archivar"

**Crear → Flujo de nube programado** · Repetir cada **2 minutos**.
En la configuración del desencadenador (⋯ → Configuración) activa **Control de simultaneidad = 1**.

### 1. `Pendientes` — Office 365 Outlook › Enviar una solicitud HTTP
- **Método:** `GET`
- **URI:**
```
https://graph.microsoft.com/v1.0/me/messages?$top=10&$select=id,conversationId,parentFolderId,categories&$filter=categories/any(c:c eq 'ET-Pendiente')&$expand=singleValueExtendedProperties($filter=id eq 'String {00020329-0000-0000-C000-000000000046} Name cecp-7c1e6f0a-4d2b-4b8e-9a3f-2e5d8c1b6a47')
```

### 2. `Cada_pendiente` — Aplicar a cada uno
- Entrada (*expresión*): `body('Pendientes')?['value']`

Dentro de `Cada_pendiente`:

### 3. `Tiene_decision` — Condición
- Izquierda (*expresión*): `empty(items('Cada_pendiente')?['singleValueExtendedProperties'])`
- Operador: **es igual a** · Derecha (*expresión*): `false`

#### Rama **Sí**

**3.1 `Decision` — Redactar (Compose)**
*expresión:*
```
json(first(items('Cada_pendiente')?['singleValueExtendedProperties'])?['value'])?['d']
```

**3.2 `Marcar` — Enviar una solicitud HTTP**
- Método `PATCH` · URI: `https://graph.microsoft.com/v1.0/me/messages/` + contenido dinámico *expresión* `items('Cada_pendiente')?['id']`
- Cuerpo (*expresión*):
```
if(equals(outputs('Decision')?['fl'], 1), addProperty(setProperty(json('{}'), 'categories', outputs('Decision')?['c']), 'flag', json('{"flagStatus":"flagged"}')), setProperty(json('{}'), 'categories', outputs('Decision')?['c']))
```
*(Esto pone las categorías elegidas, la bandera si la pediste y quita `ET-Pendiente`.)*

**3.3 `Hilo` — Enviar una solicitud HTTP**
- Método `GET` · URI:
```
https://graph.microsoft.com/v1.0/me/messages?$top=50&$select=id,parentFolderId,isDraft&$filter=conversationId eq '@{items('Cada_pendiente')?['conversationId']}'
```

**3.4 `Destinos` — Filtrar matriz**
- De (*expresión*): `body('Hilo')?['value']`
- Editar en modo avanzado y pegar:
```
@and(not(equals(item()?['isDraft'], true)), not(equals(item()?['parentFolderId'], 'AAMkAGYwODM0MDA3LWI1MjUtNDc4ZS05MThlLWYzMzM3NWU1NmVlNwAuAAAAAAB3VnRzaEwBQYKqdBldNTSOAQDO-xdreMjfQ6cEllJkstX3AAAAAAEKAAA=')), or(equals(outputs('Decision')?['th'], 1), equals(item()?['id'], items('Cada_pendiente')?['id'])))
```
*(Si elegiste "aplicar a todo el hilo" toma todos los correos de la conversación; si no, solo este.)*

**3.5 `Cada_destino` — Aplicar a cada uno**, entrada *expresión* `body('Destinos')`. Dentro:

- **`Es_otro_del_hilo` — Condición:** *expresión* `items('Cada_destino')?['id']` **no es igual a** *expresión* `items('Cada_pendiente')?['id']`
  - **Sí → `Etiquetar_hilo`** — Enviar una solicitud HTTP · `PATCH` · URI `https://graph.microsoft.com/v1.0/me/messages/@{items('Cada_destino')?['id']}` · Cuerpo (*expresión*): `setProperty(json('{}'), 'categories', outputs('Decision')?['c'])`
- **`Hay_que_mover` — Condición** (debajo de la anterior, no dentro):
  - Izquierda (*expresión*): `and(not(empty(outputs('Decision')?['f'])), not(equals(items('Cada_destino')?['parentFolderId'], outputs('Decision')?['f'])))`
  - **es igual a** `true`
  - **Sí → `Mover`** — Enviar una solicitud HTTP · `POST` · URI `https://graph.microsoft.com/v1.0/me/messages/@{items('Cada_destino')?['id']}/move` · Cuerpo (*expresión*): `setProperty(json('{}'), 'destinationId', outputs('Decision')?['f'])`

#### Rama **No** (alguien puso `ET-Pendiente` a mano: solo se limpia)

**`Sin_marca` — Filtrar matriz** · De: `items('Cada_pendiente')?['categories']` · modo avanzado: `@not(equals(item(), 'ET-Pendiente'))`

**`Limpiar` — Enviar una solicitud HTTP** · `PATCH` · URI `https://graph.microsoft.com/v1.0/me/messages/@{items('Cada_pendiente')?['id']}` · Cuerpo (*expresión*): `setProperty(json('{}'), 'categories', body('Sin_marca'))`

---

## Flujo 2 · "EmailTags · Enviados"

**Crear → Flujo de nube automatizado** · desencadenador **Office 365 Outlook › Cuando llega un nuevo correo electrónico (V3)** · **Carpeta: Elementos enviados**.

### 1. `Enviado` — Enviar una solicitud HTTP
- `GET` · URI:
```
https://graph.microsoft.com/v1.0/me/messages/@{triggerOutputs()?['body/id']}?$select=id,conversationId,categories,parentFolderId&$expand=singleValueExtendedProperties($filter=id eq 'String {00020329-0000-0000-C000-000000000046} Name cecp-7c1e6f0a-4d2b-4b8e-9a3f-2e5d8c1b6a47')
```

### 2. `Clasificado_en_panel` — Condición
- Izquierda (*expresión*): `empty(body('Enviado')?['singleValueExtendedProperties'])` **es igual a** `false`

#### Rama **Sí** — lo clasificaste antes de enviar
**`Avisar_archivo`** — Enviar una solicitud HTTP · `PATCH` · URI `https://graph.microsoft.com/v1.0/me/messages/@{body('Enviado')?['id']}` · Cuerpo (*expresión*):
```
setProperty(json('{}'), 'categories', union(coalesce(body('Enviado')?['categories'], json('[]')), json('["ET-Pendiente"]')))
```
*(El flujo 1 lo archiva en ≤ 2 min, y si marcaste "todo el hilo" también archiva el correo original con los mismos atributos.)*

#### Rama **No** — no lo clasificaste: hereda del hilo

**`Hilo_enviado`** — Enviar una solicitud HTTP · `GET` · URI:
```
https://graph.microsoft.com/v1.0/me/messages?$top=50&$select=id,categories,parentFolderId&$filter=conversationId eq '@{body('Enviado')?['conversationId']}'
```

**`Clasificados`** — Filtrar matriz · De: `body('Hilo_enviado')?['value']` · modo avanzado:
```
@and(greater(length(coalesce(item()?['categories'], json('[]'))), 0), not(contains(coalesce(item()?['categories'], json('[]')), 'ET-Pendiente')), not(equals(item()?['parentFolderId'], body('Enviado')?['parentFolderId'])), not(equals(item()?['parentFolderId'], 'AAMkAGYwODM0MDA3LWI1MjUtNDc4ZS05MThlLWYzMzM3NWU1NmVlNwAuAAAAAAB3VnRzaEwBQYKqdBldNTSOAQDO-xdreMjfQ6cEllJkstX3AAAAAAEMAAA=')), not(equals(item()?['parentFolderId'], 'AAMkAGYwODM0MDA3LWI1MjUtNDc4ZS05MThlLWYzMzM3NWU1NmVlNwAuAAAAAAB3VnRzaEwBQYKqdBldNTSOAQDO-xdreMjfQ6cEllJkstX3AAAAAAEKAAA=')))
```
*(Busca en la conversación un correo que ya tenga categorías y esté archivado en una carpeta tuya; sirve también para lo que ya clasificaban tus reglas actuales.)*

**`Hay_base` — Condición:** *expresión* `length(body('Clasificados'))` **es mayor que** `0`

- **Sí:**
  - **`Base`** — Redactar · *expresión* `last(body('Clasificados'))`
  - **`Valor_decision`** — Redactar · *expresión*:
    ```
    concat('{"d":{"f":"', outputs('Base')?['parentFolderId'], '","c":', string(outputs('Base')?['categories']), ',"fl":0,"th":0}}')
    ```
  - **`Heredar`** — Enviar una solicitud HTTP · `PATCH` · URI `https://graph.microsoft.com/v1.0/me/messages/@{body('Enviado')?['id']}` · Cuerpo (*expresión*):
    ```
    setProperty(setProperty(json('{}'), 'categories', union(outputs('Base')?['categories'], json('["ET-Pendiente"]'))), 'singleValueExtendedProperties', createArray(setProperty(json('{"id":"String {00020329-0000-0000-C000-000000000046} Name cecp-7c1e6f0a-4d2b-4b8e-9a3f-2e5d8c1b6a47"}'), 'value', outputs('Valor_decision'))))
    ```
- **No:** nada (el correo se queda en Elementos enviados).

---

## Prueba rápida (5 min)

1. Abre cualquier correo → **Clasificar** → elige `1. Acción / 1. Hoy` → **Aplicar**. Verás las categorías al instante y `ET-Pendiente`.
2. Espera ≤ 2 min: el correo aparece en *1. Hoy* sin `ET-Pendiente`. En Power Automate, el historial de "Archivar" muestra la ejecución en verde.
3. Responde ese correo sin clasificar y envíalo: el aviso no te detiene (el hilo ya es conocido) y tu respuesta termina también en *1. Hoy*.

## Si algo falla

| Síntoma | Causa probable | Arreglo |
|---|---|---|
| Error 403 en "Enviar una solicitud HTTP" | Política DLP de HITSS bloquea esa acción | Pedir a TI que permita la acción *Send an HTTP request* del conector Office 365 Outlook (es estándar, sin licencia premium). |
| `ErrorInvalidIdMalformed` al mover | Cambiaste/creaste una carpeta | Actualiza el ID en `src/config.js` (pídele a Claude que relea tus carpetas). |
| El correo no se mueve y `ET-Pendiente` se queda | El flujo 1 está apagado o falló | Revisa el historial del flujo; se reintenta solo cada 2 min. |
| Demasiadas ejecuciones | Límite diario de Power Platform | Sube el intervalo de "Archivar" a 3–5 min. |
