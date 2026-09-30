# EmailTags HITSS — clasificación tipo *Standss EmailTags* para el nuevo Outlook

## Qué hace

**1. Botón "Clasificar" en cada correo recibido.** Abre un panel que en menos de un segundo te sugiere carpeta y categorías (con % de confianza y el porqué), te muestra **todas** tus carpetas (con buscador) y **todas** tus categorías (agrupadas: Cliente, Proyecto, Tipo, Estado, Prioridad y las demás de Outlook). Si la sugerencia te sirve, presionas **Enter**. Si fijas el panel (📌), la sugerencia se actualiza sola cada vez que seleccionas otro correo.

**2. Antes de enviar.**
- Si el hilo ya está clasificado, el correo sale sin preguntar y hereda carpeta y categorías.
- Si el hilo no está clasificado, Outlook te detiene con el botón **🧠 Clasificar**. Eliges carpeta y categorías, presionas Enviar, y **tu correo y el correo original** quedan archivados con los mismos atributos.
- Si prefieres no clasificarlo en ese momento, usa **Enviar de todos modos** y el flujo intenta heredar la clasificación del hilo.

**3. Aprende de ti.** Cada decisión alimenta un pequeño modelo guardado dentro de tu buzón (hilos, contactos, dominios y palabras del asunto). Cuando corriges una sugerencia, esa corrección pesa el doble. Después de más de 5 decisiones iguales para un contacto la confianza sube, y después de más de 10 la sugerencia aparece marcada como **DIRECTA**. Todo esto funciona sin Excel, sin SharePoint y sin servidores externos.

## Arquitectura (y por qué)

```
┌──────────── Nuevo Outlook ────────────┐        ┌──── Power Automate (estándar) ────┐
│ Add-in "Clasificar" (panel)           │        │ Flujo "Archivar" (cada 2 min)     │
│  · sugiere (motor local, instantáneo) │ marca  │  · mueve correo / hilo a carpeta  │
│  · aplica categorías al instante      │ ─────▶ │  · pone bandera, quita la marca   │
│  · guarda la decisión en el correo    │        │ Flujo "Enviados" (al enviar)      │
│ Smart Alert al enviar                 │        │  · archiva o hereda del hilo      │
└───────────────────────────────────────┘        └───────────────────────────────────┘
```

En el nuevo Outlook, un add-in **no puede mover correos** ni poner categorías mientras redactas; son limitaciones documentadas de Office.js. Por eso el movimiento lo hace un flujo que usa solo el conector estándar de Outlook. Las categorías de un correo recibido se aplican al instante, y el movimiento ocurre en un máximo de 2 minutos.

## Instalación (≈ 30 min, sin pedirle nada a TI si tu buzón lo permite)

### Paso 1 — Publicar los archivos (HTTPS)
Un add-in de Outlook se carga desde una URL HTTPS. La opción gratis y rápida:
1. Crea en GitHub un repositorio **público** llamado `emailtags` y sube **todo el contenido** de esta carpeta (`src/`, `assets/`, etc.).
2. En el repositorio, entra a **Settings → Pages → Deploy from a branch → main / (root)**.
3. Tu URL base queda así: `https://TU-USUARIO.github.io/emailtags`

*(También lo puedes subir a una carpeta de oktenia.com, porque cualquier hosting con HTTPS sirve. El código no contiene contraseñas, pero `src/config.js` sí muestra los nombres de tus carpetas.)*

### Paso 2 — Poner la URL en el manifest
Abre `manifest.xml` en el Bloc de notas, usa **Ctrl+H** para reemplazar `__BASE__` por tu URL base (sin `/` al final) y guarda.

### Paso 3 — Instalar el add-in en tu Outlook
1. En el navegador, abre **https://aka.ms/olksideload**. Se abre Outlook web en *Complementos para Outlook*.
2. Ve a **Mis complementos → Complementos personalizados → Agregar un complemento personalizado → Agregar desde archivo…** y elige `manifest.xml`.
3. Reinicia el nuevo Outlook. Al abrir un correo, el botón aparece en la cinta o en el ícono **Aplicaciones** con el nombre **EmailTags › Clasificar**.

> Si no aparece la opción *Complementos personalizados*, TI bloqueó la instalación de add-ins propios. En ese caso pide que publiquen el `manifest.xml` solo para ti desde *Implementación centralizada* en el Centro de administración de Microsoft 365; lo pueden hacer en 5 minutos.

### Paso 4 — Crear los 2 flujos
Sigue `flows/FLUJOS-POWER-AUTOMATE.md`. Ahí están las URL y expresiones exactas para copiar y pegar.

### Paso 5 — Probar
En la sección "Prueba rápida" de esa misma guía hay una prueba de 3 pasos.

## Uso diario
| Acción | Cómo |
|---|---|
| Clasificar un correo | Selecciónalo → **Clasificar** → **Enter** |
| Clasificar muchos seguidos | Fija el panel 📌 y ve bajando con ↓; la sugerencia cambia sola |
| Buscar una carpeta | Tecla **/** y escribe (p. ej. "waiting") |
| Solo etiquetar, sin mover | Elige "— No mover —" |
| Que no se aplique al hilo | Desmarca "Aplicar a todo el hilo" |
| Respaldo del aprendizaje | Panel → Aprendizaje → *Copiar respaldo* |

## Personalizar (`src/config.js`)
- **folders**: tus carpetas con su ID. Vienen solo las 14 de tu estructura nueva: 1. Acción, 2. Proyectos, 3. Gestión y 4. Comercial. Las carpetas viejas de Archivo no aparecen.
- **groups**: el orden de las categorías y cuáles son exclusivas (Estado, Prioridad).
- **rules**: palabras clave con las que arranca el sistema antes de tener historial (telcel, insumos, swf, incidente, propuesta, etc.).
- **askOnSend**: `"unknown"` (valor por defecto), `"always"` o `"never"`.

Después de editar, sube el archivo de nuevo a GitHub. No hace falta reinstalar el add-in.

## Límites que debes conocer
- **El movimiento tarda entre 1 y 2 minutos**, porque el flujo corre cada 2 minutos. Mientras tanto verás la categoría `ET-Pendiente`.
- **Al redactar, las categorías no se ven** en el borrador, porque el nuevo Outlook no lo permite. Se aplican al correo enviado y al original justo después de enviar.
- **Outlook móvil** no muestra el aviso previo al envío. Aun así, el flujo "Enviados" hereda la clasificación del hilo.
- **El aprendizaje es de tu buzón**: te sigue entre el nuevo Outlook y Outlook web. Tiene un tope de unos 28 KB y se depura solo, empezando por lo más viejo.
- **La acción "Enviar una solicitud HTTP"** del conector de Outlook está en versión preliminar. Si una política DLP de HITSS la bloquea, eso es lo único que tendrías que pedir a TI.
