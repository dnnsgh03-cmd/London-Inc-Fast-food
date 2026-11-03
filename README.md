---
title: London Inc Fast Food
sdk: docker
app_port: 7860
---

# London Inc Fast Food

## Archivos del proyecto

- `templates/index.html`: estructura de las secciones del panel.
- `style.css`: estilos y diseño adaptable.
- `script.js`: menú, pedidos, pagos, caja, inventario, balance simple y acciones de la interfaz.
- `app.py`: servidor Flask y conexión con el libro Excel.
- `Procfile`: comando de Gunicorn para hosting Linux.
- `static/manifest.json`: configuración para instalar la app como PWA.
- `londoninc.xlsx`: archivo de datos local, excluido del repositorio; se crea si no existe.

## Subir el proyecto a GitHub

1. Crea un repositorio **privado** en GitHub. El código incluye nombres del equipo y configuración propia del negocio; revisa esos datos antes de cambiar el repositorio a público.
2. Sube `app.py`, `requirements.txt`, `Procfile`, `Dockerfile`, `.dockerignore`, `.gitignore`, este `README.md`, `templates/`, `static/` y las imágenes necesarias (`menu.jpg` y `menu1.jpg`, si se usan).
3. **No subas** `londoninc.xlsx` ni copias o respaldos del libro: contiene pedidos y datos del negocio. `.gitignore` los excluye. Tampoco subas archivos `.env`, contraseñas, `SECRET_KEY` ni webhooks de Discord; configura esos valores en los secretos del servicio donde despliegues la aplicación.
4. Después de crear el repositorio, usa **Add file → Upload files** en GitHub o GitHub Desktop para cargar esos archivos y carpetas. Confirma la lista de archivos antes de hacer el primer commit. Si un secreto o el Excel se publicó por accidente, borrarlo en un commit posterior no lo elimina del historial; cambia las credenciales/webhook expuestos y limpia el historial del repositorio.

## Dónde modificar la plataforma

- Apariencia, colores, tipografía, tamaños y diseño adaptable: `static/style.css`.
- Título, encabezado y estructura principal: `templates/index.html`.
- Textos de las pantallas, botones y distribución de cada sección: plantillas `V` en `script.js`.
- Productos, precios, inventario, balance, pagos, pedidos y caja: pestañas visibles del libro `londoninc.xlsx`. Después de editarlo, usa **Importar desde Excel** en la plataforma para cargar esos cambios.

## Desplegar en Render o Railway

1. Sube el proyecto a un repositorio privado y crea un servicio web Python. La plataforma instala `requirements.txt` y usa `Procfile`; Gunicorn se enlaza al puerto `PORT` que proporciona el hosting.
2. Configura en el panel de secretos del hosting `APP_PASSWORD` (contraseña larga del panel), `SECRET_KEY` (clave aleatoria persistente), `SESSION_COOKIE_SECURE=true` y, si usarás Discord, `DISCORD_WEBHOOK_URL`. Genera `SECRET_KEY` con `python -c "import secrets; print(secrets.token_hex(32))"`. No guardes secretos en el repositorio. La ruta `/healthz` sirve como health check.
3. El panel requiere contraseña y clave de sesión; sin ambas no permite entrar. Las sesiones usan cookies `HttpOnly`, `SameSite=Lax`, HTTPS en producción y vencen a las 8 horas.
4. **Persistencia de Excel:** el archivo se guarda en `LONDON_FOOD_EXCEL` si defines esa variable; si no, se crea junto al código. El disco local de un servicio cloud puede borrarse al reiniciar o redesplegar. Monta un volumen persistente y establece `LONDON_FOOD_EXCEL` en una ruta dentro de él (por ejemplo `/var/data/londoninc.xlsx` en un volumen de Render o `/data/londoninc.xlsx` en Railway). Los discos persistentes de Render dependen del plan; verifica que tu plan los incluya. `*.xlsx` está excluido por `.gitignore` para no publicar datos del negocio. Copia el libro existente al volumen antes de iniciar la app si necesitas conservar el historial. El `Procfile` limita Gunicorn a un worker para evitar escritores concurrentes sobre Excel. Para múltiples instancias o mayor concurrencia, migra a una base de datos administrada; SQLite también necesita un volumen persistente y no es adecuada para escritura multiinstancia.
5. El manifiesto y el service worker permiten instalar la PWA y almacenar en caché recursos estáticos. Las operaciones y datos requieren conexión al servidor; las respuestas privadas de la API no se guardan en caché. El sitio debe servirse por HTTPS para instalar la PWA en producción.

## Desplegar en Hugging Face Spaces (Docker)

1. Crea un Space nuevo con SDK **Docker** y sube el `Dockerfile`, `requirements.txt`, `app.py`, `templates/` y `static/`. El encabezado YAML de este README declara el SDK y el puerto 7860.
2. En **Settings → Variables and secrets**, configura `APP_PASSWORD` y `SECRET_KEY` como secretos. Genera una clave persistente con `python -c "import secrets; print(secrets.token_hex(32))"`. Configura también `SESSION_COOKIE_SECURE=true`; agrega `DISCORD_WEBHOOK_URL` solo si vas a usar las notificaciones. No subas secretos ni el libro de Excel al repositorio.
3. Crea un **Storage Bucket** y adjúntalo al Space como volumen de lectura/escritura montado en `/data`. El contenedor guarda el libro en `/data/londoninc.xlsx`; el disco efímero del Space no protege los datos de reinicios o reconstrucciones. Si vas a migrar un libro local, súbelo al bucket con el nombre `londoninc.xlsx` antes de adjuntarlo. Conserva además una copia de respaldo.
4. En **Files and versions**, mantén en la raíz del Space `Dockerfile`, `.dockerignore`, `README.md`, `requirements.txt`, `app.py`, `templates/` y `static/`. No subas el Excel con el código: usa el Storage Bucket para los datos. Si el libro no existe en `/data`, el servidor crea uno nuevo automáticamente.
5. El contenedor inicia Gunicorn con un worker en `0.0.0.0:7860`. El puerto 7860 también es el que declara el Space. `pandas` no se requiere: la aplicación usa `openpyxl` para los libros de Excel.

Al registrar un pedido, el servidor envía a Discord el cliente, dirección, productos, comentarios y total. Configura `DISCORD_WEBHOOK_URL` en el entorno; no guardes el webhook en el código ni lo compartas. Si falta la variable, el servidor imprime una advertencia y no envía la alerta. Si Discord no está disponible, el pedido queda guardado y la plataforma avisa que falló la notificación. Revoca cualquier webhook que se haya compartido anteriormente y usa uno nuevo.

Cada cambio hecho desde la plataforma se envía al servidor y se guarda en `londoninc.xlsx`; el indicador debajo de las pestañas muestra si se está guardando, si terminó o si hubo un error. La confirmación de pedido, pago o gasto aparece después del guardado. Si Excel está abierto, el panel bloquea la escritura para proteger el archivo y conserva los cambios pendientes en el navegador; cierra Excel y pulsa **Sincronizar con Excel** para guardarlos.

Los datos se organizan en las hojas `Menu`, `Pedidos`, `Pagos`, `Gastos`, `Inventario`, `Balance` y `Configuracion`. La hoja `Balance` registra concepto, categoría, valor e impacto; el balance neto resta los ítems marcados como “No suma / Resta” de los que “Suma al negocio”. Insumos y equipos se marcan para sumar por defecto; deudas y gastos, para restar. Puedes cambiar la etiqueta para señalar mermas o insumos vencidos. La hoja oculta `AppState` conserva el estado íntegro de la aplicación. Los pedidos permiten guardar comentarios o instrucciones de preparación, que también se muestran en el historial y se guardan en Excel. En la pestaña **Excel** puedes probar la conexión, sincronizar, importar y descargar el libro. También puedes descargar un informe Excel de un día específico o de todo el historial, con resumen, registros y una hoja de detalle de pedidos que incluye fecha, hora, cliente, tipo, medio de pago, producto, cantidad, precio, subtotal, total, estado y comentarios. Cierra el libro en Excel mientras la plataforma escribe; vuelve a abrirlo para ver los cambios guardados. Los cambios hechos directamente en Excel se cargan con **Importar desde Excel**.

En **Pedidos**, usa **Editar** en el historial para cargar un pedido registrado, ajustar productos, cliente, tipo de pedido, pago o comentarios y guardar los cambios sobre el mismo número de pedido. Se conservan su fecha, hora, estado y cualquier dirección que ya tuviera. Después de guardar, Discord recibe una notificación de edición con los campos modificados resaltados y los valores anteriores/nuevos; si Discord falla, el pedido queda guardado y la plataforma muestra un aviso. Los productos antiguos que ya no están en el menú se mantienen en el pedido al editarlo.

La sección **Personal** muestra los nombres y roles del equipo registrados. Las vacantes marcadas como «por decidir» se muestran en el equipo, pero no aparecen en la lista para registrar pagos.

Para cambiar la ruta del libro, define `LONDON_FOOD_EXCEL` antes de iniciar `app.py`.
