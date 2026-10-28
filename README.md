# Cruzador Excel Web

Prototipo web local que reutiliza el lector y el motor de cruce del Cruzador de escritorio. No modifica ni reemplaza la aplicacion PySide6.

## Alcance implementado, pendiente de verificacion

El flujo web se amplio para acercarse a las funciones del cliente de escritorio:

- Carga de 2 a 10 archivos `.xlsx` o `.xls`, selector de hoja o combinacion de hojas y correccion manual de encabezados. Limites: 20 MB por archivo, 100 MB descomprimidos por `.xlsx`, 250.000 filas y 500 columnas por archivo.
- Limpieza configurable con las operaciones que expone el motor compartido.
- Cruces secuenciales `inner`, `left` y `outer`, hasta cinco pares de claves compuestas por paso y advertencias de duplicados.
- Continuacion de un cruce usando el resultado anterior como primera fuente.
- Columnas calculadas con reglas acotadas (formula binaria, condicion SI/ENTONCES/SINO y diferencia entre fechas); no se ejecutan expresiones arbitrarias.
- Vista previa, busqueda, filtros AND, edicion y eliminacion de filas, exportacion Excel y recetas persistentes que guardan estructura y reglas, no filas ni archivos.
- Estadisticas descriptivas y resumen agrupado configurable sobre el resultado del cruce. Las medidas incluyen conteos, vacios, suma, promedio, minimo y maximo, con decimales exactos para agregaciones numericas.
- Autenticacion local contra SQLite, proteccion CSRF, limites de sesion y auditoria diaria encadenada por hash. La auditoria no debe incluir valores de celdas.
- Navegacion guiada por pasos, regreso a pantallas anteriores conservando la carga y parámetros condicionales que solo se muestran para la opción seleccionada; las opciones auxiliares permanecen plegadas.
- Al seleccionar una receta, se puede elegir carga paso a paso o ejecución automática de encabezados, limpieza, cruce y columnas calculadas guardadas. La receta se aplica solo si puede identificar sin ambigüedad las columnas/hojas esperadas; ante diferencias, se conservan los archivos para continuar manualmente. Las claves duplicadas siempre requieren confirmación antes de cruzar. La ejecución automática termina en la vista previa; la exportación sigue siendo una acción del operador.

Esta ampliacion **no tiene aun ejecucion de pruebas ni QA independiente**. No representa paridad completa: las recetas no guardan configuraciones estadísticas, las estadisticas no permiten elegir un archivo antes del cruce, agrupar fechas por periodo ni filtrar la fuente estadistica; las expresiones calculadas no cubren la gramatica completa de porcentajes y condiciones compuestas definida en requisitos; los cruces encadenados no se pueden guardar como receta. La exportacion entrega el resultado de datos, no un libro con todas las herramientas del escritorio. Validar estos limites en [requisitos de estadisticas](../../salidas/requisitos/estadisticas-cruzador-excel.md).

## Seguridad y uso permitido

Esta entrega sigue siendo un prototipo para datos sinteticos. Por defecto, Flask/Gunicorn escucha solo en `127.0.0.1`; el VirtualHost Apache de ejemplo exige conexiones locales. No lo publiques en una LAN o Internet ni cargues datos reales.

La autenticacion local SQLite no es SSO institucional ni aporta por si sola autorizacion corporativa. Los datos de trabajo permanecen en memoria de un unico proceso y se pierden al reiniciar; por eso Gunicorn usa un solo worker. No es una arquitectura de produccion multiusuario. La opcion de continuar el cruce solo utiliza archivos que el operador carga manualmente: no conecta con Forms, Power Automate, ANSES, correo, CRM ni sistemas core.

Antes de habilitar datos reales se requiere diseno y aprobacion de IT, Seguridad y Compliance: SSO/MFA, autorizacion, sesiones persistentes seguras, CSRF revisado, almacenamiento temporal cifrado y aislado por usuario, limites de capacidad, politica de retencion, copias, monitoreo, restauracion, pruebas de seguridad y QA independiente.

## Ejecutar localmente para probar con datos sinteticos

Desde el repositorio:

```bash
cd desarrollo/cruzador_excel_web
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python scripts/create_user.py operador
export CRUZADOR_WEB_COOKIE_SECURE=0
export CRUZADOR_WEB_LOG_DIR="$HOME/.local/state/banco-piano/cruzador-excel-web/logs"
python app.py
```

El script solicita la contrasena de forma interactiva y la guarda en la base local de usuarios; no la pases como argumento. Abrir `http://127.0.0.1:8000` e iniciar sesion. `CRUZADOR_WEB_COOKIE_SECURE=0` solo es aceptable en esta prueba limitada a loopback y HTTP; en el sitio Apache HTTPS debe permanecer en `1`.

Para ejecutar las pruebas web y las del motor compartido:

```bash
python -m pytest tests/
cd ../cruzador_excel
python -m pytest tests/
```

Los cambios de ampliacion no deben considerarse verificados hasta ejecutar esta suite y completar QA independiente.

## Prueba temporal desde otra PC de la LAN

El acceso remoto esta deshabilitado por defecto. Para una prueba controlada con datos sinteticos, configura en la PC servidor la IP de esa PC y autoriza unicamente la IP cliente. Por ejemplo, si el servidor es `192.168.5.160` y el cliente `192.168.1.70`, inicia el servidor asi:

```bash
cd desarrollo/cruzador_excel_web
source .venv/bin/activate
CRUZADOR_WEB_BIND_HOST=192.168.5.160 \
CRUZADOR_WEB_ALLOWED_CLIENTS=192.168.1.70 \
CRUZADOR_WEB_COOKIE_SECURE=0 \
python app.py
```

Desde la PC cliente, abre `http://192.168.5.160:8000`. Si no conecta, confirma en el servidor que la interfaz tiene esa IP y que el puerto esta escuchando (`sudo ss -ltnp | grep ':8000'`); en Windows, `Test-NetConnection 192.168.5.160 -Port 8000`. Como `192.168.5.x` y `192.168.1.x` suelen ser subredes distintas, la red debe enrutar entre ambas y el firewall del servidor debe permitir TCP/8000 desde `192.168.1.70`.

La lista de IPs es solo una restriccion de red y no reemplaza el login local ni acredita identidad institucional. HTTP tampoco cifra el trafico. No uses archivos reales, no abras el puerto a toda la LAN/Internet y detene el proceso al terminar (`Ctrl+C`). Para uso compartido real se necesita HTTPS, autenticacion institucional y revision de IT/Seguridad/Compliance. El servicio systemd y Apache de este prototipo permanecen intencionalmente limitados a loopback.

## Preparar Apache2 local con HTTPS

Se necesita Linux, Apache2, un certificado TLS local valido para el nombre utilizado y permisos administrativos. La configuracion es una plantilla; reemplaza rutas de certificado y `ServerName` por los aprobados para el entorno.

1. Instala Apache2 y Python venv; habilita los modulos `proxy`, `proxy_http`, `ssl` y `headers`.
2. Despliega juntos `cruzador_excel_web/` y `cruzador_excel/` bajo `/opt/banco-piano/desarrollo/`, y crea `/opt/banco-piano/venv` con las dependencias de este README. Copia solo codigo/documentacion necesaria: excluye `.venv*`, `build/`, `*.db*`, `.logs/`, `recipes/` y archivos Excel sinteticos o de trabajo.
3. Configura `/etc/cruzador-excel-web.env` con un valor secreto aleatorio para `CRUZADOR_WEB_SECRET` y `CRUZADOR_WEB_LOG_DIR=/var/log/cruzador-excel-web`. Protege ese archivo para que solo root y el usuario del servicio puedan leerlo.
4. Copia `systemd/cruzador-excel-web.service` a `/etc/systemd/system/`.
5. Configura un certificado en las rutas indicadas por `apache/cruzador-excel-web.conf`, revisa el VirtualHost y copialo a la carpeta `sites-available` de Apache.
6. Habilita el sitio, valida la configuracion de Apache, reinicia Apache y arranca el servicio systemd. En Debian/Ubuntu, el flujo es:

   ```bash
   sudo apt install apache2 python3-venv
   sudo a2enmod proxy proxy_http ssl headers
   sudo cp apache/cruzador-excel-web.conf /etc/apache2/sites-available/
   sudo cp systemd/cruzador-excel-web.service /etc/systemd/system/
   sudo install -m 0640 -o root -g www-data /dev/null /etc/cruzador-excel-web.env
   python3 -c 'import secrets; print(secrets.token_hex(32))'
   sudoedit /etc/cruzador-excel-web.env
   sudo a2ensite cruzador-excel-web.conf
   sudo apachectl configtest
   sudo systemctl daemon-reload
   sudo systemctl enable --now cruzador-excel-web
   sudo systemctl reload apache2
   curl -I https://localhost/
   ```

   En el archivo de entorno establece `CRUZADOR_WEB_SECRET` con el secreto generado, y `CRUZADOR_WEB_LOG_DIR=/var/log/cruzador-excel-web`. Los paquetes Python deben instalarse antes en el entorno virtual indicado por el servicio:

   ```bash
   sudo python3 -m venv /opt/banco-piano/venv
   sudo /opt/banco-piano/venv/bin/pip install -r /opt/banco-piano/desarrollo/cruzador_excel_web/requirements.txt
   ```

   Ajusta los paths del unit file si la ubicacion cambia. Si aun no hay un certificado local, IT debe emitir uno; un certificado autofirmado sirve solo para prueba personal y generara advertencias en el navegador.
7. Comprueba desde el mismo servidor `https://localhost/`. La directiva `Require local` bloquea clientes remotos de forma intencional.

Apache actua como terminador HTTPS y proxy inverso. Gunicorn atiende solo en `127.0.0.1:8000`; no debe exponerse directamente. No cambies `Require local` para habilitar acceso remoto: para eso se necesita autenticacion institucional, controles de red y una revision formal de Seguridad.

## Deteccion de encabezados

El lector comun ahora inspecciona hasta 20 filas, puntua encabezados candidatos por cantidad y unicidad de etiquetas y consistencia de tipos en las filas siguientes. Conserva los nombres de columna y la deteccion queda registrada en `FileMetadata`; puede elegirse explicitamente otra fila mediante `ExcelFileReader.read(..., header_row=<indice 0-based>)`. En la web se muestra la fila en numeracion humana (1-based) para que el operador la confirme.

La heuristica no es infalible. Siempre debe revisarse la fila y el significado de las columnas antes del cruce; no valida identidades ni reglas de negocio.

El lector todavia no completa celdas combinadas ni advierte sobre filas/hojas ocultas. Si un identificador fue almacenado como numero en Excel, los ceros a la izquierda pueden haberse perdido antes de la lectura; confirma ese tipo de casos antes de usar el resultado.

## Pruebas

Los tests del lector estan en `../cruzador_excel/tests/test_file_reader.py`; el flujo web tiene pruebas en `tests/test_web_flow.py`. Ejecuta ambas suites en sus entornos con dependencias instaladas y datos sinteticos. Esta entrega aun requiere QA de seguridad independiente antes de cualquier piloto.
