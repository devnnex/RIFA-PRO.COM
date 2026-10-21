# Rifa Pro

Sitio estático para GitHub Pages con Google Apps Script y Google Sheets como servidor compartido. Diseño crema y azul oscuro inspirado en la boleta, vista móvil, modo oscuro, instalación PWA y múltiples rifas.

## Despliegue del cliente y del servidor

El repositorio contiene el cliente para GitHub Pages y el archivo `backend/Code.gs`, listo para aplicar manualmente en Apps Script. Ambos despliegues son independientes. La interfaz permite consultar el servicio anterior, pero bloquea las escrituras hasta recibir `protocol: 2`; no usa la ruta anterior que permitía sobrescribir números.

El acceso al resumen privado y las nuevas reservas **requieren desplegar `backend/Code.gs`**. Cambiar únicamente los archivos de GitHub Pages no permite hacer privadas las respuestas del servidor anterior.

## Comportamiento

- Los participantes eligen un número libre, ingresan nombre y celular y confirman. La reserva pertenece a la primera solicitud que el servidor acepta dentro de su bloqueo, incluso si llegan dentro del mismo segundo. Abrir el formulario no reserva un número.
- La otra persona recibe un modal con el nombre de quien lo reservó y la hora registrada por el servidor, con segundos y zona horaria de Colombia.
- Cada clic actualiza la vista inmediatamente con la marca **Confirmando**. Las operaciones solo se muestran como confirmadas cuando responde el servidor. Los reintentos conservan su ID y se mantienen únicamente en la memoria de la pestaña, sin almacenar datos de participantes en el dispositivo. Si falla la conexión, **Reintentar** conserva la solicitud original.
- El servidor valida la contraseña y entrega una sesión de 30 minutos. El token permanece solo en memoria; cerrar sesión o recargar elimina el acceso local. El servidor puede terminar antes una sesión si su caché se elimina.
- Solo el administrador recibe los KPI, los celulares, los estados de pago y sus medios. La respuesta pública agrupa apartados y pagados como reservados, por lo que no revela el recaudo al inspeccionar la red.
- Al marcar **Pagado**, se exige escoger **Nequi / transferencia** (violeta, teléfono) o **Efectivo** (verde, billete). El resumen muestra ambos subtotales. Registrar un medio es un registro manual, no una confirmación bancaria ni una integración con Nequi.
- Las ediciones y liberaciones administrativas usan una revisión para impedir que una pestaña desactualizada sobrescriba cambios de otra.
- El administrador puede cambiar el premio y el nombre desde **Editar premio**, junto a la boleta. La vista previa cambia al escribir; el guardado conserva números, participantes, precio y pagos. Una revisión impide sobrescribir una edición de otra sesión.
- El sitio conserva solo los metadatos públicos de las rifas para consulta sin conexión. No guarda teléfonos, medios de pago, totales ni tokens en caché o almacenamiento persistente.
- **Elige tu número** pulsa suavemente y lleva directamente al tablero, sin consultas de red. Se respeta la preferencia de movimiento reducido; instalar la PWA sigue disponible en el pie de página.

## Cartones independientes y activación

Cualquier persona puede elegir **Crear mi cartón**. Completa su rifa y celular y elige un **PIN de cuatro dígitos**, repetido para evitar errores. El servidor registra una solicitud por **$1.000 COP**. Una vez aprobado el pago por el desarrollador, el organizador pulsa **Ya tengo un cartón**, escribe únicamente su PIN y recupera su rifa con los números, reservas y pagos guardados. No se pide una URL. Cada PIN es único entre cartones; si está ocupado, el formulario pide elegir otro antes de registrar la solicitud. Cada organizador gestiona únicamente su cartón. La página general no publica un catálogo de rifas. Todas se consultan por `?rifa=ID`.

El enlace compartido siempre abre la vista de **participante**, sin sesión administrativa. El organizador ve **Copiar enlace para clientes** dentro de su sesión; el botón copia únicamente `?rifa=ID`, muestra confirmación y nunca agrega PIN, token ni clave privada. El botón no aparece a participantes. La pantalla del pago permite guardar o descargar el enlace del cartón; el PIN se recuerda por separado y no se almacena en el navegador. Cerrar sesión vuelve a la vista pública del mismo cartón, oculta sus KPI y borra los campos de contraseña.

Cada PIN se guarda como hash SHA-256 con sal individual y un secreto del servidor, en `RifaPro_Accesos`. Los intentos fallidos se guardan por cartón para el ingreso desde su enlace: ocho fallos bloquean ese acceso por cinco minutos. El ingreso general solo con PIN tiene además un límite global persistente de ocho fallos por cinco minutos, para evitar búsquedas de códigos; el acceso del desarrollador tiene su propio límite independiente. Un índice hash permite localizar el cartón sin publicar el PIN. No se crea una propiedad de secuencia de comandos por PIN. `OWNER_PIN_PEPPER` es un único secreto generado automáticamente; consérvalo al respaldar el proyecto.

Los accesos privados anteriores con `#acceso=…` siguen funcionando como recuperación. Su dueño puede entrar y configurar su PIN si aún no lo tiene. También puede cambiarlo desde **Cambiar mi PIN** dentro del cartón. El desarrollador puede actualizar el PIN desde la lista de cartones para resolver códigos antiguos repetidos; el ingreso general nunca selecciona una rifa arbitraria ante una coincidencia múltiple. La clave del acceso anterior conserva su hash y se verifica en el servidor; nunca se incluye en el enlace para clientes. Durante la preparación del pago se conserva ese acceso de recuperación en `sessionStorage` y se elimina al cerrar sesión. Los tokens de sesión permanecen en memoria.

La pantalla de pago ofrece la llave **3208759798**, Nequi / Bre-B, beneficiario **Elkin Rangel**, copia de llave y descarga PNG de una tarjeta con los datos y la referencia. No genera un QR bancario: la imagen original adjunta en el chat no está disponible como archivo en el repositorio y debe incorporarse como archivo original para preservar su contenido.

**Ya hice mi pago** abre WhatsApp al **+57 3246394689**, con el texto sobre el pago del cartón, su referencia y la indicación de adjuntar el comprobante. La persona adjunta la imagen manualmente. Abrir WhatsApp solo marca la solicitud como pendiente de revisión; no verifica transferencias ni activa el cartón.

El **desarrollador** entra desde el botón **Desarrollador**, junto a **Administrador**, o mediante `?panel=desarrollador`, con su PIN exclusivo de cuatro dígitos configurado en el servidor, y ve todos los cartones: pendientes, activos, vencidos y eliminados, incluidas las rifas anteriores. Puede filtrar la lista y abrir el enlace público de cada rifa activa. Después de verificar el ingreso y el comprobante, aprueba la referencia correspondiente. Solo entonces se habilitan el ingreso por PIN y el enlace público. Si el creador dejó abierta la pantalla del pago, esta consulta el estado y habilita el acceso al recibir la aprobación.

El acceso del desarrollador está separado del organizador tanto en la interfaz como en el servidor. La contraseña de una rifa no sirve para entrar al panel del desarrollador, consultar el listado global ni aprobar cartones. Las sesiones globales de versiones anteriores se revocan al aplicar esta separación. El panel de desarrollador muestra el estado de venta de los cartones; el recaudo de los números se gestiona desde el acceso del organizador de cada rifa.

Un cartón vence **al terminar el segundo día calendario después del sorteo, en America/Bogota**. Por ejemplo, un sorteo del 31 de octubre queda disponible hasta el 2 de noviembre a las 23:59:59; cierra el 3 de noviembre a las 00:00. El servidor impide nuevas reservas y ediciones de números o premios después del vencimiento, incluso desde clientes desactualizados. Los registros se conservan para consulta; no se borran automáticamente.

## Configurar el servidor

1. En el proyecto de Apps Script vinculado a tu hoja original, respalda el código anterior y copia el nuevo `backend/Code.gs` completo. Se conserva la pestaña `rifas`, sus 13 columnas originales y las filas `__META__`.
2. Copia `backend/Code.gs`. En **Configuración del proyecto → Propiedades de secuencia de comandos**, agrega:

   | Propiedad | Valor |
   | --- | --- |
   | `SPREADSHEET_ID` | Opcional si ejecutas `prepararSistema` desde el proyecto vinculado a la hoja original; esa función guarda su ID automáticamente. Requerido para proyectos independientes. |
   | `DEVELOPER_PIN` | PIN exclusivo del desarrollador, de cuatro dígitos. Configúralo solo en las propiedades del servidor; nunca en archivos públicos. Sustituye la anterior propiedad `DEVELOPER_PASSWORD`. |
   | `ADMIN_PIN` | Solo para importar el PIN de la rifa anterior al ejecutar `prepararSistema`. Después el ingreso se valida desde `RifaPro_Accesos`. Los nuevos organizadores eligen su PIN en el formulario, sin editar propiedades. |
   | `LEGACY_RAFFLE_ID` | ID de la rifa anterior asociada a `ADMIN_PIN`. Opcional si existe una sola rifa anterior; obligatorio si hay varias. Para la boleta suministrada: `boyaca-31-octubre-2026`. |

3. Ejecuta `prepararSistema` desde el editor y autoriza el acceso a Sheets. Añade las columnas originales faltantes y crea `RifaPro_v2_Solicitudes`, `RifaPro_Cartones` y **`RifaPro_Accesos`**. Genera el único secreto `OWNER_PIN_PEPPER` e importa el PIN anterior a su cartón sin reemplazar credenciales existentes. Añade también la columna `pinLookup` a los accesos anteriores, sin reemplazar sus hashes. **Ejecuta esta función de nuevo al actualizar al ingreso solo con PIN**, aunque ya la hayas ejecutado en una versión anterior. No mueve ni borra las rifas o reservas anteriores. Si encuentra una cabecera personalizada incompatible, se detiene sin sobrescribirla.
4. Las rifas y los números existentes se leen directamente desde `rifas`, sin importar datos a otro formato. Los pagos anteriores cuentan en el recaudo y aparecen como **sin medio registrado** hasta que el administrador elija transferencia o efectivo. Los celulares y los estados reales de pago solo se devuelven en consultas autenticadas.
5. Para publicar la boleta adjunta, revisa y ejecuta explícitamente `crearRifaDeLaBoleta`. Incluye premio **$500.000**, valor **$20.000**, números **00–99**, lotería **Boyacá** y límite de pago **5 p. m.** del **31 de octubre**. El año **2026** se tomó del contexto actual porque no aparece en la boleta; revisa también la transcripción **Arelys B. Mendoza B.** antes de publicarla. Comparte `?rifa=boyaca-31-octubre-2026`; su organizador pulsa **Administrador** e ingresa su PIN. Para vender otros cartones, usa **Crear mi cartón** y apruébalos desde el panel del desarrollador.
6. **Implementar → Administrar implementaciones → Editar → Nueva versión**. Ejecutar como el propietario y permitir acceso a cualquiera. Actualizar la implementación existente conserva su URL `/exec`. Las antiguas acciones de escritura y de lectura privada por GET quedan deshabilitadas en el servidor v2.
7. Si se genera otra URL, cambia `apiUrl` en `config.js`. No agregues contraseñas ni tokens al cliente.
8. Prueba desde dos navegadores independientes: el mismo número debe producir una sola reserva y un conflicto. Abre el enlace del cartón en otro navegador y comprueba que aparezca como participante, sin KPI ni botón de copiar. Entra desde **Administrador** con el PIN de ese cartón. Verifica que no pueda consultar otra rifa ni aprobar un pago, y que el panel global requiera `DEVELOPER_PIN`. Comprueba también **Ya tengo un cartón** desde el inicio: debe pedir solo el PIN y restaurar su propia rifa.

El cliente usa JSONP solamente para consultas públicas y POST con cuerpo `text/plain` para autenticación y escrituras. No trata una respuesta opaca `no-cors` como confirmación. Los tiempos de espera piden reintentar la misma solicitud.

## Interacción y rendimiento

La interfaz usa la tipografía del sistema, sin descargar fuentes. En Apple toma la fuente nativa. Los campos tienen al menos 16 px para evitar el zoom automático al enfocarlos en Safari; no se desactiva el zoom manual. Los diálogos móviles se presentan como paneles inferiores y los controles táctiles amplían su superficie.

Al abrir un cartón aparece **Preparando números**, con tarjetas animadas. Se oculta cuando llega la respuesta, también ante errores o falta de conexión. No tiene un tiempo mínimo ni retrasa la carga para terminar la animación; las consultas periódicas no vuelven a mostrarlo. El movimiento respeta la preferencia de animación reducida. El formulario de creación móvil usa una columna, cabecera y acción fijas, desplazamiento vertical interno y campos de al menos 16 px. Los controles nativos de fecha y hora están limitados al ancho disponible, incluida su presentación en Safari. No se desactiva el zoom manual. **Ya tengo un cartón** tiene mayor contraste y un pulso suave, desactivado si el dispositivo solicita movimiento reducido.

El nuevo `Code.gs` anuncia sus capacidades. Tras la consulta inicial, una sola petición trae catálogo y números (`getSnapshot`) o catálogo, números y resumen privado (`adminSnapshot`). Las escrituras devuelven el registro confirmado y sus totales, sin exigir dos lecturas posteriores. El servidor reutiliza la conexión a Sheets y la matriz leída dentro de cada solicitud.

Los registros administrativos se pueden seguir editando mientras se guarda el anterior. Los cambios acumulados se envían en lotes de hasta 25 con un bloqueo compartido, lectura única y escrituras de rangos contiguos. Cada resultado se confirma o rechaza por separado. El navegador conserva los botones sin reconstruir toda la cuadrícula en cada actualización, para mantener foco y reducir trabajo.

La respuesta visual es inmediata; la confirmación depende de la red, las cuotas y el tiempo de ejecución de Apps Script. La pantalla nunca convierte un cambio pendiente en una reserva o pago confirmado por anticipado.

**Para activar el ingreso solo con PIN, copia `backend/Code.gs`, configura `DEVELOPER_PIN`, ejecuta `prepararSistema` y publica una nueva versión de la misma implementación.** Conserva el `ADMIN_PIN` anterior hasta importarlo. La creación de nuevos cartones requiere la capacidad `pinOnlyLogin` del servidor. Las consultas públicas y reservas existentes siguen disponibles por su enlace durante la actualización.

El bloqueo se implementa con [LockService de Apps Script](https://developers.google.com/apps-script/reference/lock/lock-service). Para volúmenes altos, hay que dimensionar el servicio conforme a las cuotas de Apps Script; esta implementación no se ha sometido a una prueba de carga en producción.

## GitHub Pages

El workflow `.github/workflows/pages.yml` ejecuta las pruebas y publica solo los archivos públicos. En el repositorio, elige **Settings → Pages → Source → GitHub Actions**. El servidor de Apps Script se despliega por separado.

Para conservar el historial remoto, utiliza un clon auténtico en otra carpeta y copia allí los archivos modificados de esta entrega. No inicialices un historial nuevo ni hagas force push sobre el repositorio existente. Después de configurar el servidor y comprobar los datos:

```sh
git clone https://github.com/devnnex/RIFA-PRO.COM.git
# Copiar los archivos de esta entrega al clon, sin .qa, .tools ni node_modules.
cd RIFA-PRO.COM
npm test
git add .
git commit -m "Improve raffle experience, private admin summary and payment methods"
git push origin main
```

El sitio admite la ruta de proyecto `/RIFA-PRO.COM/` y enlaces a una rifa concreta mediante `?rifa=ID`.

## Desarrollo y verificación

No hay dependencias de npm ni proceso de compilación.

```sh
npm test
npm start
```

Abre `http://127.0.0.1:4173`. También se admite `http://127.0.0.1:4173/RIFA-PRO.COM/` para comprobar las rutas de GitHub Pages.

`tests/backend.test.cjs` ejecuta 58 pruebas con dobles de Sheets, caché y bloqueo. Verifica ganador único, conflictos, reintentos, aislamiento público y entre organizadores, credenciales de desarrollador separadas, PIN por cartón, hashes y sales, bloqueo por intentos, importación del PIN anterior, cierre de sesión, liberaciones, medios de pago, edición del premio, lotes, aprobación de cartones, vencimiento y compatibilidad con la hoja original. También cuenta lecturas y escrituras. Estas pruebas verifican la lógica; no sustituyen las pruebas entre clientes contra el despliegue real.

`scripts/browser-qa.cjs` verifica Chrome con el servicio aislado `tests/browser-fixture.cjs`: mide la respuesta visual antes de recibir confirmación, conserva nodos y foco, comprueba la lectura agrupada, conflictos, medios de pago, premio editable y su restauración si se rechaza, creación, eliminación, lotes y reintentos con el mismo ID. Revisa el formulario con fecha vacía y completa a 320/375/390/430 px, además de la página a 320/375/390/768/1024, campos de 16 px, iconos y PWA sin conexión. Las capturas quedan en `.qa/`, que no se publica. La medición visual de esta prueba no representa la latencia del servidor real.

## Iconos PWA

`images/logo.svg` contiene una marca vectorial de boleta con los colores de la referencia. Los PNG de 192 y 512 píxeles, el icono maskable de 512 y el de Apple de 180 están incluidos. `scripts/create-icons.ps1` permite regenerarlos con System.Drawing en Windows.

El service worker instala la interfaz y los iconos como un conjunto y los sirve desde la caché de esa versión para abrir la PWA sin esperar a la red. Cambia su identificador `CACHE` en cada publicación de recursos. Nunca almacena respuestas del servidor. La instalación requiere HTTPS (GitHub Pages) o localhost; las reservas requieren conexión. En iPhone, usa Safari → Compartir → Agregar a inicio.
