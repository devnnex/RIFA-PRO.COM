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

Cualquier persona puede elegir **Crear mi cartón**. Completa su rifa y celular; el servidor registra una solicitud por **$1.000 COP** y entrega acceso privado mediante un enlace con una clave aleatoria. Cada organizador gestiona únicamente su cartón. Las rifas independientes no aparecen en el catálogo general: el enlace público muestra solo esa rifa.

El creador debe guardar o descargar su enlace privado. La clave se envía mediante POST y el servidor conserva únicamente su hash SHA-256. El enlace se recuerda en `sessionStorage` para recuperarlo en la misma pestaña y se elimina al cerrar sesión. Los tokens de sesión permanecen en memoria. El enlace privado usa el fragmento `#acceso=…`, se retira de la barra al abrirlo y nunca se incluye al compartir con participantes. Quien tenga ese enlace puede administrar el cartón; guárdalo de forma privada. Para volver desde otro dispositivo, abre el enlace guardado o pégalo en **Ya tengo un cartón**.

La pantalla de pago ofrece la llave **3208759798**, Nequi / Bre-B, beneficiario **Elkin Rangel**, copia de llave y descarga PNG de una tarjeta con los datos y la referencia. No genera un QR bancario: la imagen original adjunta en el chat no está disponible como archivo en el repositorio y debe incorporarse como archivo original para preservar su contenido.

**Ya hice mi pago** abre WhatsApp al **+57 3246394689**, con el texto sobre el pago del cartón, su referencia y la indicación de adjuntar el comprobante. La persona adjunta la imagen manualmente. Abrir WhatsApp solo marca la solicitud como pendiente de revisión; no verifica transferencias ni activa el cartón.

Después de verificar el ingreso y el comprobante, el administrador de la plataforma entra con su contraseña habitual, pulsa **Actualizar** en **Activar cartones** y aprueba la referencia correspondiente. Solo entonces se habilitan la administración del organizador y el enlace público. El creador vuelve con su enlace privado; si dejó la app abierta, esta consulta el estado y habilita el acceso al recibir la aprobación. El administrador de plataforma conserva acceso a las rifas anteriores y puede crear rifas propias como antes.

Un cartón vence **al terminar el segundo día calendario después del sorteo, en America/Bogota**. Por ejemplo, un sorteo del 31 de octubre queda disponible hasta el 2 de noviembre a las 23:59:59; cierra el 3 de noviembre a las 00:00. El servidor impide nuevas reservas y ediciones de números o premios después del vencimiento, incluso desde clientes desactualizados. Los registros se conservan para consulta; no se borran automáticamente.

## Configurar el servidor

1. En el proyecto de Apps Script vinculado a tu hoja original, respalda el código anterior y copia el nuevo `backend/Code.gs` completo. Se conserva la pestaña `rifas`, sus 13 columnas originales y las filas `__META__`.
2. Copia `backend/Code.gs`. En **Configuración del proyecto → Propiedades de secuencia de comandos**, agrega:

   | Propiedad | Valor |
   | --- | --- |
   | `SPREADSHEET_ID` | Opcional si ejecutas `prepararSistema` desde el proyecto vinculado a la hoja original; esa función guarda su ID automáticamente. Requerido para proyectos independientes. |
   | `ADMIN_PIN` | La contraseña de cuatro dígitos solicitada por el organizador, configurada únicamente en el servidor |

3. Ejecuta `prepararSistema` desde el editor y autoriza el acceso a Sheets. Añade al final cuatro columnas: `paymentMethod`, `revision`, `requestId` y `condiciones`, si faltan. Crea las pestañas auxiliares `RifaPro_v2_Solicitudes` para reintentos y `RifaPro_Cartones` para solicitudes, hashes de acceso y aprobaciones. **Ejecuta esta función de nuevo al actualizar a cartones independientes**, aunque ya la hayas ejecutado en una versión anterior. No mueve ni borra las rifas o reservas anteriores. Si encuentra una cabecera personalizada incompatible, se detiene sin sobrescribirla.
4. Las rifas y los números existentes se leen directamente desde `rifas`, sin importar datos a otro formato. Los pagos anteriores cuentan en el recaudo y aparecen como **sin medio registrado** hasta que el administrador elija transferencia o efectivo. Los celulares y los estados reales de pago solo se devuelven en consultas autenticadas.
5. Para publicar la boleta adjunta, revisa y ejecuta explícitamente `crearRifaDeLaBoleta`. Incluye premio **$500.000**, valor **$20.000**, números **00–99**, lotería **Boyacá** y límite de pago **5 p. m.** del **31 de octubre**. El año **2026** se tomó del contexto actual porque no aparece en la boleta; revisa también la transcripción **Arelys B. Mendoza B.** antes de publicarla. Para otras rifas, usa el formulario del administrador.
6. **Implementar → Administrar implementaciones → Editar → Nueva versión**. Ejecutar como el propietario y permitir acceso a cualquiera. Actualizar la implementación existente conserva su URL `/exec`. Las antiguas acciones de escritura y de lectura privada por GET quedan deshabilitadas en el servidor v2.
7. Si se genera otra URL, cambia `apiUrl` en `config.js`. No agregues contraseñas ni tokens al cliente.
8. Prueba desde dos navegadores independientes: el mismo número debe producir una sola reserva y un conflicto. Verifica el ingreso con la contraseña configurada, los pagos de ambos medios y que las consultas públicas no incluyan teléfonos ni recaudo.

El cliente usa JSONP solamente para consultas públicas y POST con cuerpo `text/plain` para autenticación y escrituras. No trata una respuesta opaca `no-cors` como confirmación. Los tiempos de espera piden reintentar la misma solicitud.

## Interacción y rendimiento

La interfaz usa la tipografía del sistema, sin descargar fuentes. En Apple toma la fuente nativa. Los campos tienen al menos 16 px para evitar el zoom automático al enfocarlos en Safari; no se desactiva el zoom manual. Los diálogos móviles se presentan como paneles inferiores y los controles táctiles amplían su superficie.

El nuevo `Code.gs` anuncia sus capacidades. Tras la consulta inicial, una sola petición trae catálogo y números (`getSnapshot`) o catálogo, números y resumen privado (`adminSnapshot`). Las escrituras devuelven el registro confirmado y sus totales, sin exigir dos lecturas posteriores. El servidor reutiliza la conexión a Sheets y la matriz leída dentro de cada solicitud.

Los registros administrativos se pueden seguir editando mientras se guarda el anterior. Los cambios acumulados se envían en lotes de hasta 25 con un bloqueo compartido, lectura única y escrituras de rangos contiguos. Cada resultado se confirma o rechaza por separado. El navegador conserva los botones sin reconstruir toda la cuadrícula en cada actualización, para mantener foco y reducir trabajo.

La respuesta visual es inmediata; la confirmación depende de la red, las cuotas y el tiempo de ejecución de Apps Script. La pantalla nunca convierte un cambio pendiente en una reserva o pago confirmado por anticipado.

**Para activar los cartones independientes y el vencimiento, copia `backend/Code.gs`, ejecuta `prepararSistema` y publica una nueva versión de la misma implementación.** La creación pública se habilita únicamente cuando el servidor anuncia `independentRaffles`. Mientras actualizas el servidor, las rifas anteriores siguen funcionando con sus operaciones disponibles.

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

`tests/backend.test.cjs` ejecuta 42 pruebas con dobles de Sheets, caché y bloqueo. Verifica ganador único, conflictos, reintentos, aislamiento público y entre organizadores, contraseña, cierre de sesión, límites, liberaciones, medios de pago, edición del premio, lotes, aprobación de cartones, vencimiento y compatibilidad con la hoja original. También cuenta lecturas y escrituras. Estas pruebas verifican la lógica; no sustituyen las pruebas entre clientes contra el despliegue real.

`scripts/browser-qa.cjs` verifica Chrome con el servicio aislado `tests/browser-fixture.cjs`: mide la respuesta visual antes de recibir confirmación, conserva nodos y foco, comprueba la lectura agrupada, conflictos, medios de pago, premio editable y su restauración si se rechaza, creación, eliminación, lotes y reintentos con el mismo ID. Revisa anchos de 320/375/390/768/1024, campos de 16 px, iconos y PWA sin conexión. Las capturas quedan en `.qa/`, que no se publica. La medición visual de esta prueba no representa la latencia del servidor real.

## Iconos PWA

`images/logo.svg` contiene una marca vectorial de boleta con los colores de la referencia. Los PNG de 192 y 512 píxeles, el icono maskable de 512 y el de Apple de 180 están incluidos. `scripts/create-icons.ps1` permite regenerarlos con System.Drawing en Windows.

El service worker instala la interfaz y los iconos como un conjunto y los sirve desde la caché de esa versión para abrir la PWA sin esperar a la red. Cambia su identificador `CACHE` en cada publicación de recursos. Nunca almacena respuestas del servidor. La instalación requiere HTTPS (GitHub Pages) o localhost; las reservas requieren conexión. En iPhone, usa Safari → Compartir → Agregar a inicio.
