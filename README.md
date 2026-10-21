# Rifa Pro

Sitio estático para GitHub Pages con Google Apps Script y Google Sheets como servidor compartido. Diseño crema y azul oscuro inspirado en la boleta, vista móvil, modo oscuro, instalación PWA y múltiples rifas.

## Despliegue del cliente y del servidor

El repositorio contiene el cliente para GitHub Pages y el archivo `backend/Code.gs`, listo para aplicar manualmente en Apps Script. Ambos despliegues son independientes. La interfaz permite consultar el servicio anterior, pero bloquea las escrituras hasta recibir `protocol: 2`; no usa la ruta anterior que permitía sobrescribir números.

El acceso al resumen privado y las nuevas reservas **requieren desplegar `backend/Code.gs`**. Cambiar únicamente los archivos de GitHub Pages no permite hacer privadas las respuestas del servidor anterior.

## Comportamiento

- Los participantes eligen un número libre, ingresan nombre y celular y confirman. La reserva pertenece a la primera solicitud que el servidor acepta dentro de su bloqueo, incluso si llegan dentro del mismo segundo. Abrir el formulario no reserva un número.
- La otra persona recibe un modal con el nombre de quien lo reservó y la hora registrada por el servidor, con segundos y zona horaria de Colombia.
- No se confirman operaciones sin respuesta del servidor ni se reenvían reservas desde una cola sin conexión. Cada solicitud tiene un ID para que un reintento no duplique la operación.
- El servidor valida la contraseña y entrega una sesión de 30 minutos. El token permanece solo en memoria; cerrar sesión o recargar elimina el acceso local. El servidor puede terminar antes una sesión si su caché se elimina.
- Solo el administrador recibe los KPI, los celulares, los estados de pago y sus medios. La respuesta pública agrupa apartados y pagados como reservados, por lo que no revela el recaudo al inspeccionar la red.
- Al marcar **Pagado**, se exige escoger **Nequi / transferencia** (violeta, teléfono) o **Efectivo** (verde, billete). El resumen muestra ambos subtotales. Registrar un medio es un registro manual, no una confirmación bancaria ni una integración con Nequi.
- Las ediciones y liberaciones administrativas usan una revisión para impedir que una pestaña desactualizada sobrescriba cambios de otra.
- El sitio conserva solo los metadatos públicos de las rifas para consulta sin conexión. No guarda teléfonos, medios de pago, totales ni tokens en caché o almacenamiento persistente.

## Configurar el servidor

1. En el proyecto de Apps Script vinculado a tu hoja original, respalda el código anterior y copia el nuevo `backend/Code.gs` completo. Se conserva la pestaña `rifas`, sus 13 columnas originales y las filas `__META__`.
2. Copia `backend/Code.gs`. En **Configuración del proyecto → Propiedades de secuencia de comandos**, agrega:

   | Propiedad | Valor |
   | --- | --- |
   | `SPREADSHEET_ID` | Opcional si ejecutas `prepararSistema` desde el proyecto vinculado a la hoja original; esa función guarda su ID automáticamente. Requerido para proyectos independientes. |
   | `ADMIN_PIN` | La contraseña de cuatro dígitos solicitada por el organizador, configurada únicamente en el servidor |

3. Ejecuta `prepararSistema` desde el editor y autoriza el acceso a Sheets. Añade al final cuatro columnas: `paymentMethod`, `revision`, `requestId` y `condiciones`. También crea una pestaña auxiliar `RifaPro_v2_Solicitudes` para evitar duplicar operaciones al reintentar. No mueve ni borra las rifas o reservas anteriores. Si encuentra una cabecera personalizada incompatible, se detiene sin sobrescribirla.
4. Las rifas y los números existentes se leen directamente desde `rifas`, sin importar datos a otro formato. Los pagos anteriores cuentan en el recaudo y aparecen como **sin medio registrado** hasta que el administrador elija transferencia o efectivo. Los celulares y los estados reales de pago solo se devuelven en consultas autenticadas.
5. Para publicar la boleta adjunta, revisa y ejecuta explícitamente `crearRifaDeLaBoleta`. Incluye premio **$500.000**, valor **$20.000**, números **00–99**, lotería **Boyacá** y límite de pago **5 p. m.** del **31 de octubre**. El año **2026** se tomó del contexto actual porque no aparece en la boleta; revisa también la transcripción **Arelys B. Mendoza B.** antes de publicarla. Para otras rifas, usa el formulario del administrador.
6. **Implementar → Administrar implementaciones → Editar → Nueva versión**. Ejecutar como el propietario y permitir acceso a cualquiera. Actualizar la implementación existente conserva su URL `/exec`. Las antiguas acciones de escritura y de lectura privada por GET quedan deshabilitadas en el servidor v2.
7. Si se genera otra URL, cambia `apiUrl` en `config.js`. No agregues contraseñas ni tokens al cliente.
8. Prueba desde dos navegadores independientes: el mismo número debe producir una sola reserva y un conflicto. Verifica el ingreso con la contraseña configurada, los pagos de ambos medios y que las consultas públicas no incluyan teléfonos ni recaudo.

El cliente usa JSONP solamente para consultas públicas y POST con cuerpo `text/plain` para autenticación y escrituras. No trata una respuesta opaca `no-cors` como confirmación. Los tiempos de espera piden reintentar la misma solicitud.

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

`tests/backend.test.cjs` ejecuta 26 pruebas con dobles de Sheets, caché y bloqueo. Verifica ganador único, conflictos, reintentos, aislamiento público, contraseña, cierre de sesión, límites, liberaciones, medios de pago y compatibilidad con la hoja original de 13 columnas y sus filas `__META__`, incluidos pagos antiguos, fechas, duplicados y otras rifas. Estas pruebas verifican la lógica; no sustituyen las pruebas entre clientes contra el despliegue real.

`scripts/browser-qa.cjs` verifica la interfaz con Chrome en Windows y datos aislados, sin escribir al servicio real: reserva y conflicto, ingreso correcto e incorrecto, selección obligatoria del medio, subtotales, cierre de sesión, anchos de 320/375/390/768/1024, iconos y carga PWA sin conexión. Las capturas quedan en `.qa/`, que no se publica.

## Iconos PWA

`images/logo.svg` contiene una marca vectorial de boleta con los colores de la referencia. Los PNG de 192 y 512 píxeles, el icono maskable de 512 y el de Apple de 180 están incluidos. `scripts/create-icons.ps1` permite regenerarlos con System.Drawing en Windows.

El service worker almacena exclusivamente la interfaz y los iconos. La instalación requiere HTTPS (GitHub Pages) o localhost; las reservas requieren conexión. En iPhone, usa Safari → Compartir → Agregar a inicio.
