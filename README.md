# roundtrip-eventos

Sitio de Sisifuz / Round Trip Eventos (GitHub Pages, dominio `www.roundtrip-eventos.com`).

| Archivo | Qué es |
|---|---|
| `index.html` | Página principal: presentación del sistema Round Trip (invitaciones con QR, eventos, pasarela de pagos, control de promotores y artistas). |
| `cuadre-wapaz.html` | Cuadre de la co-producción Wapaz. Pide un código de acceso y guarda en Google Sheets. No está enlazada desde la página principal y lleva `noindex`. |
| `lista.html` | Lista de invitados de Round Trip (antes era `index.html`). Los links con `?promotor=` de la página principal redirigen aquí. |
| `cuadre/index.html` | Cuadre de los miércoles de Sisifuz (Titanium) para que Nery cargue cajas, taquilla, totem, planilla, gastos y fotos. Se abre en `/cuadre`, pide un código de acceso y lleva `noindex`. |
| `apps-script/cuadre-sisifuz.gs` | Script de Google que recibe el cuadre de Sisifuz y lo guarda en Drive: carpeta `Cuadre AAAA-MM-DD`, hoja `Datos Nery DD.MM` y subcarpeta `Fotos`. Instrucciones al inicio del archivo. |
| `apps-script/cuadre-wapaz.gs` | Script de Google que recibe el cuadre. Se instala dentro del archivo "Proyección Viernes - Wapaz"; las instrucciones están al inicio del archivo. |

## Cuadre Wapaz: qué vive dónde

- En este repositorio (público) solo está el código. No hay montos, términos del acuerdo, nombres de la planilla ni el código de acceso.
- Todo se guarda en el archivo de Google "Proyección Viernes - Wapaz". El script no modifica ni borra ninguna otra pestaña del archivo:
  - `cuadre web`: una fila por noche con cada ingreso y gasto en su columna. La fila 2 es de TOTALES (suma automática) y los cuadres empiezan en la fila 3. El detalle de la planilla de cada noche queda como nota en la celda "Planilla staff".
  - `Negociación AAAA-MM-DD`: una hoja por cada versión de la negociación (fecha de inicio en el nombre). Cuando el acuerdo cambia se crea una hoja nueva y las anteriores quedan como historial. Cada cuadre usa la negociación vigente a su fecha.
- El código de acceso vive solo en las propiedades del script de Google (`CODIGO`).
- Después de instalar el script, pegar su URL `/exec` en la constante `WEB_APP_URL` de `cuadre-wapaz.html`.

## Cuadre Sisifuz: qué vive dónde

- En este repositorio (público) solo está el código. No hay montos, nombres de la planilla, el código de acceso ni el ID de la carpeta de Drive.
- El código (`CODIGO`) y la carpeta de destino (`CARPETA_ID`) viven solo en las propiedades del script de Google.
- Después de publicar el script, pegar su URL `/exec` en la constante `WEB_APP_URL` de `cuadre/index.html`. Mientras esté vacía, la página solo permite llenar y copiar el resumen.
