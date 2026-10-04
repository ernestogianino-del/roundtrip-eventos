# roundtrip-eventos

Sitio de Sisifuz / Round Trip Eventos (GitHub Pages, dominio `www.roundtrip-eventos.com`).

| Archivo | Qué es |
|---|---|
| `index.html` | Página principal: presentación del sistema Round Trip (invitaciones con QR, eventos, pasarela de pagos, control de promotores y artistas). |
| `cuadre-wapaz.html` | Cuadre de la co-producción Wapaz. Pide un código de acceso y guarda en Google Sheets. No está enlazada desde la página principal y lleva `noindex`. |
| `lista.html` | Lista de invitados de Round Trip (antes era `index.html`). Los links con `?promotor=` de la página principal redirigen aquí. |
| `apps-script/cuadre-wapaz.gs` | Script de Google que recibe el cuadre. Las instrucciones de instalación están al inicio del archivo. |

## Cuadre Wapaz: qué vive dónde

- En este repositorio (público) solo está el código. No hay montos, términos del acuerdo, nombres de la planilla ni el código de acceso.
- En la hoja de Google quedan los cuadres guardados, los términos del acuerdo y la planilla base.
- El código de acceso vive solo en las propiedades del script de Google (`CODIGO`).
- Después de instalar el script, pegar su URL `/exec` en la constante `WEB_APP_URL` de `cuadre-wapaz.html`.
