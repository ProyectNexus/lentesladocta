# Lentes La Docta · web + panel

- **Tienda:** `https://TU-USUARIO.github.io/lentes-la-docta/`
- **Panel (app):** `https://TU-USUARIO.github.io/lentes-la-docta/panel/`

Las dos se pueden instalar en el celular como app, con ícono propio.
Los datos siguen en tu planilla de Google: GitHub solo muestra las páginas.

## Qué hay en cada carpeta

| Carpeta / archivo | Para qué es |
|---|---|
| `index.html` | La tienda |
| `panel/` | El panel de gestión (pide la clave) |
| `config.js` | **Lo único que completás:** los 2 links de Apps Script |
| `fotos/` | Anteojos de la intro |
| `icons/`, `panel/icons/` | Íconos de las apps y foto de perfil de Instagram |
| `sw.js`, `manifest.webmanifest` | Lo que hace que se instalen como app |
| `apps-script/` | Copia del código de Apps Script (no se usa en GitHub, es respaldo) |

## Puesta en marcha

1. **Apps Script · Panel:** reemplazá `Código.gs` por `apps-script/Panel.gs`.
   Implementar → Gestionar implementaciones → lápiz → Versión: **Nueva versión** → Implementar.
   Acceso: **Cualquier persona**. Copiá el link que termina en `/exec`.
2. **Apps Script · Tienda:** copiá el link `/exec` de su implementación (acceso: Cualquier persona).
3. **config.js:** pegá los dos links entre las comillas (`tiendaApi` y `panelApi`).
4. **GitHub:** creá el repositorio `lentes-la-docta` (público) → *Add file → Upload files* →
   arrastrá **todo lo que está adentro** de esta carpeta → *Commit changes*.
5. **Settings → Pages →** Source: *Deploy from a branch* → `main` / `(root)` → Save.
   En 1 o 2 minutos queda online.

## Instalar como app

- **Android (Chrome):** abrí el link → menú ⋮ → *Instalar app*.
- **iPhone (Safari):** abrí el link → Compartir → *Agregar a inicio*.

Instalá la tienda y el panel por separado: cada uno tiene su ícono.

## Link para Instagram

`https://TU-USUARIO.github.io/lentes-la-docta/?src=ig-bio`
(el panel cuenta las visitas que llegan desde ahí)

## Cuando haya cambios

Reemplazá los archivos nuevos en GitHub y cambiá `VERSION` en `sw.js`
(por ejemplo `2026-11-03-2`): así los celulares con la app instalada bajan la versión nueva.

> La clave del panel **no** va en GitHub: queda guardada solo en Apps Script.
