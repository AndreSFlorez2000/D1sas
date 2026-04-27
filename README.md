# Control OC Mantenimiento - Proyecto organizado

Este paquete separa el HTML original en una estructura más profesional sin cambiar la lógica del aplicativo.

## Estructura

```text
index.html
css/
  styles.css
js/
  app.js
  theme.js
  auth.js
assets/
```

## Qué se separó

- `index.html`: estructura principal de la página.
- `css/styles.css`: todos los estilos visuales.
- `js/app.js`: lógica principal, lectura de Excel, filtros, tablas, estadísticas y cruce.
- `js/theme.js`: manejo de tema claro/oscuro.
- `js/auth.js`: login actual temporal.

## Cómo subir a GitHub Pages

1. Entra al repositorio actual.
2. Reemplaza el archivo anterior por estos archivos.
3. Asegúrate de que `index.html` quede en la raíz del repositorio.
4. Sube también las carpetas `css`, `js` y `assets`.
5. Haz clic en **Commit changes**.
6. Espera a que GitHub Pages actualice la página.

## Importante

No cambies nombres de IDs del HTML ni nombres de archivos sin actualizar las rutas.
