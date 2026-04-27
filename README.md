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



## Importante

No cambies nombres de IDs del HTML ni nombres de archivos sin actualizar las rutas.
