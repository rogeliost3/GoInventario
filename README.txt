INVENTARIO DE EQUIPOS - PWA

1. Publica esta carpeta en un servidor HTTPS (por ejemplo GitHub Pages).
2. Abre la URL con Chrome para Android.
3. Instala la PWA desde Chrome (Añadir a pantalla de inicio / Instalar aplicación).
4. La primera vez que pulses la cámara, selecciona Documents o una carpeta equivalente.
5. Autoriza cámara y escritura en la carpeta.
6. Selecciona PC/PTT/MON/IMP/MF y escanea Code 128 con el botón rojo.

IMPORTANTE
- La aplicación utiliza BarcodeDetector nativo de Chrome para Code 128.
- Android/Chrome necesita tener disponible la detección de códigos de barras de Google Play Services.
- SheetJS se carga desde CDN. El service worker intenta almacenarlo en caché para uso posterior sin conexión.
- Para garantizar un primer arranque sin Internet, puede sustituirse el script CDN por una copia local de xlsx.full.min.js.
- El archivo diario se llama INVENTARIO-yymmdd.xlsx.
- Cada lectura reescribe el XLSX completo y lo cierra inmediatamente; no depende de un botón Guardar.
- Las lecturas se conservan también en IndexedDB como respaldo local.
