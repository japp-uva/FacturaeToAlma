# Changelog

Todas las modificaciones relevantes de **FacturaeToAlma** se documentan en este archivo.

---

## [3.0] - Versión estable

La versión `3.0` estabiliza las mejoras introducidas en `2.5_dev` y añade personalización institucional mediante un logo opcional.

### Añadido

- Logo institucional opcional en la cabecera de la aplicación.
- Selección del logo desde la pestaña **Personalizar**.
- Compatibilidad con imágenes:
  - PNG, incluida la transparencia;
  - JPG;
  - JPEG.
- Redimensionado proporcional del logo hasta una altura máxima de 100 píxeles.
- Conservación del tamaño original cuando el logo mide menos de 100 píxeles de alto.
- Persistencia de la ruta del logo en el fichero de configuración.
- Botón **Eliminar logo** para retirar la imagen de la cabecera.
- Gestión segura de logos inexistentes, desplazados o no válidos, sin impedir el inicio de la aplicación ni la conversión.
- Dependencia de Pillow para cargar, convertir y redimensionar imágenes.

### Cambiado

- Licencia actualizada a GNU General Public License 3.0 únicamente, identificada mediante `GPL-3.0-only`.
- Cabecera reorganizada para mostrar primero el logo y, debajo, el título de la aplicación.
- Formato del título actualizado a:

```text
[Institución] - FacturaeToAlma 3.0
```

- Altura inicial de la ventana ampliada a 940 píxeles.
- Altura mínima de la ventana ampliada a 820 píxeles.
- La tabla **Facturas del lote** muestra inicialmente diez filas en lugar de seis.
- El nombre propuesto para el Excel generado en modo lote utiliza el patrón:

```text
Lote_[nombre de la primera factura]_Alma.xlsx
```

- La pestaña **Ayuda** muestra la licencia `GPL-3.0-only`.

### Gestión de la ruta del logo

- La aplicación guarda la ruta absoluta del archivo seleccionado, no una copia de la imagen.
- La ruta se guarda normalmente en:

```text
%APPDATA%\FacturaeToAlma\config.json
```

- Si `%APPDATA%` no está disponible, se utiliza un fichero de configuración junto a la aplicación.
- Cuando la imagen se mueve o elimina, la aplicación se inicia sin logo y mantiene operativa la conversión.
- Para distribuir una personalización institucional entre equipos, el logo debe existir en una ruta válida en cada equipo.

### Dependencias externas del código fuente

```text
openpyxl
defusedxml
Pillow
```

Instalación:

```bash
python -m pip install openpyxl defusedxml pillow
```

### Conservado sin cambios funcionales

- Interfaz por pestañas: **Convertir**, **Personalizar** y **Ayuda**.
- Preferencias persistentes de institución y directorios.
- Gestión visual de lotes.
- Detección de rutas y números de factura duplicados.
- Conversión en segundo plano.
- Emparejamiento por ISBN, ISSN y título normalizado.
- Comparación entre `Quantity` y `Quantity for Pricing`.
- Conservación del `Quantity` indicado en la factura XML.
- Aviso `REVISAR CANTIDAD`.
- Estructura `HINV`, `INV`, `HIL` e `IL`.
- Fondos, reporting codes, fechas de suscripción e IVA.
- Desplegable `Line type`.
- Procesamiento XML seguro mediante `defusedxml`.
- Sanitización de textos antes de escribirlos en Excel.

---

## [2.5_dev] - Versión de desarrollo

### Añadido

- Interfaz estructurada mediante las pestañas **Convertir**, **Personalizar** y **Ayuda**.
- Institución y directorios personalizables.
- Persistencia de preferencias en `%APPDATA%`.
- Memoria de los últimos directorios y de la última plantilla.
- Gestión visual del lote mediante `ttk.Treeview`.
- Vista previa del número de factura, proveedor y fichero.
- Funciones para añadir, quitar, ordenar y vaciar facturas del lote.
- Detección de rutas repetidas y números de factura duplicados.
- Conversión en un hilo secundario para evitar el bloqueo de la interfaz.
- Comunicación con Tkinter mediante `queue.Queue` y `root.after()`.
- Mensajes finales diferenciados según existan o no incidencias.
- Clasificación más clara de errores.

### Cambiado

- El botón **Convertir a formato Alma** pasó a mostrarse con mayor tamaño y texto en negrita.
- Los enlaces de GitHub y la Biblioguía pasaron a la pestaña **Ayuda**.
- El modo lote pasó a gestionarse mediante una lista visible y editable.
- La carpeta inicial de los selectores pasó a depender de preferencias y rutas utilizadas anteriormente.

---

## [2.0] - Versión estable

- Estabilización del modo individual y del modo lote.
- Lotes restringidos a un mismo proveedor y una misma biblioteca.
- Uso de un único fichero Excel de Alma Analytics por conversión.
- Extracción robusta del proveedor para empresas y personas físicas.
- Procesamiento seguro mediante `defusedxml`.
- Protección frente a fórmulas de Excel.
- Lectura de `Quantity for Pricing`.
- Comparación entre `Quantity` y `Quantity for Pricing`.
- Generación del aviso `REVISAR CANTIDAD` cuando ambas cantidades no coinciden.

---

## [1.3_dev] - Versión de desarrollo

- Incorporó la lectura correcta de `Quantity for Pricing`.
- Añadió la comparación con `Quantity` de la factura.
- Añadió el aviso `REVISAR CANTIDAD` y ambas cantidades al informe de validación.

---

## [1.2_dev] - Versión de desarrollo

- Sustituyó el parser XML estándar por `defusedxml`.
- Añadió protección frente a estructuras XML peligrosas.
- Añadió sanitización de textos escritos en Excel.
- Reforzó la comprobación de la ruta usada para abrir la carpeta de salida.

---

## [1.1_dev] - Versión de desarrollo

- Añadió el modo lote para facturas del mismo proveedor y la misma biblioteca.
- Permitió escribir varias facturas en un único Excel de carga.
- Mejoró la identificación de proveedores empresa y autónomos.
- Añadió el número de factura al informe de validación.

---

## [1.0] - Versión estable inicial

- Primera versión estable de FacturaeToAlma.
- Conversión individual de XSIG/XML/TXT a Excel compatible con Alma.
- Uso dinámico de la plantilla de Alma.
- Localización de PO Lines por ISBN, ISSN o título.
- Incorporación de reporting codes, fondos y fechas de suscripción.
- Desplegable de `Line type`.
- Generación de informe de validación.

---

## Limitaciones conocidas de Alma

Soporte de Ex Libris confirmó que la plantilla Excel actual de Alma no permite:

- indicar `Line Exclusive` o «Línea exclusiva»;
- indicar explícitamente si una línea queda parcial o completamente facturada.

La comparación con `Quantity for Pricing` es una ayuda para la revisión y no un estado transmitido a Alma.

## Estado actual

```text
Versión estable: 3.0
Rama estable: stable
Rama de desarrollo: dev
```
