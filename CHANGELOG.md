# Changelog

Todas las modificaciones relevantes de **FacturaeToAlma** se documentan en este archivo.

---

## [3.1] - Versión estable

La versión `3.1` mantiene la funcionalidad de `3.0` y corrige los avisos de bajo impacto detectados mediante una auditoría estática con Bandit.

### Seguridad y robustez

- Revisado el uso del módulo estándar `subprocess`.
- Sustituidas las llamadas a `subprocess.Popen()` por `subprocess.run()` en la apertura de carpetas de macOS y Linux.
- Configuración explícita de `shell=False` en las llamadas externas.
- Resolución previa de la ruta absoluta de `open` o `xdg-open` mediante `shutil.which()`.
- Comprobación de que la utilidad externa encontrada utiliza una ruta absoluta antes de ejecutarla.
- Normalización de la carpeta de salida mediante `os.path.abspath()` y `os.path.realpath()`.
- Verificación de que la ruta corresponde a un directorio existente antes de abrirla.
- Conservación de `os.startfile()` en Windows como mecanismo nativo para abrir carpetas.
- Incorporación de anotaciones `# nosec` específicas y justificadas para los avisos revisados de Bandit:
  - `B404`, importación deliberada de `subprocess`;
  - `B606`, uso controlado de `os.startfile()`;
  - `B603`, ejecución sin shell con ejecutable absoluto y argumento validado.
- Las supresiones se limitan a cada identificador concreto y no desactivan genéricamente el análisis de seguridad.

### Corregido

- Sustituido el bloque silencioso `except Exception: pass` utilizado al aplicar el tema gráfico.
- Captura específica de `tk.TclError` cuando el tema `clam` no está disponible.
- Registro del fallo del tema gráfico mediante `logging.warning()`, manteniendo el tema predeterminado como alternativa.
- Eliminados los usos de rutas parciales para ejecutar `open` o `xdg-open`.

### Auditoría

El análisis original de `3.0` notificó siete avisos de severidad baja y ninguno de severidad media o alta:

- `B404`: importación de `subprocess`;
- `B110`: excepción ignorada mediante `pass`;
- `B606`: apertura de recurso mediante `os.startfile()`;
- `B607`: ejecución de utilidades mediante rutas parciales;
- `B603`: ejecución de procesos sin shell pendiente de revisión de argumentos.

La versión `3.1` corrige los patrones mejorables y documenta expresamente los usos residuales considerados necesarios y seguros.

### Conservado sin cambios funcionales

- Conversión individual y por lotes.
- Gestión visual y editable de lotes.
- Detección de rutas y números de factura duplicados.
- Restricción de los lotes a un mismo proveedor y una misma biblioteca.
- Uso de un único fichero Excel de Alma Analytics por conversión.
- Preferencias persistentes e institución personalizable.
- Logo institucional opcional.
- Conversión en segundo plano.
- Emparejamiento por ISBN, ISSN y título normalizado.
- Comparación entre `Quantity` y `Quantity for Pricing`.
- Aviso `REVISAR CANTIDAD`.
- Generación del Excel de Alma y del informe de validación.
- Procesamiento XML mediante `defusedxml`.
- Sanitización de textos antes de escribirlos en Excel.

### Dependencias externas del código fuente

```text
openpyxl
defusedxml
Pillow
```

`shutil` forma parte de la biblioteca estándar de Python y no requiere instalación adicional.

Instalación:

```bash
python -m pip install openpyxl defusedxml pillow
```

---

## [3.0] - Versión estable

- Estabilización de las mejoras introducidas en `2.5_dev`.
- Interfaz organizada en las pestañas **Convertir**, **Personalizar** y **Ayuda**.
- Logo institucional opcional en PNG, JPG o JPEG.
- Conservación de la transparencia en imágenes PNG.
- Redimensionado proporcional del logo hasta 100 píxeles de altura.
- Título situado debajo del logo.
- Institución y directorios personalizables.
- Preferencias persistentes en `%APPDATA%`.
- Gestión visual de facturas en modo lote.
- Conversión en segundo plano.
- Mayor altura de la ventana y del listado de facturas.
- Nombre de salida del lote con el patrón `Lote_[primera factura]_Alma.xlsx`.
- Cambio de licencia a `GPL-3.0-only`.

---

## [2.5_dev] - Versión de desarrollo

- Introdujo la interfaz por pestañas.
- Añadió preferencias persistentes.
- Añadió la lista visual y editable de facturas de lote.
- Incorporó detección de rutas y facturas duplicadas.
- Incorporó conversión en segundo plano.
- Añadió mensajes finales diferenciados según existieran incidencias.

---

## [2.0] - Versión estable

- Estabilización del modo individual y del modo lote.
- Lotes restringidos a un mismo proveedor y una misma biblioteca.
- Uso de un único fichero Excel de Alma Analytics por conversión.
- Extracción robusta de proveedores empresa y personas físicas.
- Procesamiento XML seguro mediante `defusedxml`.
- Protección frente a fórmulas de Excel.
- Lectura de `Quantity for Pricing`.
- Comparación con `Quantity` y generación de `REVISAR CANTIDAD`.

---

## [1.3_dev] - Versión de desarrollo

- Incorporó la lectura correcta de `Quantity for Pricing`.
- Añadió la comparación con `Quantity` de la factura.
- Añadió ambas cantidades al informe de validación.

---

## [1.2_dev] - Versión de desarrollo

- Sustituyó el parser XML estándar por `defusedxml`.
- Añadió protección frente a estructuras XML peligrosas.
- Añadió sanitización de textos escritos en Excel.
- Reforzó la comprobación de rutas.

---

## [1.1_dev] - Versión de desarrollo

- Añadió el modo lote.
- Mejoró la identificación de proveedores.
- Añadió el número de factura al informe de validación.

---

## [1.0] - Versión estable inicial

- Conversión individual de XSIG/XML/TXT a Excel compatible con Alma.
- Uso dinámico de la plantilla de Alma.
- Localización de PO Lines por ISBN, ISSN o título.
- Incorporación de reporting codes, fondos y fechas de suscripción.
- Desplegable de `Line type`.
- Generación del informe de validación.

---

## Limitaciones conocidas de Alma

Soporte de Ex Libris confirmó que la plantilla Excel actual de Alma no permite:

- indicar `Line Exclusive` o «Línea exclusiva»;
- indicar explícitamente si una línea queda parcial o completamente facturada.

La comparación con `Quantity for Pricing` es una ayuda para la revisión y no un estado transmitido a Alma.

## Estado actual

```text
Versión estable: 3.1
Rama estable: stable
Rama de desarrollo: dev
```
