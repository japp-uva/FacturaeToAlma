# Changelog

Todas las modificaciones relevantes de **FacturaeToAlma** se documentan en este archivo.

---

## [2.5_dev] - En desarrollo

La versión `2.5_dev` conserva el núcleo de conversión de la versión estable `2.0` y reorganiza de forma importante la interfaz, la gestión de lotes y las preferencias de usuario.

### Añadido

- Interfaz principal estructurada mediante tres pestañas:
  - **Convertir**;
  - **Personalizar**;
  - **Ayuda**.
- Título general situado sobre las pestañas y actualizado dinámicamente cuando se modifica la institución.
- Campo opcional para personalizar el nombre de la institución.
- Configuración opcional de directorios para:
  - facturas FACe en formato XSIG/XML/TXT;
  - fichero Excel de Alma Analytics;
  - plantilla Excel de Alma;
  - ficheros Excel finales.
- Persistencia de preferencias entre sesiones en:

```text
%APPDATA%\FacturaeToAlma\config.json
```

- Alternativa de guardado del fichero de configuración junto a la aplicación cuando `%APPDATA%` no está disponible.
- Memoria de los últimos directorios utilizados para facturas, Alma Analytics, plantilla y salida.
- Memoria de la última plantilla Alma seleccionada.
- Botones **Guardar preferencias** y **Restablecer valores**.
- Pestaña de ayuda con:
  - instrucciones para facturas individuales y lotes;
  - explicación del informe de validación;
  - explicación de `REVISAR CANTIDAD`;
  - limitaciones conocidas de Alma;
  - versión y licencia;
  - acceso al repositorio de GitHub y a la Biblioguía.
- Tabla visual para gestionar las facturas de un lote mediante `ttk.Treeview`.
- Visualización previa en el lote de:
  - número de factura;
  - proveedor detectado;
  - fichero de origen.
- Controles para:
  - añadir facturas;
  - quitar las facturas seleccionadas;
  - ordenar el lote por número de factura;
  - vaciar el lote.
- Detección de rutas de factura repetidas.
- Detección de números de factura duplicados, incluso cuando proceden de ficheros distintos.
- Análisis previo de cada factura añadida al lote para obtener su número y proveedor.
- Ejecución de la conversión en un hilo secundario para evitar que la interfaz se bloquee durante lotes o informes grandes.
- Comunicación segura entre el hilo de conversión y Tkinter mediante `queue.Queue` y `root.after()`.
- Aviso al intentar cerrar la aplicación mientras existe una conversión en curso.
- Clasificación más clara de errores de entrada, formato, permisos y ficheros inexistentes.

### Cambiado

- El botón **Convertir a formato Alma** se muestra con mayor tamaño, relleno y texto en negrita.
- Los enlaces al repositorio y a la Biblioguía dejan de aparecer en la pestaña principal y pasan a **Ayuda**.
- El modo lote deja de ser una selección opaca y pasa a gestionarse mediante una lista visible y editable.
- La carpeta inicial de cada diálogo se determina siguiendo esta prioridad:
  1. directorio personalizado válido;
  2. último directorio válido utilizado;
  3. carpeta relacionada con el paso previo, cuando corresponda;
  4. directorio de la aplicación.
- El Excel de salida se propone junto a la factura seleccionada, salvo que exista un directorio de salida personalizado.
- El cuadro de resultado comienza con un mensaje destacado:
  - `CONVERSION COMPLETA SIN INCIDENCIAS`, cuando no se genera validación;
  - `INCIDENCIAS EN LA CONVERSION`, cuando se genera un informe de validación.
- Las conversiones con incidencias se muestran mediante un cuadro de advertencia y las conversiones limpias mediante un cuadro informativo.
- Se mantiene el resumen detallado de coincidencias y avisos tras la conversión.
- La comprobación final vuelve a verificar los números de factura duplicados antes de iniciar el procesamiento.

### Conservado sin cambios funcionales

- Emparejamiento por ISBN, ISSN y título normalizado.
- Extracción de ISBN/ISSN desde `ArticleCode` e `ItemDescription`.
- Comparación entre `Quantity` y `Quantity for Pricing`.
- Conservación del `Quantity` indicado en la factura XML.
- Aviso `REVISAR CANTIDAD` cuando las cantidades no coinciden.
- Estructura `HINV`, `INV`, `HIL` e `IL` del Excel de salida.
- Fondos, reporting codes, fechas de suscripción e IVA.
- Desplegable `Line type`.
- Procesamiento seguro mediante `defusedxml`.
- Sanitización de textos antes de escribirlos en Excel.
- Uso de un único fichero de Alma Analytics por conversión.
- Restricción de los lotes a un mismo proveedor y una misma biblioteca.

### Pendiente de evaluación

- Comportamiento de la interfaz con lotes especialmente grandes.
- Posibles mejoras del informe de validación, como fichero de origen, tipo de coincidencia y puntuación por título.
- Revisión futura del tratamiento numérico con `Decimal`, únicamente si se detecta un problema real de precisión.

---

## [2.0] - Versión estable

La versión `2.0` estabiliza las funcionalidades desarrolladas y probadas en `1.1_dev`, `1.2_dev` y `1.3_dev`.

### Añadido

- Conversión individual y por lotes.
- Lotes limitados a facturas del mismo proveedor y de la misma biblioteca.
- Uso de un único fichero Excel de Alma Analytics por conversión.
- Escritura consecutiva de varias facturas mediante bloques `HINV`, `INV`, `HIL` e `IL`.
- Identificación de la factura correspondiente en cada incidencia del informe de validación.
- Extracción robusta del proveedor desde `SellerParty`, tanto para empresas como para personas físicas o autónomos.
- Uso interno de `TaxIdentificationNumber` para validar el proveedor de un lote, sin escribirlo en el Excel de Alma.
- Procesamiento seguro de XML mediante `defusedxml`.
- Protección frente a textos interpretables como fórmulas de Excel.
- Lectura de `Quantity for Pricing` en Alma Analytics.
- Comparación entre `Quantity` de la factura y `Quantity for Pricing` de la línea de orden de compra.
- Aviso `REVISAR CANTIDAD` cuando ambas cantidades no coinciden.

### Comportamiento de cantidades

- `IL > Quantity` siempre procede de la factura XML.
- Si `Quantity` coincide con `Quantity for Pricing`, la información es compatible con una facturación total.
- Si no coincide, la aplicación no determina automáticamente si la facturación es parcial o total.
- La diferencia puede corresponder a una facturación parcial o a una última entrega que complete una línea facturada previamente.
- La línea conserva su PO Line y se incorpora al Excel, pero se genera un aviso de revisión.

### Seguridad y robustez

- Uso de `defusedxml.ElementTree`.
- Sanitización de textos antes de escribirlos en Excel.
- Normalización y comprobación de rutas antes de abrir la carpeta de salida.
- Registro técnico en `facturae_alma.log`.

---

## [1.3_dev] - Versión de desarrollo

- Incorporó la lectura correcta del campo `Quantity for Pricing` de Alma Analytics.
- Añadió la comparación con `Quantity` de la factura.
- Añadió el aviso `REVISAR CANTIDAD` y ambas cantidades al informe de validación.
- Sirvió como candidata funcional para la versión estable `2.0`.

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
- Mantuvo un único fichero Excel de Alma Analytics por conversión.

---

## [1.0] - Versión estable inicial

- Primera versión estable de FacturaeToAlma.
- Conversión individual de XSIG/XML/TXT a Excel compatible con Alma.
- Uso dinámico de la plantilla de Alma.
- Localización de PO Lines por ISBN, ISSN o título.
- Extracción de ISBN/ISSN desde `ArticleCode` e `ItemDescription`.
- Stopwords multilingües para comparación de títulos.
- Incorporación de reporting codes, fondos y fechas de suscripción.
- `VAT In Invoice Line Level` informado como `YES`.
- Desplegable de `Line type`.
- Generación de informe de validación.

---

## Limitaciones conocidas de Alma

Soporte de Ex Libris ha confirmado que la plantilla Excel actual de Alma no permite:

- indicar `Line Exclusive` o «Línea exclusiva»;
- indicar explícitamente si una línea queda parcial o completamente facturada.

La comparación con `Quantity for Pricing` es una ayuda de validación y no un estado transmitido a Alma.

## Estado actual

```text
Versión estable: 2.0
Versión de desarrollo: 2.5_dev
Rama estable: stable
Rama de desarrollo: dev
```
