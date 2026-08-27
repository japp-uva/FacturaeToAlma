# Changelog

Todas las modificaciones relevantes del proyecto **FacturaeToAlma** se documentan en este archivo.

El formato sigue una estructura sencilla basada en versiones, con indicación de funcionalidades añadidas, cambios, correcciones, seguridad y limitaciones conocidas.

---

## [1.3_dev] - En desarrollo

### Añadido

- Lectura del campo `Quantity for price` del fichero obtenido mediante Alma Analytics.
- Uso de `Quantity for price` para determinar la cantidad total solicitada en cada línea de orden de compra.
- Comparación entre la cantidad facturada y la cantidad solicitada para deducir si la facturación de cada ítem es parcial o total.
- Incorporación de esta información al procesamiento interno de las líneas de factura.
- Identificación de la factura y de la línea correspondiente en los resultados de validación.

### Cambiado

- El fichero de Alma Analytics aporta ahora, además de los datos utilizados para localizar la `PO Line`, la cantidad solicitada mediante el campo `Quantity for price`.
- La aplicación puede distinguir internamente entre:
  - facturación parcial, cuando la cantidad facturada es inferior a `Quantity for price`;
  - facturación total, cuando la cantidad facturada alcanza o supera `Quantity for price`.
- Se mantiene un único fichero de Alma Analytics por conversión, correspondiente a una misma biblioteca.
- Los lotes continúan limitados a facturas del mismo proveedor y de la misma biblioteca.

### Importante

- La deducción de facturación parcial o total se realiza mediante la comparación de cantidades.
- Esta información puede utilizarse para diagnóstico, validación o revisión posterior.
- La plantilla Excel actual de Alma no dispone de un campo que permita importar directamente el estado de facturación parcial o completa.
- Por este motivo, la aplicación puede deducir e informar el estado, pero no transmitirlo automáticamente a Alma mediante la plantilla de carga.

### Limitaciones conocidas

- La plantilla Excel actual de Alma no permite indicar la opción `Line Exclusive` o «Línea exclusiva».
- La plantilla Excel actual de Alma no permite indicar directamente si una línea está parcial o completamente facturada.
- Las limitaciones anteriores han sido confirmadas por Soporte de Ex Libris.
- La deducción basada en `Quantity for price` depende de que el campo esté informado correctamente en Alma Analytics.

---

## [1.2_dev] - Versión de desarrollo

### Añadido

- Uso de `defusedxml` para reforzar la seguridad en el procesamiento de ficheros XSIG/XML.
- Protección frente a XML maliciosos con entidades externas, expansión de entidades y estructuras potencialmente peligrosas.
- Sanitización de textos antes de escribirlos en el Excel de Alma y en el informe de validación.
- Protección frente a cadenas de texto que comiencen por caracteres interpretables por Excel como fórmulas, como `=`, `+`, `-` o `@`.
- Mensaje de error específico cuando un XML es bloqueado por motivos de seguridad.

### Cambiado

- Se mantiene la lógica funcional de `1.1_dev`.
- Se prioriza la seguridad del procesamiento de ficheros sin alterar la estructura del Excel generado para Alma.
- Se mantiene `tkinter/ttk` como interfaz gráfica.
- Se mantiene un único fichero Excel de Alma por conversión.

### Seguridad

- Sustitución del parser XML estándar por `defusedxml.ElementTree`.
- Prevención de ataques de expansión de entidades y estructuras XML peligrosas.
- Prevención de fórmulas inyectadas desde títulos, nombres de proveedores u otros textos procedentes de las facturas.
- Normalización de la ruta de salida mediante `os.path.abspath()` antes de abrir la carpeta.
- Comprobación mediante `os.path.isdir()` de que la ruta de salida corresponde a un directorio existente.

### Dependencias externas

Para ejecutar el código fuente se requieren:

```text
openpyxl
defusedxml
```

Instalación:

```bash
python -m pip install openpyxl defusedxml
```

### Pendiente de evaluación

- Ejecución de conversiones en un hilo secundario para evitar el bloqueo visual de la interfaz en lotes grandes.
- Revisión de la conversión de importes `Decimal` a valores numéricos compatibles con Excel y Alma.
- Mejora diferenciada de los mensajes de error según el tipo de fallo.
- Incorporación futura de un apartado de ayuda o ventana «Acerca de».

---

## [1.1_dev] - Versión de desarrollo

### Añadido

- Modo de conversión por lotes.
- Selección múltiple de facturas XSIG/XML/TXT para generar un único Excel de carga en Alma.
- Validación de que todas las facturas del lote pertenecen al mismo proveedor.
- Restricción funcional del modo lote a facturas del mismo proveedor y de la misma biblioteca.
- Uso de un único fichero Excel de Alma por conversión, tanto en modo individual como en modo lote.
- Escritura de varias facturas consecutivas en el Excel final siguiendo esta estructura, sin filas en blanco entre facturas:

```text
HINV
INV
HIL
IL
IL
HINV
INV
HIL
IL
...
```

- Inclusión del número de factura en el informe de validación.
- Extracción más robusta del proveedor desde `SellerParty`.
- Soporte para proveedores definidos como empresa o como persona física/autónomo.
- Búsqueda del nombre del proveedor en este orden:
  1. `LegalEntity > CorporateName`
  2. `Individual > CorporateName`
  3. `LegalEntity > TradeName`
  4. `Individual > TradeName`
  5. `Individual > Name + FirstSurname + SecondSurname`
- Uso interno de `TaxIdentificationNumber` para validar que las facturas del lote pertenecen al mismo proveedor.
- Mensajes de interfaz adaptados al modo lote.
- Mención explícita en la interfaz de que los lotes deben corresponder al mismo proveedor y a la misma biblioteca.

### Cambiado

- La aplicación pasa a trabajar internamente con una lista de facturas, incluso en modo individual.
- `convert_facturae_to_alma_excel()` acepta una factura o varias.
- El exportador permite escribir varias facturas en un mismo Excel.
- El informe de validación identifica la factura a la que pertenece cada incidencia.
- Se mantiene un único fichero Excel de Alma para evitar mapeos por biblioteca.
- Se descarta la carga simultánea de varios Excel de Alma, ya que dichos ficheros no contienen información suficiente para distinguir bibliotecas.
- El `TaxIdentificationNumber` no se escribe en el Excel de Alma.
- Se descarta el uso de `ttkbootstrap`; la interfaz se mantiene en `tkinter/ttk`.

### Corregido

- Mejora del texto de notas para evitar cortes en la interfaz.
- Ajuste de terminología en la interfaz:
  - se evita `Excel(es)`;
  - se usa `Fichero Excel de Alma`.
- Reducción de textos largos en botones para evitar puntos suspensivos.
- Validación de lotes con proveedores distintos, impidiendo la conversión.

### Limitaciones conocidas

- El modo lote solo debe utilizarse con facturas del mismo proveedor y de la misma biblioteca.
- Cada conversión utiliza un único fichero Excel de Alma.
- La aplicación no puede verificar automáticamente que todas las facturas correspondan a la misma biblioteca, ya que dicha información no está disponible en el fichero Excel de Alma.
- Si una línea no aparece en el Excel de líneas abiertas de Alma, la aplicación no inventa la `PO Line` y genera un aviso de validación.

---

## [1.0] - Versión estable inicial

### Añadido

- Primera versión estable de FacturaeToAlma.
- Conversión de facturas XSIG/XML/TXT a Excel compatible con la carga en Alma.
- Interfaz gráfica en `tkinter/ttk`.
- Selección manual de factura, fichero Excel de Alma, plantilla Excel de Alma y fichero de salida.
- Uso dinámico de la plantilla de Alma para respetar el orden de columnas, la estructura `HINV`, `INV`, `HIL`, `IL` y los estilos básicos.
- Generación del Excel final con estructura reconocida por Alma.
- Generación de informe de validación cuando existen incidencias.
- Enlaces clicables al repositorio del proyecto y a la Biblioguía.
- Botón para limpiar datos.
- Botón para abrir la carpeta de salida.
- Línea de estado y resumen final de conversión.

### Matching con Alma

- Búsqueda de líneas de orden de compra por ISBN e ISSN.
- Búsqueda alternativa por título cuando no hay ISBN ni ISSN.
- Comparación de títulos mediante normalización textual y stopwords multilingües.
- Stopwords integradas para español, inglés, francés, alemán, italiano, catalán, gallego, euskera y portugués.
- Extracción de ISBN/ISSN desde `ArticleCode` e `ItemDescription`.
- Limpieza del título cuando el ISBN/ISSN aparece añadido al final de `ItemDescription`.

### Datos de Alma incorporados

- `PO Line`.
- `Title`.
- `Reporting Code`.
- `Secondary Reporting Code`.
- `Tertiary Reporting Code`.
- `Fourth Reporting Code`.
- `Fifth Reporting Code`.
- `Fund and percent`.
- `Start subs date`.

### Reglas aplicadas

- `PO Line` solo se toma del Excel de Alma y nunca del XSIG/XML.
- Si no hay coincidencia o existen duplicados, `PO Line` queda vacío y se genera validación.
- Si un reporting code contiene `-1`, ese campo concreto se deja vacío.
- `VAT In Invoice Line Level` se informa como `YES`.
- `Report TAX` se deja vacío.
- `Price` se toma de `GrossAmount`, con respaldo prudente en `UnitPriceWithoutTax`.
- El IVA de línea se toma de rutas XML concretas.

### Excel de salida

- Inclusión de desplegable en la columna `Line type` con los valores:
  - `REGULAR`
  - `OVERHEAD`
  - `OTHER`
  - `SHIPMENT`
  - `DISCOUNT`
  - `INSURANCE`
  - `ADDITIONAL_CHARGES`
- Conservación de la estructura de la plantilla de Alma.
- Generación de informe `_validacion.xlsx` para líneas sin coincidencia, duplicados y títulos ambiguos.

### Limitaciones conocidas

- La plantilla Excel actual de Alma no permite indicar la opción `Line Exclusive`.
- La plantilla Excel actual de Alma no permite indicar si una línea está parcial o completamente facturada.
- Estas limitaciones han sido confirmadas por Soporte de Ex Libris.
- En caso de líneas cerradas o ya cargadas en Alma, la aplicación no encontrará `PO Line` si dichas líneas no aparecen en el Excel de líneas abiertas.

---

## Notas generales

### Filosofía del proyecto

FacturaeToAlma automatiza únicamente los emparejamientos seguros.

Si existe duda, duplicidad o ausencia de coincidencia, la aplicación deja la `PO Line` vacía y genera un informe de validación para revisión manual.

### Ramas recomendadas

```text
stable
  Versiones estables validadas con Alma.

dev
  Desarrollo de nuevas funcionalidades y pruebas.
```

### Versión estable actual

```text
1.0
```

### Versión de desarrollo actual

```text
1.3_dev
```
