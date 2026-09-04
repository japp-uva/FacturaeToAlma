# Changelog

Todas las modificaciones relevantes de **FacturaeToAlma** se documentan en este archivo.

---

## [2.0] - Versión estable

La versión `2.0` estabiliza las funcionalidades desarrolladas y probadas en `1.1_dev`, `1.2_dev` y `1.3_dev`.

### Añadido

- Conversión de una factura individual o de un lote de facturas.
- Selección múltiple de facturas XSIG/XML/TXT en modo lote.
- Validación de que todas las facturas del lote correspondan al mismo proveedor.
- Generación de varias facturas consecutivas en un único Excel mediante bloques `HINV`, `INV`, `HIL` e `IL`, sin filas vacías entre facturas.
- Identificación de la factura correspondiente en cada incidencia del informe de validación.
- Extracción robusta del proveedor desde `SellerParty`, tanto para empresas como para personas físicas o autónomos.
- Búsqueda del nombre del proveedor en `CorporateName`, `TradeName` o mediante la concatenación de `Name`, `FirstSurname` y `SecondSurname`.
- Uso interno de `TaxIdentificationNumber` para comprobar que las facturas de un lote pertenecen al mismo proveedor. Este dato no se escribe en el Excel de Alma.
- Procesamiento seguro de XML mediante `defusedxml`.
- Protección frente a textos que Excel pudiera interpretar como fórmulas.
- Lectura del campo `Quantity for Pricing` del informe de Alma Analytics.
- Comparación entre `Quantity` de la factura y `Quantity for Pricing` de la línea de orden de compra.
- Generación del aviso `REVISAR CANTIDAD` cuando ambas cantidades no coinciden.
- Inclusión en el informe de validación de la PO Line, la cantidad facturada y `Quantity for Pricing`.

### Comportamiento de cantidades

- El valor escrito en `IL > Quantity` siempre procede de la factura XML y nunca se sustituye por `Quantity for Pricing`.
- Si `Quantity` coincide con `Quantity for Pricing`, la información es compatible con una facturación total.
- Si las cantidades no coinciden, la aplicación no determina automáticamente si la facturación es parcial o total.
- La diferencia puede corresponder a una facturación parcial o a una última entrega que complete una línea facturada previamente.
- Cuando existe diferencia, la línea conserva su PO Line, se incorpora al Excel de carga y se genera un aviso para revisión manual.

### Seguridad y robustez

- Uso de `defusedxml.ElementTree` para bloquear expansión de entidades y otras estructuras XML peligrosas.
- Sanitización de textos procedentes de las facturas antes de escribirlos en Excel.
- Normalización y comprobación de la ruta antes de abrir la carpeta de salida.
- Registro de incidencias técnicas en `facturae_alma.log`.

### Restricciones del modo lote

- Todas las facturas deben pertenecer al mismo proveedor.
- Todas las facturas deben corresponder a la misma biblioteca.
- Cada conversión utiliza un único fichero Excel de Alma Analytics.
- La biblioteca no puede comprobarse automáticamente porque el informe de Alma Analytics utilizado no contiene esa información.

### Dependencias externas del código fuente

```text
openpyxl
defusedxml
```

Instalación:

```bash
python -m pip install openpyxl defusedxml
```

### Limitaciones conocidas de Alma

- La plantilla Excel actual de Alma no permite indicar `Line Exclusive` o «Línea exclusiva».
- La plantilla Excel actual de Alma no permite indicar explícitamente si una línea queda parcial o completamente facturada.
- Estas limitaciones han sido confirmadas por Soporte de Ex Libris.

---

## [1.3_dev] - Versión de desarrollo

- Incorporó la lectura del campo `Quantity for Pricing` de Alma Analytics.
- Añadió la comparación con `Quantity` de la factura.
- Añadió el aviso `REVISAR CANTIDAD` y las cantidades al informe de validación.
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

## Filosofía del proyecto

FacturaeToAlma automatiza únicamente los emparejamientos considerados suficientemente seguros. Si existe duda, duplicidad o ausencia de coincidencia, evita inventar datos y genera un informe para revisión manual.

## Estado actual

```text
Versión estable: 2.0
Rama estable: stable
Rama de desarrollo: dev
```
