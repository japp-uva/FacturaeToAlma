# FacturaeToAlma

**FacturaeToAlma** es una aplicación de escritorio en Python que convierte facturas electrónicas XSIG, XML o TXT a un fichero Excel compatible con la carga de facturas en Alma.

## Estado del proyecto

- **Versión estable actual:** `2.0`
- **Rama estable:** `stable`
- **Rama de desarrollo:** `dev`

## Funcionalidades principales

- Conversión individual o por lotes.
- Lotes limitados a facturas del mismo proveedor y de la misma biblioteca.
- Un único fichero Excel de Alma Analytics por conversión.
- Generación consecutiva de bloques `HINV`, `INV`, `HIL` e `IL`.
- Búsqueda de PO Lines por ISBN, ISSN y título normalizado.
- Extracción de ISBN/ISSN desde `ArticleCode` e `ItemDescription`.
- Soporte para proveedores empresa y autónomos.
- Incorporación de reporting codes, fondos y fechas de suscripción.
- Comparación entre `Quantity` de la factura y `Quantity for Pricing` de Alma Analytics.
- Informe de validación con factura, línea, PO Line, cantidades y mensaje explicativo.
- Procesamiento XML seguro con `defusedxml`.
- Protección frente a textos interpretables como fórmulas de Excel.

## Control de cantidades

`Quantity` es siempre el número de ejemplares indicado en la factura XML y es el valor que se escribe en la línea `IL`.

`Quantity for Pricing` procede de Alma Analytics y se utiliza exclusivamente como control:

- si coincide con `Quantity`, la información es compatible con una facturación total;
- si no coincide, puede tratarse de una facturación parcial o de una última entrega que complete una línea facturada previamente.

Cuando las cantidades no coinciden, la aplicación conserva la PO Line y el `Quantity` de la factura, incorpora la línea al Excel y genera el aviso `REVISAR CANTIDAD`.

## Requisitos para ejecutar el código fuente

Se requiere Python 3 con Tkinter y estas dependencias:

```bash
python -m pip install openpyxl defusedxml
```

Para ejecutar:

```bash
python facturae_to_alma_2_0.py
```

Para comprobar la sintaxis:

```bash
python -m py_compile facturae_to_alma_2_0.py
```

## Uso

### Factura individual

1. Selecciona **Factura individual**.
2. Elige la factura XSIG/XML/TXT.
3. Selecciona el fichero Excel de Alma Analytics.
4. Selecciona la plantilla Excel de Alma.
5. Revisa el destino y pulsa **Convertir a formato Alma**.

### Lote

1. Selecciona **Lote de facturas del mismo proveedor y biblioteca**.
2. Selecciona varias facturas.
3. Selecciona un único fichero Excel de Alma Analytics.
4. Selecciona la plantilla de Alma.
5. Define el Excel de salida y ejecuta la conversión.

La aplicación bloquea lotes con proveedores distintos. La pertenencia a una misma biblioteca es una condición de uso que no puede comprobarse automáticamente con el informe actual.

## Columnas requeridas en Alma Analytics

- `PO Line`
- `Reporting Code`
- `Secondary Reporting Code`
- `Tertiary Reporting Code`
- `Start subs date`
- `Title`
- `ISSN`
- `ISBN`
- `Quantity for Pricing`

También puede utilizar `Fourth Reporting Code`, `Fifth Reporting Code` y columnas `Fund and percent` cuando estén disponibles.

## Informe de validación

Cuando existen incidencias se genera un fichero con el sufijo `_validacion.xlsx`.

Puede incluir estos avisos:

- `SIN COINCIDENCIA`
- `DUPLICADO ISBN`
- `DUPLICADO ISSN`
- `TITULO AMBIGUO`
- `REVISAR CANTIDAD`

## Limitaciones conocidas de Alma

Soporte de Ex Libris ha confirmado que la plantilla Excel actual no permite:

- indicar `Line Exclusive` o «Línea exclusiva»;
- indicar explícitamente si una línea queda parcial o completamente facturada.

Por ello, la comparación de cantidades es una ayuda de validación y no un estado transmitido a Alma.

## Seguridad

- Parseo XML mediante `defusedxml`.
- Bloqueo de estructuras XML peligrosas.
- Sanitización de textos antes de escribirlos en Excel.
- Apertura de carpeta sin `shell=True`.
- Registro de errores en `facturae_alma.log`.

## Empaquetado para Windows

```bash
python -m PyInstaller --clean --onefile --windowed --noupx --name "FacturaeToAlma_2_0" --collect-submodules=openpyxl --collect-data=openpyxl --collect-all=defusedxml facturae_to_alma_2_0.py
```

El ejecutable generado puede ejecutarse en Windows sin instalar Python ni las dependencias, siempre que PyInstaller las haya incluido correctamente.

## Licencia

El proyecto se distribuye bajo **GNU General Public License version 2.0 only** (`GPL-2.0-only`).

## Enlaces

- Repositorio: <https://github.com/japp-uva/FacturaeToAlma>
- Biblioguía: <https://biblioguias.uva.es/facturae_to_alma>

Consulta [`CHANGELOG.md`](CHANGELOG.md) para ver el historial de versiones.
