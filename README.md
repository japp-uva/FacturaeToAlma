# FacturaeToAlma

**FacturaeToAlma** es una aplicación de escritorio desarrollada en Python para convertir facturas electrónicas en formato **XSIG, XML o TXT** a un fichero Excel compatible con la carga de facturas en **Alma**.

La aplicación está orientada inicialmente al tratamiento de facturas procedentes de Universitas XXI y FACe, aunque su diseño pretende facilitar su adaptación y uso por parte de otras bibliotecas e instituciones.

## Estado del proyecto

- **Versión estable:** `1.0`
- **Versión de desarrollo actual:** `1.3_dev`
- **Rama estable:** `stable`
- **Rama de desarrollo:** `dev`

Las versiones marcadas como `dev` deben considerarse versiones de prueba. Se recomienda comprobar cuidadosamente el Excel generado antes de cargarlo en Alma.

## Funcionalidades principales

- Conversión de facturas `XSIG`, `XML` y `TXT` a Excel compatible con Alma.
- Uso de una plantilla Excel de Alma seleccionada por el usuario.
- Conversión de una factura individual o de un lote de facturas.
- Generación de varias facturas consecutivas en un único Excel mediante bloques `HINV`, `INV`, `HIL` e `IL`.
- Validación de que las facturas de un lote correspondan al mismo proveedor.
- Uso de un único fichero de Alma Analytics por conversión y biblioteca.
- Localización de líneas de orden de compra mediante:
  - ISBN;
  - ISSN;
  - título normalizado, cuando no existe una coincidencia válida por identificador.
- Extracción de ISBN e ISSN desde `ArticleCode` y `ItemDescription`.
- Eliminación del ISBN o ISSN del título cuando aparece añadido al final de `ItemDescription`.
- Normalización de títulos mediante stopwords en varios idiomas.
- Incorporación de datos procedentes de Alma Analytics:
  - `PO Line`;
  - `Title`;
  - `Reporting Code`;
  - `Secondary Reporting Code`;
  - `Tertiary Reporting Code`;
  - `Fourth Reporting Code`;
  - `Fifth Reporting Code`;
  - `Fund and percent`;
  - `Start subs date`;
  - `Quantity for price`.
- Desplegable de tipos de línea en la columna `Line type`.
- Generación de un informe Excel de validación cuando existen incidencias.
- Protección frente a estructuras XML potencialmente peligrosas mediante `defusedxml`.
- Protección frente a la interpretación de textos no confiables como fórmulas de Excel.

## Versiones

### 1.0

Primera versión estable del proyecto.

Incluye la conversión individual, búsqueda de PO Lines por ISBN, ISSN o título, lectura dinámica de la plantilla de Alma, fondos, reporting codes, IVA por línea, validaciones y una interfaz gráfica basada en Tkinter/ttk.

### 1.1_dev

Añade el procesamiento por lotes.

El modo lote permite seleccionar varias facturas siempre que correspondan:

- al mismo proveedor;
- a la misma biblioteca;
- al mismo fichero de Alma Analytics.

El informe de validación identifica la factura concreta a la que pertenece cada incidencia.

### 1.2_dev

Añade mejoras de seguridad y robustez:

- uso de `defusedxml` para procesar XML no confiable;
- sanitización de textos escritos en Excel;
- validación y normalización de la ruta usada para abrir la carpeta de salida.

### 1.3_dev

Añade el control de cantidades mediante el campo `Quantity for price` del informe de Alma Analytics.

Cuando la cantidad de una línea de factura no coincide con la cantidad solicitada en la PO Line, la aplicación:

- conserva la PO Line encontrada;
- incorpora la línea al Excel de carga;
- mantiene en `IL > Quantity` la cantidad indicada por la factura;
- genera un aviso `REVISAR CANTIDAD` en el informe de validación.

La diferencia de cantidades puede indicar una posible facturación parcial, aunque también puede corresponder a una entrega final que complete una línea facturada anteriormente. La aplicación no toma una decisión automática y solicita revisión manual.

Para conocer todos los cambios, consulta [`CHANGELOG.md`](CHANGELOG.md).

## Requisitos para ejecutar el código fuente

### Python

Se necesita una instalación reciente de Python 3 con soporte para Tkinter.

Para comprobar que Tkinter está disponible:

```bash
python -m tkinter
```

### Dependencias externas

Instala las dependencias mediante:

```bash
python -m pip install openpyxl defusedxml
```

Paquetes utilizados:

- `openpyxl`: lectura y escritura de ficheros Excel.
- `defusedxml`: procesamiento seguro de ficheros XML y XSIG.

El resto de módulos utilizados pertenecen a la biblioteca estándar de Python.

## Ejecución desde el código fuente

```bash
python facturae_to_alma_1_3_dev.py
```

Antes de ejecutar, puede verificarse la sintaxis con:

```bash
python -m py_compile facturae_to_alma_1_3_dev.py
```

## Uso de la aplicación

### Factura individual

1. Selecciona **Factura individual**.
2. Elige el fichero de factura `XSIG`, `XML` o `TXT`.
3. Selecciona el fichero Excel exportado desde Alma Analytics.
4. Selecciona la plantilla Excel de carga de facturas de Alma.
5. Revisa o modifica la ruta del Excel de salida.
6. Pulsa **Convertir a formato Alma**.
7. Revisa el mensaje final y, si se genera, el fichero de validación.

### Lote de facturas

1. Selecciona **Lote de facturas del mismo proveedor y biblioteca**.
2. Selecciona varias facturas mediante el diálogo de selección múltiple.
3. Selecciona un único fichero Excel de Alma Analytics.
4. Selecciona la plantilla Excel de Alma.
5. Define el fichero Excel de salida.
6. Pulsa **Convertir a formato Alma**.

El lote debe contener exclusivamente facturas:

- del mismo proveedor;
- de la misma biblioteca;
- compatibles con el único fichero de Alma Analytics seleccionado.

Si se detectan proveedores distintos, la conversión se cancela.

## Fichero de Alma Analytics

La versión `1.3_dev` espera las siguientes columnas:

- `PO Line`
- `Reporting Code`
- `Secondary Reporting Code`
- `Tertiary Reporting Code`
- `Start subs date`
- `Title`
- `ISSN`
- `ISBN`
- `Quantity for price`

También puede utilizar, cuando estén disponibles:

- `Fourth Reporting Code`
- `Fifth Reporting Code`
- una o varias columnas `Fund and percent`

`Quantity for price` corresponde al número de ejemplares solicitado en la línea de orden de compra y se utiliza únicamente como control de validación.

## Estructura del Excel generado

Cada factura se escribe mediante la siguiente estructura:

```text
HINV
INV
HIL
IL
IL
...
```

En un lote, el siguiente bloque comienza inmediatamente después de la última línea `IL`, sin filas vacías:

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

## Criterios de emparejamiento

La aplicación intenta localizar la PO Line en este orden:

1. ISBN.
2. ISSN.
3. Título normalizado.

La PO Line se toma exclusivamente del fichero Excel de Alma Analytics. Nunca se toma del XSIG o XML.

Si no existe una coincidencia suficientemente segura, la PO Line se deja vacía y se genera una incidencia para revisión manual.

## Informe de validación

Cuando existen incidencias, se crea un fichero con el sufijo:

```text
_validacion.xlsx
```

El informe puede incluir:

- factura y línea afectadas;
- PO Line encontrada o candidatas;
- ISBN e ISSN;
- título;
- cantidad facturada;
- `Quantity for price` de Alma;
- tipo de incidencia;
- mensaje explicativo.

Tipos habituales de incidencia:

- `SIN COINCIDENCIA`
- `DUPLICADO ISBN`
- `DUPLICADO ISSN`
- `TITULO AMBIGUO`
- `REVISAR CANTIDAD`

## Tipos de línea

El Excel generado incorpora un desplegable en `Line type` con los siguientes valores:

- `REGULAR`
- `OVERHEAD`
- `OTHER`
- `SHIPMENT`
- `DISCOUNT`
- `INSURANCE`
- `ADDITIONAL_CHARGES`

El valor predeterminado es `REGULAR`.

## Tratamiento del IVA

La aplicación utiliza actualmente:

```text
Inclusive = False
VAT In Invoice Line Level = YES
Report TAX = vacío
```

Los importes y porcentajes de IVA se incorporan en las líneas cuando están disponibles en la factura.

## Limitaciones conocidas de Alma

Soporte de Ex Libris ha confirmado que la plantilla Excel actual de Alma no permite:

- indicar el tipo de IVA `Line Exclusive` o «Línea exclusiva»;
- indicar explícitamente si una línea queda parcial o completamente facturada.

Por este motivo, FacturaeToAlma no puede informar directamente esas opciones en el Excel de carga.

La comparación entre `Quantity` y `Quantity for price` es solo una ayuda para detectar casos que requieren revisión manual. No determina automáticamente si una factura es parcial o completa.

## Seguridad

La aplicación incluye las siguientes medidas:

- procesamiento XML mediante `defusedxml`;
- bloqueo de estructuras XML peligrosas;
- sanitización de textos que comienzan por caracteres interpretables como fórmulas;
- apertura de carpetas sin `shell=True`;
- normalización y comprobación de rutas antes de abrirlas;
- registro de errores en `facturae_alma.log`.

## Ejecutable para Windows

El proyecto puede distribuirse como ejecutable generado con PyInstaller.

Un ejecutable correctamente empaquetado puede ejecutarse en Windows sin instalar Python, `openpyxl` ni `defusedxml`, porque estas dependencias se incluyen en el paquete.

Ejemplo de empaquetado:

```bash
python -m PyInstaller --clean --onefile --windowed --noupx \
  --name "FacturaeToAlma_1_3_dev" \
  --collect-submodules=openpyxl \
  --collect-data=openpyxl \
  --collect-all=defusedxml \
  facturae_to_alma_1_3_dev.py
```

En Windows CMD puede ser necesario escribir el comando en una sola línea o sustituir `\` por `^`.

Algunos antivirus pueden señalar falsos positivos en ejecutables generados con PyInstaller. Para mejorar la confianza en la distribución se recomienda:

- publicar también el código fuente;
- distribuir el ejecutable desde GitHub Releases;
- publicar su hash SHA-256;
- evitar UPX;
- firmar digitalmente el ejecutable, si se dispone de certificado.

## Filosofía del proyecto

FacturaeToAlma automatiza únicamente los emparejamientos considerados suficientemente seguros.

Cuando existe duda, duplicidad o ausencia de coincidencia, la aplicación evita inventar datos y genera un informe de validación para revisión manual.

## Licencia

Este proyecto se distribuye bajo la licencia **GNU General Public License, versión 2.0 únicamente** (`GPL-2.0-only`).

Consulta el fichero de licencia del repositorio para obtener el texto completo.

## Enlaces

- Repositorio: <https://github.com/japp-uva/FacturaeToAlma>
- Biblioguía: <https://biblioguias.uva.es/facturae_to_alma>

## Contribuciones e incidencias

Las propuestas de mejora, facturas con estructuras no contempladas y errores de conversión pueden comunicarse mediante las incidencias del repositorio.

Al informar de un problema, evita publicar datos personales, fiscales o información sensible contenida en facturas reales. Siempre que sea posible, utiliza ejemplos anonimizados.
