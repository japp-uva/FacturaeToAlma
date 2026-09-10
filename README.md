# FacturaeToAlma

**FacturaeToAlma** es una aplicación de escritorio desarrollada en Python para convertir facturas electrónicas XSIG, XML o TXT a un fichero Excel compatible con la carga de facturas en Alma.

La aplicación se orienta inicialmente a facturas procedentes de Universitas XXI y FACe, pero permite personalizar la institución y los directorios para facilitar su utilización por otras bibliotecas.

## Estado del proyecto

- **Versión estable:** `2.0`
- **Versión de desarrollo actual:** `2.5_dev`
- **Rama estable:** `stable`
- **Rama de desarrollo:** `dev`

Las versiones `dev` deben probarse cuidadosamente antes de utilizar el Excel resultante en producción.

## Novedades de 2.5_dev

### Interfaz por pestañas

La aplicación se organiza en:

- **Convertir**: conversión individual, gestión de lotes, selección de Alma Analytics, plantilla y salida.
- **Personalizar**: institución, directorios opcionales y preferencias persistentes.
- **Ayuda**: instrucciones, limitaciones conocidas, licencia y enlaces del proyecto.

### Personalización persistente

Se pueden definir opcionalmente:

- institución;
- directorio de facturas FACe;
- directorio del fichero Excel de Alma Analytics;
- directorio de la plantilla Alma;
- directorio de los Excel finales.

La configuración se guarda normalmente en:

```text
%APPDATA%\FacturaeToAlma\config.json
```

Si no se establece un directorio personalizado, la aplicación recuerda el último directorio válido. En una primera ejecución utiliza la carpeta de la aplicación. La salida se propone junto a la factura, salvo que se haya definido un directorio de salida personalizado.

### Gestión visual de lotes

El modo lote muestra una tabla con:

- número de factura;
- proveedor;
- fichero de origen.

Permite:

- añadir facturas;
- quitar las seleccionadas;
- ordenar por número de factura;
- vaciar el lote;
- detectar rutas repetidas;
- detectar números de factura duplicados.

Todos los documentos del lote deben pertenecer al mismo proveedor y a la misma biblioteca. Cada conversión utiliza un único fichero de Alma Analytics.

### Conversión en segundo plano

La conversión se ejecuta en un hilo secundario para evitar que la ventana aparezca bloqueada durante lotes grandes. La interfaz recibe el resultado mediante una cola y actualizaciones programadas con Tkinter.

### Resultado destacado

El cuadro final comienza con uno de estos mensajes:

```text
CONVERSION COMPLETA SIN INCIDENCIAS
```

```text
INCIDENCIAS EN LA CONVERSION
```

Se mantiene el resumen detallado de coincidencias, duplicados, líneas sin PO Line y demás avisos.

## Funcionalidades principales

- Conversión individual o por lotes.
- Uso dinámico de la plantilla Excel de Alma.
- Generación consecutiva de bloques `HINV`, `INV`, `HIL` e `IL`.
- Localización de PO Lines por ISBN, ISSN y título normalizado.
- Extracción de ISBN/ISSN desde `ArticleCode` e `ItemDescription`.
- Soporte para proveedores empresa y autónomos.
- Incorporación de:
  - `PO Line`;
  - título;
  - cinco reporting codes;
  - fondos;
  - fechas de suscripción;
  - datos de IVA.
- Desplegable de `Line type`.
- Comparación entre `Quantity` y `Quantity for Pricing`.
- Informe de validación por factura y línea.
- Procesamiento seguro con `defusedxml`.
- Protección frente a textos interpretables como fórmulas de Excel.

## Requisitos para ejecutar el código fuente

Se necesita Python 3 con Tkinter y estas dependencias externas:

```bash
python -m pip install openpyxl defusedxml
```

Comprobación de sintaxis:

```bash
python -m py_compile facturae_to_alma_2_5_dev.py
```

Ejecución:

```bash
python facturae_to_alma_2_5_dev.py
```

## Uso de la pestaña Convertir

### Factura individual

1. Selecciona **Factura individual**.
2. Elige una factura XSIG/XML/TXT.
3. Selecciona el fichero Excel de Alma Analytics.
4. Selecciona la plantilla Excel de Alma.
5. Revisa el destino propuesto.
6. Pulsa **Convertir a formato Alma**.

### Lote

1. Selecciona **Lote de facturas del mismo proveedor y biblioteca**.
2. Añade las facturas al listado.
3. Revisa número, proveedor y fichero.
4. Quita u ordena elementos si es necesario.
5. Selecciona un único fichero Excel de Alma Analytics.
6. Selecciona la plantilla y define la salida.
7. Ejecuta la conversión.

La aplicación bloquea lotes con proveedores diferentes. La pertenencia a una misma biblioteca es una condición de uso que no puede comprobarbarse automáticamente con el informe actual de Alma Analytics.

## Uso de la pestaña Personalizar

Todos los campos son opcionales.

- **Institución** modifica el título de la aplicación.
- Los directorios personalizados determinan la carpeta inicial de cada selector.
- **Guardar preferencias** conserva los valores entre sesiones.
- **Restablecer valores** devuelve la institución y los directorios a sus valores iniciales.

## Uso de la pestaña Ayuda

Incluye:

- instrucciones básicas;
- explicación del modo lote;
- información sobre el informe de validación;
- significado de `REVISAR CANTIDAD`;
- limitaciones confirmadas por Ex Libris;
- licencia;
- enlaces a GitHub y a la Biblioguía.

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

Cuando están disponibles, también se utilizan:

- `Fourth Reporting Code`
- `Fifth Reporting Code`
- una o varias columnas `Fund and percent`

## Control de cantidades

`IL > Quantity` siempre conserva el número de ejemplares indicado en la factura XML.

`Quantity for Pricing` procede de Alma Analytics y se utiliza exclusivamente como control:

- si coincide con `Quantity`, la información es compatible con una facturación total;
- si no coincide, puede tratarse de una facturación parcial o de una última entrega que complete una línea facturada previamente.

Una diferencia genera `REVISAR CANTIDAD`, pero no modifica la PO Line ni la cantidad de la factura.

## Informe de validación

Cuando existen incidencias se genera un fichero con el sufijo:

```text
_validacion.xlsx
```

Puede incluir:

- factura;
- línea;
- PO Line;
- ISBN e ISSN;
- título;
- cantidad facturada;
- `Quantity for Pricing`;
- candidatas;
- mensaje explicativo.

Avisos habituales:

- `SIN COINCIDENCIA`
- `DUPLICADO ISBN`
- `DUPLICADO ISSN`
- `TITULO AMBIGUO`
- `REVISAR CANTIDAD`

## Tratamiento del IVA

La aplicación utiliza actualmente:

```text
Inclusive = False
VAT In Invoice Line Level = YES
Report TAX = vacío
```

## Limitaciones conocidas de Alma

Soporte de Ex Libris ha confirmado que la plantilla Excel actual no permite:

- indicar `Line Exclusive` o «Línea exclusiva»;
- indicar explícitamente si una línea queda parcial o completamente facturada.

Por ello, la comparación de cantidades es una ayuda de revisión y no un estado transmitido a Alma.

## Seguridad

- XML procesado mediante `defusedxml`.
- Bloqueo de estructuras XML peligrosas.
- Sanitización de textos antes de escribirlos en Excel.
- Apertura de carpetas sin `shell=True`.
- Registro de errores en `facturae_alma.log`.
- Preferencias almacenadas como JSON local, sin contraseñas ni credenciales.

## Empaquetado para Windows

```bash
python -m PyInstaller --clean --onefile --windowed --noupx --name "FacturaeToAlma_2_5_dev" --collect-submodules=openpyxl --collect-data=openpyxl --collect-all=defusedxml facturae_to_alma_2_5_dev.py
```

El ejecutable correctamente empaquetado puede ejecutarse en Windows sin instalar Python ni las dependencias.

## Licencia

El proyecto se distribuye bajo **GNU General Public License version 2.0 only** (`GPL-2.0-only`).

## Enlaces

- Repositorio: <https://github.com/japp-uva/FacturaeToAlma>
- Biblioguía: <https://biblioguias.uva.es/facturae_to_alma>

Consulta [`CHANGELOG.md`](CHANGELOG.md) para ver el historial de versiones.
