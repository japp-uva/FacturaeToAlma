# FacturaeToAlma

**FacturaeToAlma** es una aplicación de escritorio desarrollada en Python para convertir facturas electrónicas XSIG, XML o TXT a un fichero Excel compatible con la carga de facturas en Alma.

La aplicación se orienta inicialmente a facturas procedentes de Universitas XXI y FACe, pero permite personalizar la institución, sus directorios de trabajo y su logo para facilitar el uso por otras bibliotecas.

## Estado del proyecto

- **Versión estable actual:** `3.0`
- **Rama estable:** `stable`
- **Rama de desarrollo:** `dev`

## Novedades de la versión 3.0

### Personalización mediante logo

La pestaña **Personalizar** permite seleccionar un logo institucional opcional.

Formatos admitidos:

- PNG, incluida la transparencia;
- JPG;
- JPEG.

El logo:

- aparece encima del título de la aplicación;
- mantiene sus proporciones originales;
- se reduce cuando supera los 100 píxeles de altura;
- no se amplía si su altura original es menor;
- no modifica el archivo de imagen original.

El encabezado utiliza este formato:

```text
[LOGO]

[Institución] - FacturaeToAlma 3.0
```

Si el logo se mueve, elimina o deja de ser válido, la aplicación continúa funcionando y muestra únicamente el título.

### Mayor espacio para lotes

La tabla **Facturas del lote** muestra inicialmente diez filas y puede crecer al ampliar la ventana.

### Nombre de salida de los lotes

El nombre propuesto para un Excel generado en modo lote sigue el patrón:

```text
Lote_[nombre de la primera factura]_Alma.xlsx
```

El usuario puede modificar libremente la ruta y el nombre mediante **Elegir destino**.

## Interfaz

La aplicación se organiza en tres pestañas:

### Convertir

Contiene:

- modo factura individual;
- modo lote;
- selección del informe de Alma Analytics;
- selección de la plantilla de Alma;
- destino del Excel final;
- botones de conversión y gestión;
- estado del proceso.

### Personalizar

Permite configurar opcionalmente:

- institución;
- directorio de facturas FACe;
- directorio del fichero Excel de Alma Analytics;
- directorio de plantilla Alma;
- directorio de Excel finales;
- logo institucional.

Incluye los botones:

- **Guardar preferencias**;
- **Restablecer valores**;
- **Seleccionar logo...**;
- **Eliminar logo**.

### Ayuda

Incluye:

- instrucciones básicas;
- explicación del modo lote;
- información sobre el informe de validación;
- explicación de `REVISAR CANTIDAD`;
- limitaciones confirmadas por Ex Libris;
- licencia;
- enlaces a GitHub y a la Biblioguía.

## Gestión de preferencias

La configuración se guarda normalmente en:

```text
%APPDATA%\FacturaeToAlma\config.json
```

El fichero puede incluir:

- institución;
- directorios personalizados;
- últimos directorios utilizados;
- última plantilla seleccionada;
- ruta del logo.

La aplicación guarda la **ruta del logo**, no una copia de la imagen. Por ello, el archivo debe permanecer en una ubicación estable.

Ejemplo:

```text
C:\FacturaeToAlma\recursos\logo.png
```

Si se copia la configuración a otro ordenador, la ruta solo funcionará si el logo existe allí en la misma ubicación. Cada institución puede seleccionar de nuevo su logo desde la pestaña **Personalizar**.

## Funcionalidades principales

- Conversión individual o por lotes.
- Uso dinámico de la plantilla Excel de Alma.
- Generación consecutiva de bloques `HINV`, `INV`, `HIL` e `IL`.
- Localización de PO Lines por ISBN, ISSN y título normalizado.
- Extracción de ISBN/ISSN desde `ArticleCode` e `ItemDescription`.
- Soporte para proveedores empresa y personas físicas o autónomos.
- Incorporación de reporting codes, fondos, fechas de suscripción e IVA.
- Desplegable de `Line type`.
- Comparación entre `Quantity` y `Quantity for Pricing`.
- Informe de validación por factura y línea.
- Ejecución de la conversión en segundo plano.
- Procesamiento XML seguro con `defusedxml`.
- Protección frente a textos interpretables como fórmulas de Excel.

## Requisitos para ejecutar el código fuente

Se necesita Python 3 con Tkinter y estas dependencias externas:

```text
openpyxl
defusedxml
Pillow
```

Instalación:

```bash
python -m pip install openpyxl defusedxml pillow
```

Comprobación de sintaxis:

```bash
python -m py_compile facturae_to_alma_3_0.py
```

Ejecución:

```bash
python facturae_to_alma_3_0.py
```

Para comprobar Tkinter:

```bash
python -m tkinter
```

## Uso con una factura individual

1. Abre la pestaña **Convertir**.
2. Selecciona **Factura individual**.
3. Elige una factura XSIG/XML/TXT.
4. Selecciona el fichero Excel de Alma Analytics.
5. Selecciona la plantilla Excel de Alma.
6. Revisa el destino propuesto.
7. Pulsa **Convertir a formato Alma**.
8. Revisa el mensaje final y el informe de validación, si se genera.

## Uso en modo lote

1. Selecciona **Lote de facturas del mismo proveedor y biblioteca**.
2. Añade las facturas al listado.
3. Revisa el número de factura, el proveedor y el nombre del fichero.
4. Elimina u ordena elementos si es necesario.
5. Selecciona un único fichero Excel de Alma Analytics.
6. Selecciona la plantilla de Alma.
7. Revisa el nombre `Lote_[primera factura]_Alma.xlsx` propuesto para la salida.
8. Ejecuta la conversión.

La aplicación detecta:

- rutas repetidas;
- números de factura duplicados;
- facturas que no puedan analizarse;
- proveedores distintos.

Todos los documentos del lote deben corresponder al mismo proveedor y a la misma biblioteca. La biblioteca no puede verificarse automáticamente con el informe actual de Alma Analytics.

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

## Resultado de la conversión

Sin incidencias, el cuadro final comienza con:

```text
CONVERSION COMPLETA SIN INCIDENCIAS
```

Cuando se genera un informe de validación, comienza con:

```text
INCIDENCIAS EN LA CONVERSION
```

## Tratamiento del IVA

La aplicación utiliza actualmente:

```text
Inclusive = False
VAT In Invoice Line Level = YES
Report TAX = vacío
```

## Limitaciones conocidas de Alma

Soporte de Ex Libris confirmó que la plantilla Excel actual no permite:

- indicar `Line Exclusive` o «Línea exclusiva»;
- indicar explícitamente si una línea queda parcial o completamente facturada.

Por ello, la comparación de cantidades es una ayuda para la revisión y no un estado transmitido a Alma.

## Seguridad

- XML procesado mediante `defusedxml`.
- Bloqueo de estructuras XML peligrosas.
- Sanitización de textos antes de escribirlos en Excel.
- Apertura de carpetas sin `shell=True`.
- Configuración JSON sin contraseñas ni credenciales.
- Registro de errores en `facturae_alma.log`.

## Empaquetado para Windows

Para crear un ejecutable autónomo con PyInstaller:

```bash
python -m PyInstaller --clean --onefile --windowed --noupx --name "FacturaeToAlma_3_0" --collect-submodules=openpyxl --collect-data=openpyxl --collect-all=defusedxml --collect-all=PIL facturae_to_alma_3_0.py
```

El ejecutable correctamente empaquetado puede ejecutarse en Windows sin instalar Python, `openpyxl`, `defusedxml` ni Pillow.

El logo no se integra en el ejecutable, porque es una preferencia configurable. Cada instalación debe conservar el logo en una ruta accesible.

## Licencia

El proyecto se distribuye bajo la **GNU General Public License version 3.0 only** (`GPL-3.0-only`).

La cabecera del código utiliza:

```text
SPDX-License-Identifier: GPL-3.0-only
```

## Enlaces

- Repositorio: <https://github.com/japp-uva/FacturaeToAlma>
- Biblioguía: <https://biblioguias.uva.es/facturae_to_alma>

Consulta [`CHANGELOG.md`](CHANGELOG.md) para ver el historial de versiones.
