# FacturaeToAlma

**FacturaeToAlma** es una aplicación de escritorio desarrollada en Python para convertir facturas electrónicas XSIG, XML o TXT a un fichero Excel compatible con la carga de facturas en Alma.

La aplicación se orienta inicialmente a facturas procedentes de Universitas XXI y FACe, pero permite personalizar la institución, los directorios de trabajo y el logo para facilitar su utilización por otras bibliotecas.

## Estado del proyecto

- **Versión estable actual:** `3.1`
- **Versión de desarrollo actual:** `3.2_dev`
- **Rama estable:** `stable`
- **Rama de desarrollo:** `dev`
- **Licencia:** GNU General Public License 3.0 únicamente (`GPL-3.0-only`)

Las versiones `dev` deben probarse con facturas reales antes de promoverse a la rama estable.

## Novedades de 3.2_dev

La versión `3.2_dev` mejora la localización de PO Lines por título, especialmente en facturas de:

- bases de datos;
- plataformas electrónicas;
- paquetes de recursos;
- suscripciones;
- servicios con descripciones comerciales extensas.

### Inclusión segura de títulos

Cuando no existe una coincidencia válida por ISBN o ISSN, la aplicación compara títulos en este orden:

1. Coincidencia exacta del título normalizado.
2. Inclusión segura de una secuencia completa de palabras.
3. Similitud textual tradicional.

La normalización ignora:

- mayúsculas y minúsculas;
- tildes y signos diacríticos;
- puntuación y separadores;
- stopwords incluidas en el programa.

Por ejemplo:

```text
MATHSCINET - ONLINE PACKAGE
```

puede relacionarse con:

```text
MathSciNet
```

También se puede reconocer un nombre distintivo como `AENORmás` dentro de una descripción comercial extensa.

### Prevención de falsos positivos

La inclusión se realiza por palabras completas. Por tanto:

```text
art
```

no coincide con:

```text
artificial
```

Para títulos de una sola palabra se exige:

- un mínimo de cinco caracteres;
- que no se trate de un término genérico.

Se excluyen como identificadores únicos términos como:

- `online`;
- `package`;
- `database`;
- `subscription`;
- `platform`;
- `service`;
- y equivalentes en español incluidos en el código.

Si varias PO Lines cumplen la inclusión, la aplicación no selecciona ninguna automáticamente y genera `TITULO AMBIGUO`.

### Prioridad de identificadores

La prioridad general continúa siendo:

1. ISBN.
2. ISSN.
3. Título.

Por ello, la nueva lógica de inclusión no afecta a un libro o revista cuando ya se ha encontrado una coincidencia válida por ISBN o ISSN.

## Dependencias

Para ejecutar el código fuente se requieren:

```bash
python -m pip install openpyxl defusedxml pillow
```

Paquetes externos:

- `openpyxl`: lectura y escritura de Excel;
- `defusedxml`: procesamiento seguro de XML;
- `Pillow`: carga y redimensionado del logo institucional.

`shutil` y el resto de módulos auxiliares utilizados pertenecen a la biblioteca estándar de Python.

## Comprobación y ejecución

Comprobación de sintaxis:

```bash
python -m py_compile facturae_to_alma_3_2_dev.py
```

Ejecución:

```bash
python facturae_to_alma_3_2_dev.py
```

Auditoría con Bandit:

```bash
python -m pip install bandit
python -m bandit -r facturae_to_alma_3_2_dev.py
```

## Interfaz

La aplicación se organiza en tres pestañas:

### Convertir

Contiene la conversión individual, el modo lote, la selección del informe de Alma Analytics, la plantilla, el destino y el estado del proceso.

### Personalizar

Permite configurar la institución, los directorios de trabajo y un logo opcional.

### Ayuda

Incluye instrucciones, limitaciones conocidas, licencia y enlaces al repositorio y a la Biblioguía.

## Uso con una factura individual

1. Abre **Convertir**.
2. Selecciona **Factura individual**.
3. Elige la factura XSIG/XML/TXT.
4. Selecciona el fichero de Alma Analytics.
5. Selecciona la plantilla Excel de Alma.
6. Revisa el destino.
7. Pulsa **Convertir a formato Alma**.
8. Revisa el resultado y el informe de validación, si se genera.

## Uso en modo lote

1. Selecciona **Lote de facturas del mismo proveedor y biblioteca**.
2. Añade las facturas.
3. Revisa número, proveedor y fichero.
4. Elimina u ordena elementos si es necesario.
5. Selecciona un único fichero de Alma Analytics.
6. Selecciona la plantilla.
7. Revisa el nombre de salida.
8. Ejecuta la conversión.

El nombre propuesto para un lote sigue este patrón:

```text
Lote_[nombre de la primera factura]_Alma.xlsx
```

## Criterios de emparejamiento

### ISBN

Es el primer criterio. Cuando existe una coincidencia válida, no se ejecuta la búsqueda por título.

### ISSN

Se utiliza después del ISBN. Cuando existe una coincidencia válida, no se ejecuta la búsqueda por título.

### Título

Si no se encuentra ISBN o ISSN, se aplica:

1. igualdad exacta normalizada;
2. inclusión segura;
3. puntuación de similitud.

Si hay varias candidatas suficientemente próximas, se genera `TITULO AMBIGUO`.

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

Cuando están disponibles, también se utilizan `Fourth Reporting Code`, `Fifth Reporting Code` y columnas `Fund and percent`.

## Control de cantidades

`IL > Quantity` conserva siempre el número de ejemplares indicado en la factura XML.

`Quantity for Pricing` se utiliza exclusivamente como control:

- si coincide con `Quantity`, la información es compatible con una facturación total;
- si no coincide, puede tratarse de una facturación parcial o de una última entrega que complete una línea previamente facturada.

Una diferencia genera `REVISAR CANTIDAD`, pero no modifica la PO Line ni la cantidad de factura.

## Informe de validación

Cuando existen incidencias se genera un fichero con el sufijo `_validacion.xlsx`.

Avisos habituales:

- `SIN COINCIDENCIA`
- `DUPLICADO ISBN`
- `DUPLICADO ISSN`
- `TITULO AMBIGUO`
- `REVISAR CANTIDAD`

Cuando la ambigüedad procede del título, el mensaje indica el método utilizado, por ejemplo `EXACT`, `CONTAINMENT` o `SIMILARITY`.

## Seguridad

- XML procesado mediante `defusedxml`.
- Bloqueo de estructuras XML peligrosas.
- Sanitización de textos antes de escribirlos en Excel.
- Ejecución externa con `shell=False`.
- Resolución de ejecutables mediante `shutil.which()`.
- Validación de carpetas antes de abrirlas.
- Configuración JSON sin credenciales.
- Registro de incidencias en `facturae_alma.log`.

## Logo y preferencias

La configuración se guarda normalmente en:

```text
%APPDATA%\FacturaeToAlma\config.json
```

El logo puede ser PNG, JPG o JPEG. Los PNG conservan la transparencia y la altura máxima visible es de 100 píxeles.

La aplicación guarda la ruta del logo, no una copia de la imagen.

## Empaquetado para Windows

```bash
python -m PyInstaller --clean --onefile --windowed --noupx --name "FacturaeToAlma_3_2_dev" --collect-submodules=openpyxl --collect-data=openpyxl --collect-all=defusedxml --collect-all=PIL facturae_to_alma_3_2_dev.py
```

## Limitaciones conocidas de Alma

Soporte de Ex Libris confirmó que la plantilla Excel actual no permite:

- indicar `Line Exclusive` o «Línea exclusiva»;
- indicar explícitamente si una línea queda parcial o completamente facturada.

La comparación con `Quantity for Pricing` es una ayuda para la revisión y no un estado transmitido a Alma.

## Licencia

El proyecto se distribuye bajo la **GNU General Public License version 3.0 only** (`GPL-3.0-only`).

## Enlaces

- Repositorio: <https://github.com/japp-uva/FacturaeToAlma>
- Biblioguía: <https://biblioguias.uva.es/facturae_to_alma>

Consulta [`CHANGELOG.md`](CHANGELOG.md) para ver el historial de versiones.
