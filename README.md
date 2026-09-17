# FacturaeToAlma

**FacturaeToAlma** es una aplicación de escritorio desarrollada en Python para convertir facturas electrónicas XSIG, XML o TXT a un fichero Excel compatible con la carga de facturas en Alma.

La aplicación se orienta inicialmente a facturas procedentes de Universitas XXI y FACe, pero permite personalizar la institución, sus directorios de trabajo y su logo para facilitar el uso por otras bibliotecas.

## Estado del proyecto

- **Versión estable actual:** `3.1`
- **Rama estable:** `stable`
- **Rama de desarrollo:** `dev`
- **Licencia:** GNU General Public License 3.0 únicamente (`GPL-3.0-only`)

## Novedades de la versión 3.1

La versión `3.1` mantiene las funcionalidades de `3.0` y refuerza la seguridad y la trazabilidad a partir de un análisis estático efectuado con Bandit.

### Apertura segura de carpetas

La aplicación normaliza y valida la carpeta de salida antes de abrirla:

- convierte la ruta a absoluta;
- resuelve enlaces y componentes mediante `os.path.realpath()`;
- comprueba que la ruta sea un directorio existente.

En Windows se mantiene `os.startfile()`, que es el mecanismo nativo para abrir la carpeta.

En macOS y Linux:

- se determina la utilidad correspondiente, `open` o `xdg-open`;
- se obtiene su ruta mediante `shutil.which()`;
- se exige una ruta absoluta;
- se ejecuta mediante `subprocess.run()`;
- se utiliza expresamente `shell=False`.

### Mejora del tratamiento de excepciones

Si Tkinter no puede aplicar el tema gráfico `clam`, la aplicación:

- captura específicamente `tk.TclError`;
- registra la incidencia en el log;
- continúa utilizando el tema gráfico disponible.

Se elimina así el antiguo bloque que ignoraba silenciosamente cualquier excepción.

### Anotaciones de Bandit

Los usos revisados de `subprocess` y `os.startfile()` incluyen anotaciones `# nosec` limitadas a comprobaciones concretas:

- `B404`;
- `B606`;
- `B603`.

Estas anotaciones documentan decisiones revisadas y no desactivan globalmente el análisis de seguridad.

## Dependencias

### Dependencias externas

Para ejecutar el código fuente se necesita:

```bash
python -m pip install openpyxl defusedxml pillow
```

Paquetes externos:

- `openpyxl`: lectura y escritura de ficheros Excel;
- `defusedxml`: procesamiento seguro de XML;
- `Pillow`: carga y redimensionado del logo institucional.

### Biblioteca estándar

`shutil` forma parte de la biblioteca estándar de Python. No debe instalarse mediante `pip`.

También pertenecen a la biblioteca estándar módulos como `os`, `sys`, `json`, `queue`, `threading`, `re`, `logging`, `subprocess`, `platform`, `datetime`, `decimal`, `difflib` y `tkinter`.

## Comprobación y ejecución

Comprobación de sintaxis:

```bash
python -m py_compile facturae_to_alma_3_1.py
```

Ejecución:

```bash
python facturae_to_alma_3_1.py
```

Auditoría con Bandit:

```bash
python -m pip install bandit
python -m bandit -r facturae_to_alma_3_1.py
```

Las anotaciones específicas `# nosec` deberían evitar que Bandit vuelva a informar de los usos revisados, sin ocultar otros hallazgos futuros.

## Interfaz

La aplicación se organiza en tres pestañas:

### Convertir

Contiene:

- factura individual;
- modo lote;
- selección del informe de Alma Analytics;
- selección de la plantilla de Alma;
- destino del Excel final;
- conversión y gestión de la salida;
- estado del proceso.

### Personalizar

Permite configurar:

- institución;
- directorio de facturas FACe;
- directorio del fichero Excel de Alma Analytics;
- directorio de plantilla Alma;
- directorio de Excel finales;
- logo institucional opcional.

### Ayuda

Incluye instrucciones, limitaciones conocidas, licencia y enlaces al repositorio y a la Biblioguía.

## Logo institucional

Se admiten:

- PNG, incluida la transparencia;
- JPG;
- JPEG.

El logo:

- aparece encima del título;
- mantiene sus proporciones;
- se reduce hasta un máximo de 100 píxeles de altura;
- no se amplía cuando es más pequeño.

La aplicación guarda la ruta del logo, no una copia de la imagen.

## Preferencias

La configuración se guarda normalmente en:

```text
%APPDATA%\FacturaeToAlma\config.json
```

Puede incluir:

- institución;
- directorios personalizados;
- últimos directorios utilizados;
- última plantilla;
- ruta del logo.

## Uso con una factura individual

1. Abre **Convertir**.
2. Selecciona **Factura individual**.
3. Elige la factura XSIG/XML/TXT.
4. Selecciona el fichero Excel de Alma Analytics.
5. Selecciona la plantilla Excel de Alma.
6. Revisa el destino.
7. Pulsa **Convertir a formato Alma**.
8. Revisa el resultado y el informe de validación, si se genera.

## Uso en modo lote

1. Selecciona **Lote de facturas del mismo proveedor y biblioteca**.
2. Añade las facturas.
3. Revisa número, proveedor y fichero.
4. Elimina u ordena elementos si es necesario.
5. Selecciona un único fichero Excel de Alma Analytics.
6. Selecciona la plantilla.
7. Revisa el nombre de salida propuesto.
8. Ejecuta la conversión.

La aplicación detecta rutas repetidas, números de factura duplicados, facturas ilegibles y proveedores distintos.

El nombre propuesto para un lote sigue este patrón:

```text
Lote_[nombre de la primera factura]_Alma.xlsx
```

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

`IL > Quantity` siempre conserva el número de ejemplares indicado en la factura XML.

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

## Seguridad

- XML procesado con `defusedxml`.
- Bloqueo de estructuras XML peligrosas.
- Sanitización de textos antes de escribirlos en Excel.
- Ejecución externa sin `shell=True`.
- Resolución de rutas de ejecutables con `shutil.which()`.
- Validación de carpetas antes de abrirlas.
- Preferencias JSON sin credenciales.
- Registro de incidencias en `facturae_alma.log`.

## Empaquetado para Windows

```bash
python -m PyInstaller --clean --onefile --windowed --noupx --name "FacturaeToAlma_3_1" --collect-submodules=openpyxl --collect-data=openpyxl --collect-all=defusedxml --collect-all=PIL facturae_to_alma_3_1.py
```

El ejecutable correctamente empaquetado puede ejecutarse sin instalar Python ni las dependencias externas.

## Limitaciones conocidas de Alma

Soporte de Ex Libris confirmó que la plantilla Excel actual no permite:

- indicar `Line Exclusive` o «Línea exclusiva»;
- indicar explícitamente si una línea queda parcial o completamente facturada.

La comparación con `Quantity for Pricing` es una ayuda para la revisión, no un estado transmitido a Alma.

## Licencia

El proyecto se distribuye bajo la **GNU General Public License version 3.0 only** (`GPL-3.0-only`).

La cabecera SPDX del código es:

```text
SPDX-License-Identifier: GPL-3.0-only
```

## Enlaces

- Repositorio: <https://github.com/japp-uva/FacturaeToAlma>
- Biblioguía: <https://biblioguias.uva.es/facturae_to_alma>

Consulta [`CHANGELOG.md`](CHANGELOG.md) para ver el historial de versiones.
