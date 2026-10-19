# FacturaeToAlma 4.0

Aplicación de escritorio en Python para convertir facturas electrónicas XSIG, XML o TXT (Facturae/FACe, especialmente procedentes de Universitas XXI) a un Excel compatible con la carga de facturas en Alma.

## Estado del proyecto

- **Versión estable:** 4.0
- **Rama recomendada:** `stable`
- **Archivo principal:** `facturae_to_alma_4_0.py`
- **Licencia:** GPL-3.0-only

## Novedades de la versión 4.0

La versión 4.0 integra en una única aplicación:

- conversión individual y por lotes;
- visor de factura y XML;
- edición previa del Excel destinado a Alma;
- guardado de cambios y actualización inmediata del Excel de salida;
- verificación del fichero generado;
- selección múltiple y operaciones de edición masiva;
- menú contextual en la aplicación;
- desplazamiento con la rueda del ratón en las vistas de factura, XML y edición.

## Pestañas

### Convertir

Permite seleccionar:

- una factura o un lote de facturas;
- el Excel de Alma Analytics;
- la plantilla Excel de Alma;
- la ruta del Excel de salida.

El modo lote exige facturas del mismo proveedor y biblioteca.

### Visualizar

Permite consultar la factura original sin modificarla:

- vista legible de la factura;
- vista XML;
- selector de factura en lotes;
- desplazamiento mediante la rueda del ratón.

### Editar

Muestra el borrador del Excel final de Alma, incluyendo la cabecera `INV` y las líneas `IL`.

Campos editables principales:

- `Line type`;
- `PO Line`;
- `Vendor Ref Number`;
- `ISSN`;
- `ISBN`;
- `Title`;
- `Reporting Code`;
- `Secondary Reporting Code`;
- `Tertiary Reporting Code`;
- `Fourth Reporting Code`;
- `Fifth Reporting Code`;
- `Start subs date`;
- `End subs date`;
- `Note`;
- columnas `Fund and percent`;
- moneda, cuenta de proveedor, nota e indicador `Inclusive` de cabecera.

Campos protegidos:

- fecha de factura;
- importe total de cabecera;
- importes de línea;
- cantidades;
- importes de IVA;
- porcentajes de IVA.

La fecha de factura se muestra en formato `MM/DD/AAAA`.

#### Número de factura

El número no se edita directamente. Se ofrecen tres opciones:

- **Original**;
- **Prefijo**;
- **Sufijo**.

Se propone el año en curso entre corchetes como ayuda, aunque puede modificarse o eliminarse.

#### Edición de celdas

- Doble clic para editar.
- `Intro` confirma el valor.
- `Tabulador` confirma el valor y avanza a la siguiente celda editable.
- `Ctrl + clic` y `Mayús + clic` permiten seleccionar varias líneas.
- El menú contextual permite aplicar un mismo valor a varias líneas, rellenar blancos y restaurar los valores originales.
- Los reporting codes y fondos ofrecen sugerencias basadas en los valores existentes en el contexto de la factura o lote.

#### Guardar cambios y actualizar Excel

El botón **Guardar cambios y actualizar Excel**:

1. confirma la celda activa;
2. recoge los cambios visibles;
3. pide confirmación;
4. guarda una copia estable de los cambios;
5. genera un Excel temporal;
6. aplica y verifica los cambios;
7. sobrescribe el Excel indicado en **Convertir** solo si la verificación es correcta.

El XML o XSIG original nunca se modifica.

El Excel final incorpora una hoja de auditoría denominada **Cambios FacturaeToAlma**.

### Personalizar

Permite configurar:

- nombre de la institución;
- directorios habituales;
- logo institucional opcional.

### Ayuda

Incluye instrucciones de uso, licencia, limitaciones conocidas y enlaces del proyecto.

## Funcionamiento de la conversión

1. Seleccionar factura o lote.
2. Seleccionar Alma Analytics.
3. Seleccionar la plantilla de Alma.
4. Definir el Excel de salida.
5. Ejecutar la conversión inicial.
6. Revisar la factura en **Visualizar**.
7. Corregir el borrador en **Editar**.
8. Pulsar **Guardar cambios y actualizar Excel**.
9. Revisar el Excel final y, si procede, el informe de validación.

## Emparejamiento de PO Lines

La prioridad general es:

1. ISBN;
2. ISSN;
3. título.

Para títulos se aplican:

- coincidencia exacta normalizada;
- inclusión segura de palabras completas;
- similitud textual.

Las coincidencias ambiguas no se asignan automáticamente.

## Columnas utilizadas de Alma Analytics

Columnas principales:

- `PO Line`;
- `Reporting Code`;
- `Secondary Reporting Code`;
- `Tertiary Reporting Code`;
- `Start subs date`;
- `Title`;
- `ISSN`;
- `ISBN`;
- `Quantity for Pricing`.

También se utilizan cuando están disponibles:

- `Fourth Reporting Code`;
- `Fifth Reporting Code`;
- columnas `Fund and percent`.

## Validación

Cuando existen incidencias se genera un fichero con el sufijo `_validacion.xlsx`.

Avisos habituales:

- `SIN COINCIDENCIA`;
- `DUPLICADO ISBN`;
- `DUPLICADO ISSN`;
- `TITULO AMBIGUO`;
- `REVISAR CANTIDAD`.

`Quantity for Pricing` se utiliza como control. No modifica automáticamente la cantidad de factura ni la PO Line.

## Dependencias

```bash
python -m pip install openpyxl defusedxml pillow
```

Para la auditoría de seguridad:

```bash
python -m pip install bandit
```

## Ejecución

```bash
python facturae_to_alma_4_0.py
```

Comprobación de sintaxis:

```bash
python -m py_compile facturae_to_alma_4_0.py
```

Auditoría con Bandit:

```bash
python -m bandit -r facturae_to_alma_4_0.py
```

## Empaquetado para Windows

```bat
python -m PyInstaller --clean --onefile --windowed --noupx ^
  --name "FacturaeToAlma_4_0" ^
  --collect-submodules=openpyxl ^
  --collect-data=openpyxl ^
  --collect-all=defusedxml ^
  --collect-all=PIL ^
  facturae_to_alma_4_0.py
```

## Seguridad

- Procesamiento XML mediante `defusedxml`.
- Bloqueo de estructuras XML peligrosas.
- Sanitización de valores antes de escribir en Excel.
- Ejecución externa con `shell=False`.
- Resolución de ejecutables mediante `shutil.which()`.
- Validación de rutas de entrada y salida.
- Conversión en segundo plano.
- Generación transaccional y verificación del XLSX antes de sustituir el destino.
- Registro en `facturae_alma.log`.
- Revisión con Bandit.

## Limitaciones conocidas de Alma

La plantilla Excel actual de Alma no permite:

- indicar `Line Exclusive` o «Línea exclusiva»;
- indicar expresamente si una línea queda parcial o completamente facturada.

## Licencia

GNU General Public License version 3.0 only (`GPL-3.0-only`).

## Historial

Consulta [CHANGELOG_4.0.md](CHANGELOG_4.0.md).
