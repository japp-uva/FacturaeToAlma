# FacturaeToAlma 5.1

Aplicación de escritorio en Python para convertir facturas electrónicas XSIG, XML o TXT (Facturae/FACe, especialmente procedentes de Universitas XXI) a un Excel compatible con la carga de facturas en Alma.

## Estado del proyecto

- **Versión estable:** 5.1
- **Rama recomendada:** `stable`
- **Archivo principal:** `facturae_to_alma_5_1.py`
- **Licencia:** GPL-3.0-only

## Novedades de la versión 5.1

La versión 5.1 parte de la versión 5.0 (100 % determinista, local y sin conexión) e incorpora tres mejoras principales orientadas al cuadre exacto en Alma, la interfaz multilingüe y la gestión masiva de proveedores:

- **Gestión exacta de importes e IVA de la versión 4.0**: restablece íntegramente la lógica de extracción y exportación de importes e IVA de `facturae_to_alma_4_0.py` (`InvoiceTotal`, `GrossAmount` / `UnitPriceWithoutTax`, `TaxesOutputs/Tax/TaxAmount/TotalAmount`, `TaxesOutputs/Tax/TaxRate` y `VAT Percentage` vacío en la fila `INV`), garantizando que al cambiar la factura a **«Línea exclusiva»** (*Line Exclusive*) en Alma los importes cuadren siempre sin descuadres por redondeos ni herencia sintética de tipos impositivos.
- **Selector de lenguas oficiales y cooficiales de España (`es`, `ca`, `va`, `gl`, `eu`, `oc`)**: desplegable en **Personalizar → Preferencias** con cambio en tiempo real, persistencia en `config.json` y cobertura completa de toda la interfaz (pestañas, etiquetas de directorios y notas explicativas en **Personalizar**, desplegables de filtros y operadores en **Editar**, barra de estado y de diagnóstico/totales, pestaña **Ayuda**, visor **Visualizar** estilo PDF/DIR3 y todos los diálogos `messagebox` y selectores de archivos).
- **Importación y exportación Excel de proveedores en Personalizar (solo 2 columnas)**: permite exportar e importar en `.xlsx` las asociaciones de proveedores exclusivamente con las columnas **`NIF / CIF`** y **`Vendor Account`** (sin campos de línea de pedido), ofreciendo tres modos de importación:
  - **Incremental / acumulativo** (añadir nuevos y actualizar existentes);
  - **Solo añadir nuevos** (conservar los existentes sin sobrescribirlos);
  - **Carga completa** (reemplazar la lista actual por el contenido del Excel).

Además, conserva todas las prestaciones de productividad introducidas en 5.0:

- lectura rápida en streaming (`read_only=True`) y **caché inteligente en memoria** del Excel de Alma Analytics;
- conversión bidireccional automática entre ISBN-10 e ISBN-13 y detección de `PO Line` explícita (`POL-...`) en el XML;
- emparejamiento bibliográfico por título en 4 niveles (`EXACT`, `CONTAINMENT`, `SIMILARITY` y `DIAGNOSTIC_IDF`);
- respeto estricto a las `PO Lines` introducidas manualmente en **Editar** (`MATCH_MANUAL`), sin sobrescribir `Reporting Codes` ni `Funds`;
- buscador interactivo **«Buscar PO Line en Alma...»** sobre el catálogo de Alma Analytics;
- filtro por estado de línea, **Autofiltro multicriterio estilo Excel** por una o varias columnas y búsqueda rápida global (`Ctrl+F`);
- edición directa al teclear, tirador de relleno tipo Excel (`■`), historial `Deshacer` / `Rehacer` y **borradores de sesión (`.ftalma.json`)**.

## Pestañas

### Convertir

Permite seleccionar:

- una factura individual o un lote de facturas del mismo proveedor y biblioteca;
- el Excel de Alma Analytics con `PO Lines`;
- la plantilla Excel de Alma (`.xlsx`);
- la ruta del Excel de salida.

Hacer doble clic sobre cualquier factura de la lista **Facturas del lote** la abre directamente en la pestaña **Editar**.

### Visualizar

Permite consultar la factura original sin modificarla:

- vista legible estilo PDF (cabecera, emisor, receptor, centros administrativos DIR3, detalle económico de líneas, datos de pago y cuadro de totales, traducidos dinámicamente al idioma activo);
- vista del código XML formateado;
- selector de factura en lotes y desplazamiento con la rueda del ratón.

### Editar

Muestra el borrador del Excel final de Alma, incluyendo la cabecera `INV` y las líneas `IL`.

Campos editables principales:

- `Line type`, `PO Line`, `Vendor Ref Number`, `ISSN`, `ISBN`, `Title`;
- `Reporting Code`, `Secondary Reporting Code`, `Tertiary Reporting Code`, `Fourth Reporting Code`, `Fifth Reporting Code`;
- `Start subs date`, `End subs date`, `Note` y columnas `Fund and percent`;
- moneda (`Currency`), cuenta de proveedor (`Vendor Account`), nota (`Note`) e indicador `Inclusive` de cabecera.

Campos protegidos:

- fecha de factura (`MM/DD/AAAA`), importe total de cabecera, importes de línea, cantidades, importes de IVA y porcentajes de IVA;
- el número de factura se gestiona mediante las opciones **Original**, **Prefijo** y **Sufijo** con actualización en tiempo real.

Incluye:

- **Filtrado por estado y Autofiltro multicriterio estilo Excel** (con preservación del filtro activo aunque se cambie el idioma de la interfaz en vivo);
- **Buscador interactivo «Buscar PO Line en Alma...»** («Asignar solo PO Line» o «Asignar PO Line + códigos y fondos de Analytics»);
- **Barra de diagnóstico en vivo y sumatorio** de Base e IVA de filas visibles y seleccionadas;
- **Guardar y cargar borradores (`.ftalma.json`)** y guardado transaccional verificado en **Guardar cambios y actualizar Excel** (con hoja de auditoría `Cambios FacturaeToAlma`).

### Personalizar

Permite configurar y persistir en `config.json`:

- **Idioma de la interfaz**:
  - Español / Castellano (`es`)
  - Català (`ca`)
  - Valencià (`va`)
  - Galego (`gl`)
  - Euskara (`eu`)
  - Aranés / Occitan (`oc`)
- **Institución**, **directorios habituales** (facturas FACe, Excel de Alma Analytics, plantilla Alma y salida) y **logo institucional** opcional.
- **Proveedores recordados (`NIF / CIF → Vendor Account`)**:
  - alta, actualización, borrado individual y vaciado completo con normalización de NIF/CIF (`normalize_tax_id`);
  - **Importar Excel...** y **Exportar Excel...** exclusivamente con las 2 columnas **`NIF / CIF`** y **`Vendor Account`** (rechazando ficheros con columnas de línea de pedido como `PO Line`, `Reporting Code` o `Fund and percent`), en los modos **Incremental / acumulativo**, **Solo añadir nuevos** y **Carga completa**.

### Ayuda

Incluye instrucciones resumidas de uso, licencia, limitaciones conocidas y enlaces al repositorio y a la Biblioguía, íntegramente localizados en el idioma seleccionado.

## Dependencias

```bash
python -m pip install openpyxl defusedxml pillow
```

Para auditoría de calidad y seguridad:

```bash
python -m pip install bandit ruff
```

## Ejecución

```bash
python facturae_to_alma_5_1.py
```

Comprobación de sintaxis, calidad y seguridad:

```bash
python -m py_compile facturae_to_alma_5_1.py
ruff check --select F,E9,B facturae_to_alma_5_1.py
python -m bandit -r facturae_to_alma_5_1.py
```

## Empaquetado para Windows

```bat
python -m PyInstaller --clean --onefile --windowed --noupx ^
  --name "FacturaeToAlma_5_1" ^
  --collect-submodules=openpyxl ^
  --collect-data=openpyxl ^
  --collect-all=defusedxml ^
  --collect-all=PIL ^
  facturae_to_alma_5_1.py
```

## Seguridad

- Procesamiento XML seguro mediante `defusedxml`.
- Sanitización de celdas frente a inyección de fórmulas en Excel (tanto en el Excel de Alma como en la exportación de proveedores).
- Ejecución externa con `shell=False` y resolución mediante `shutil.which()`.
- Generación transaccional mediante archivo temporal y reemplazo atómico (`os.replace`).
- Auditoría limpia con Bandit (`0 issues`) y Ruff (`0 issues`).

## Limitaciones conocidas de Alma

La plantilla Excel actual de Alma no permite:

- indicar `Line Exclusive` o «Línea exclusiva» directamente desde el fichero de carga;
- indicar expresamente si una línea queda parcial o completamente facturada.

Gracias a la lógica exacta de importes e IVA de la versión 4.0 incorporada en 5.1, una vez cargada la factura en Alma el cambio manual a **«Línea exclusiva»** cuadra siempre con el importe total de la factura.

## Licencia

GNU General Public License version 3.0 only (`GPL-3.0-only`).

## Historial

Consulta [CHANGELOG_5.1.md](CHANGELOG_5.1.md).
