# FacturaeToAlma 5.0

Aplicación de escritorio en Python para convertir facturas electrónicas XSIG, XML o TXT (Facturae/FACe, especialmente procedentes de Universitas XXI) a un Excel compatible con la carga de facturas en Alma.

## Estado del proyecto

- **Versión estable:** 5.0
- **Rama recomendada:** `stable`
- **Archivo principal:** `facturae_to_alma_5_0.py`
- **Licencia:** GPL-3.0-only

## Novedades de la versión 5.0

La versión 5.0 mantiene la interfaz clásica 4.0 (100 % determinista, local y sin conexión) e incorpora mejoras clave en extracción XML/IVA, cruce bibliográfico, rendimiento, asistente de catálogo, filtrado y edición estilo Excel:

- extracción avanzada de IVA Facturae/FACe con herencia de porcentaje de cabecera a línea y cálculo automático de cuota;
- lectura rápida en streaming (`read_only=True`) y **caché inteligente en memoria** del Excel de Alma Analytics (por ruta, tamaño y fecha de modificación);
- conversión bidireccional automática entre ISBN-10 e ISBN-13;
- detección directa de `PO Line` explícita (`POL-...`) en campos y referencias del XML;
- clasificación automática de líneas de portes/envío (`SHIPMENT`) y descuentos (`DISCOUNT`) sin falsos avisos de `SIN COINCIDENCIA`;
- emparejamiento bibliográfico por título en 4 niveles (`EXACT`, `CONTAINMENT`, `SIMILARITY` y `DIAGNOSTIC_IDF` con expansión de más de 60 abreviaturas editoriales y título principal MARC 245 / ISBD);
- respeto estricto a las `PO Lines` introducidas a mano en **Editar** (`MATCH_MANUAL`), sin sobrescribir `Reporting Codes` ni `Funds`;
- buscador interactivo **«Buscar PO Line en Alma...»** con consulta en tiempo real sobre el Excel de Alma Analytics y doble acción («Asignar solo PO Line» por defecto, o «Asignar PO Line + códigos y fondos de Analytics»);
- actualización en tiempo real de **Prefijo** y **Sufijo** tanto en la cabecera `INV` como en la columna `Factura` de todas las líneas `IL`, con memoria por factura;
- carga inmediata de la factura al entrar en **Editar**, filtro por estado de línea, **Autofiltro multicriterio por una o varias columnas estilo Excel** y **búsqueda rápida global (`Ctrl+F`)**;
- **barra de diagnóstico en vivo y sumatorio de importes** en **Editar** (motivo exacto del aviso al seleccionar una fila, cuadre contable de factura e importe acumulado de filas visibles y seleccionadas);
- **edición directa al teclear** sobre la celda activa, opciones completas de **arrastre tipo Excel (`■`)** para propagar valores o rellenar blancos, barra de herramientas rápidas y ciclo de vida limpio de editores sin superposiciones visuales;
- **guardar y reanudar borradores de sesión (`.ftalma.json`)** para pausar y continuar facturas largas o lotes más tarde;
- memoria automática y **gestor visual de proveedores (`NIF / CIF → Vendor Account`)** en **Personalizar**, con normalización completa del CIF/NIF;
- control de cuadre contable (`REVISAR CUADRE IMPORTE`), visibilidad permanente de **3. Archivo de salida** en modo lote y salto directo a **Editar** con doble clic en cualquier factura del lote;
- **interfaz multilingüe en lenguas oficiales y cooficiales de España** (Castellano, Català, Valencià, Galego, Euskara y Aranès/Occità) configurable desde **Personalizar → Preferencias** con cambio en vivo y persistencia en `config.json`, 100 % local y sin conexión.

## Pestañas

### Convertir

Permite seleccionar:

- una factura o un lote de facturas;
- el Excel de Alma Analytics;
- la plantilla Excel de Alma;
- la ruta del Excel de salida.

El modo lote exige facturas del mismo proveedor y biblioteca. Hacer doble clic sobre cualquier factura de la lista **Facturas del lote** la abre directamente en la pestaña **Editar**.

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

Se propone el año en curso entre corchetes como ayuda, aunque puede modificarse o eliminarse. El cambio de prefijo o sufijo se refleja en tiempo real tanto en la tabla de cabecera (`INV -> Invoice Number`) como en la columna `Factura` de todas las líneas (`IL`), y se conserva para cada factura del lote al cambiar de documento en el desplegable.

#### Carga automática, filtrado estilo Excel, diagnóstico en vivo y buscador de PO Line

- Al entrar en la pestaña **Editar** con una factura cargada, sus líneas se muestran inmediatamente (y si el Excel de Alma Analytics está seleccionado, se cruzan automáticamente con las `PO Lines` aprovechando la caché en memoria).
- **Filtrado por estado de línea**: permite alternar entre **Todas las líneas**, **Solo líneas con problemas (sin PO Line o con avisos)**, **Solo líneas sin PO Line (pendientes)**, **Solo líneas correctas (con PO Line y sin avisos)** y **Solo líneas modificadas manualmente**, con contador en vivo (`Mostrando | Correctas | Con problemas`) y resaltado visual por estado.
- **Autofiltro multicriterio por una o varias columnas estilo Excel**:
  - Barra superior para añadir filtros acumulativos por cualquier columna (`Contiene / Igual`, `Igual exacto`, `Distinto de`, `Vacías`, `No vacías`) con desplegable de valores únicos existentes.
  - Clic directo en cualquier cabecera de columna (`Columna ▼`) o botón **Filtro Excel...** para abrir el selector estilo Excel con buscador instantáneo y casillas de verificación (`(Seleccionar todo)`, `(Vacías)` y cada valor distinto de la columna).
  - Posibilidad de combinar filtros en **múltiples columnas simultáneamente** y opción de menú contextual **«Filtrar por el valor de esta celda»**.
- **Búsqueda rápida global (`Ctrl+F`)**: cuadro integrado en la barra Excel para filtrar instantáneamente filas que contengan un texto en cualquiera de sus columnas (`Esc` o `✕` limpia la búsqueda).
- **Barra de diagnóstico en vivo y sumatorio**: muestra debajo de la tabla el motivo exacto del aviso de la línea seleccionada (p. ej. `REVISAR CANTIDAD`, `TITULO AMBIGUO`, `DUPLICADO ISBN` o descuadre contable de factura) junto con el sumatorio en tiempo real de la base y el IVA de las filas visibles (y de las filas seleccionadas si hay más de una).
- El botón **Buscar PO Line en Alma...** (también disponible en el menú contextual con clic derecho) abre un buscador interactivo sobre todas las líneas del fichero Excel de Alma Analytics cargado (si aún no se había seleccionado el Excel de Analytics, ofrece seleccionarlo en ese momento):
  - **Asignar solo PO Line** (acción por defecto y con doble clic): escribe únicamente la `PO Line` en la línea de factura sin alterar ni sobrescribir `Reporting Codes` ni `Funds`.
  - **Asignar PO Line + códigos y fondos de Analytics**: opción explícita para copiar también los códigos de informe y fondos del registro de Analytics cuando el usuario lo desee.

#### Edición de celdas, arrastre tipo Excel (`■`), borradores y atajos de teclado

- **Borde de selección y tirador de relleno tipo Excel (`■`)**: al seleccionar una celda o rango aparece el marco de celda activa con el cuadradito `■` en la esquina inferior derecha:
  - **Arrastrar el cuadradito `■`** hacia abajo o hacia arriba propaga el valor de la celda activa a las filas atravesadas (igual que en Excel).
  - **Doble clic en el cuadradito `■`** rellena automáticamente las celdas vacías hacia abajo en esa columna hasta encontrar el siguiente valor no vacío.
- **Edición directa al teclear**: con una celda editable seleccionada, basta con empezar a escribir directamente para editarla (como en Excel), además de doble clic o `Intro`.
- **Barra de acciones rápidas tipo Excel**: botones directos sobre la tabla para `↩ Deshacer`, `↪ Rehacer`, `⬇ Rellenar abajo (Ctrl+D)`, `⬇ Rellenar blancos (■)`, `Copiar`, `Pegar`, `Vaciar celda`, `Restablecer celda` y `Asignar valor a seleccionadas...`.
- **Guardar y cargar borradores (`.ftalma.json`)**: los botones **Guardar borrador...** y **Cargar borrador...** permiten guardar el progreso de edición de una factura o lote en un archivo ligero y reanudarlo en otra sesión sin perder ningún cambio.
- **Ciclo de vida limpio de celdas**: abrir otra celda, desplazarse o hacer clic fuera confirma y destruye automáticamente el editor anterior sin dejar cuadros superpuestos.
- `Intro` confirma el valor; `Tabulador` / `Mayús + Tabulador` confirma y avanza o retrocede a la siguiente celda editable.
- `Ctrl + Z` y `Ctrl + Y` permiten deshacer y rehacer cualquier cambio individual o relleno por arrastre.
- `Ctrl + D` rellena hacia abajo en la columna activa; `Ctrl + C` y `Ctrl + V` copian y pegan valores en las celdas seleccionadas; `Supr` o `Retroceso` vacía las celdas seleccionadas.
- `Ctrl + clic`, `Mayús + clic` y arrastre con el ratón permiten seleccionar múltiples líneas.
- Introducir una `PO Line` a mano asigna únicamente el código de línea de pedido (y evita falsos avisos de `SIN COINCIDENCIA`), sin sobrescribir automáticamente los `Reporting Codes` ni los `Funds`.
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

Si se pulsa **Convertir a formato Alma** en la pestaña **Convertir** habiendo cambios pendientes en **Editar**, la aplicación ofrece aplicarlos y guardarlos directamente en un solo paso.

El XML o XSIG original nunca se modifica.

El Excel final incorpora una hoja de auditoría denominada **Cambios FacturaeToAlma**.

### Personalizar

Permite configurar:

- nombre de la institución;
- directorios habituales;
- logo institucional opcional;
- **Proveedores recordados (`NIF / CIF → Vendor Account`)**: tabla visual para consultar, añadir, modificar o eliminar las asociaciones de proveedor guardadas en `config.json`, con botones **Importar Excel...** (modos *Incremental/Acumulativo*, *Solo añadir nuevos* o *Carga completa*) y **Exportar / Plantilla...** (cabeceras estándar `NIF / CIF` y `Vendor Account`, aceptando alias como `NIF`, `CIF`, `Tax ID`, `Vendor Code` o `Cuenta proveedor`).

### Ayuda

Incluye instrucciones de uso, licencia, limitaciones conocidas y enlaces del proyecto.

## Funcionamiento de la conversión

1. Seleccionar factura o lote.
2. Seleccionar Alma Analytics.
3. Seleccionar la plantilla de Alma.
4. Definir el Excel de salida.
5. Revisar o corregir directamente en **Editar** (o ejecutar la conversión inicial desde **Convertir**).
6. Revisar la factura original en **Visualizar**.
7. Pulsar **Guardar cambios y actualizar Excel** (o **Convertir a formato Alma**).
8. Revisar el Excel final y, si procede, el informe de validación.

## Emparejamiento de PO Lines

La prioridad general es:

1. Asignación manual en la sesión (`Editar`);
2. `PO Line` explícita presente en el XML de la factura (`POL-...`);
3. ISBN (incluyendo conversión automática ISBN-10 ↔ ISBN-13);
4. ISSN;
5. título (4 niveles).

Para títulos se aplican:

- coincidencia exacta normalizada (con expansión de abreviaturas editoriales frecuentes);
- inclusión segura de palabras completas;
- similitud textual ponderada;
- rescate diagnóstico por rareza de palabras (`IDF`) y título principal MARC 245 (`$a`, previo a `:` o `/`) con protección de números de volumen y edición.

Las coincidencias ambiguas no se asignan automáticamente. Las líneas identificadas como gastos de envío (`SHIPMENT`) o descuentos (`DISCOUNT`) se clasifican automáticamente con su `Line type` correspondiente sin generar falsos avisos de falta de `PO Line`.

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
- `REVISAR CANTIDAD`;
- `REVISAR CUADRE IMPORTE`;
- `ASIGNACION MANUAL`.

`Quantity for Pricing` se utiliza como control. No modifica automáticamente la cantidad de factura ni la PO Line.

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
python facturae_to_alma_5_0.py
```

Comprobación de sintaxis, calidad y seguridad:

```bash
python -m py_compile facturae_to_alma_5_0.py
ruff check --select F,E9,B facturae_to_alma_5_0.py
python -m bandit -r facturae_to_alma_5_0.py
```

## Empaquetado para Windows

```bat
python -m PyInstaller --clean --onefile --windowed --noupx ^
  --name "FacturaeToAlma_5_0" ^
  --collect-submodules=openpyxl ^
  --collect-data=openpyxl ^
  --collect-all=defusedxml ^
  --collect-all=PIL ^
  facturae_to_alma_5_0.py
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
- Auditoría limpia con Bandit (`0 issues`) y Ruff (`0 issues`).

## Limitaciones conocidas de Alma

La plantilla Excel actual de Alma no permite:

- indicar `Line Exclusive` o «Línea exclusiva»;
- indicar expresamente si una línea queda parcial o completamente facturada.

## Licencia

GNU General Public License version 3.0 only (`GPL-3.0-only`).

## Historial

Consulta [CHANGELOG_5.0.md](CHANGELOG_5.0.md).
