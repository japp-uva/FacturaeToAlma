# Changelog

## [5.0] - 2026-11-19

Versión 5.0 estable, 100 % determinista y sin conexión (sin dependencias de inteligencia artificial), que conserva la interfaz clásica 4.0 e incorpora mejoras de extracción XML/IVA, rendimiento con caché de Alma Analytics, emparejamiento bibliográfico avanzado, buscador interactivo de catálogo, autofiltro multicriterio, diagnóstico en vivo, borradores guardables y edición estilo Excel.

### Añadido

- **Extracción avanzada de IVA Facturae/FACe**: lectura tolerante a espacios de nombres XML, herencia del porcentaje de IVA (`TaxRate`) de cabecera a línea cuando la línea lo omite, cálculo automático de cuota de IVA (`Base × TaxRate / 100`) y exportación compatible con plantillas de Alma donde `HIL` no incluye `VAT Percentage`.
- **Lectura rápida en streaming y caché inteligente del Excel de Alma Analytics**: carga mediante `openpyxl(..., read_only=True, data_only=True)` con `iter_rows(values_only=True)` y caché en memoria por `(ruta, tamaño, fecha_modificación)` (`get_cached_alma_lookup`) para evitar releer ficheros grandes al cambiar de pestaña, actualizar borradores o convertir.
- **Barra de diagnóstico en vivo y sumatorio en Editar**: muestra debajo de la tabla el motivo exacto del aviso o estado de cruce de la fila seleccionada (y avisos de cuadre contable de factura) junto con el sumatorio en tiempo real de Base e IVA de las filas visibles y seleccionadas.
- **Guardar y reanudar borradores de sesión (`.ftalma.json`)**: botones **Guardar borrador...** y **Cargar borrador...** en **Editar** para conservar todas las facturas cargadas, rutas, ediciones de cabecera/líneas, prefijos/sufijos y asignaciones manuales entre sesiones.
- **Gestor visual de proveedores recordados (`NIF / CIF → Vendor Account`) en Personalizar**: tabla interactiva para consultar, añadir, actualizar, eliminar o vaciar las asociaciones de proveedor persistidas en `config.json`, con normalización completa de CIF/NIF (`normalize_tax_id`) que conserva las letras iniciales.
- **Productividad Excel en Editar y Lote**:
  - **Marco de selección de 2 px y tirador de relleno (`■`)**: arrastrar `■` para propagar valores hacia arriba/abajo y doble clic en `■` para rellenar celdas vacías hacia abajo.
  - **Barra de acciones rápidas tipo Excel**: `Deshacer`, `Rehacer`, `Rellenar abajo (Ctrl+D)`, `Rellenar blancos (■)`, `Copiar`, `Pegar`, `Vaciar celda`, `Restablecer celda` y `Asignar valor a seleccionadas...`.
  - **Búsqueda rápida global (`Ctrl+F`)** en cualquier columna con limpieza rápida mediante `Esc` o `✕`.
  - **Edición directa al teclear** sobre la celda activa sin necesidad de hacer doble clic primero.
  - **Doble clic en «Facturas del lote»** para saltar directamente a esa factura en la pestaña **Editar**.
- **Filtro por estado y Autofiltro multicriterio estilo Excel en Editar**:
  - Selector por estado (`Todas las líneas`, `Solo líneas con problemas (sin PO Line o con avisos)`, `Solo líneas sin PO Line (pendientes)`, `Solo líneas correctas (con PO Line y sin avisos)`, `Solo líneas modificadas manualmente`) con contador en vivo y resaltado visual por estado.
  - Filtrado simultáneo por **una o varias columnas estilo Excel**: barra de condiciones rápidas (`Contiene / Igual`, `Igual exacto`, `Distinto de`, `Vacías`, `No vacías`) y ventana emergente de Autofiltro Excel con buscador y casillas de verificación al pulsar en cualquier cabecera de columna (`Columna ▼`) o en **Filtro Excel...**.
- **Buscador interactivo «Buscar PO Line en Alma...»**: ventana de búsqueda en tiempo real sobre todo el catálogo cargado del Excel de Alma Analytics con doble botón («Asignar solo PO Line» por defecto, o «Asignar PO Line + códigos y fondos de Analytics» bajo petición expresa), y asistente guiado para seleccionar el Excel de Analytics si aún no se había cargado.
- **Prefijo y Sufijo en tiempo real**: actualización inmediata del número de factura tanto en la tabla `Cabecera de factura (INV)` como en la columna `Factura` de todas las líneas en `Editar`, conservando el modo (`Original`, `Prefijo`, `Sufijo`) de cada factura del lote.
- **Carga inmediata en Editar**: al abrir la pestaña **Editar** con una factura seleccionada, las líneas se muestran automáticamente incluso antes de seleccionar la plantilla o el informe de Analytics (y se cruzan con Alma Analytics en cuanto dicho fichero está seleccionado).
- **Conversión directa con cambios en Editar**: si se pulsa **Convertir a formato Alma** con modificaciones pendientes en **Editar**, la aplicación ofrece aplicarlas y guardarlas directamente en un paso.
- **Conversión bidireccional ISBN-10 ↔ ISBN-13**: cálculo automático del dígito de control módulo 11 / módulo 10 para emparejar facturas con ISBN-10 frente a informes de Alma Analytics con ISBN-13 (y viceversa).
- **Detección de `PO Line` explícita en XML (`POL-...`)**: reconocimiento directo del identificador `POL-\d+` en referencias de línea del XML.
- **Detección automática de portes y descuentos**: clasificación automática de líneas de gastos de envío (`SHIPMENT`) y descuentos (`DISCOUNT`), evitando falsos avisos de `SIN COINCIDENCIA`.
- **Emparejamiento bibliográfico avanzado por título (4 niveles)**: expansión de más de 60 abreviaturas bibliográficas frecuentes, extracción del título principal MARC 245 (`$a`, antes de `:` o `/`) y ponderación por rareza de palabras (`IDF`) con protección estricta de números de volumen/edición.
- **Respeto estricto a las PO Lines manuales**: introducir o asignar manualmente una `PO Line` en **Editar** actualiza únicamente la `PO Line` y suprime el falso aviso de `SIN COINCIDENCIA`, **sin sobrescribir automáticamente los `Reporting Codes` ni los `Funds`**.
- **Control de cuadre contable (`REVISAR CUADRE IMPORTE`)**: verificación automática de que la suma de líneas + IVA − descuentos generales + cargos generales coincide con el total de la factura.
- **Selector de idioma en lenguas oficiales y cooficiales de España**: interfaz disponible en **Castellano**, **Català**, **Valencià**, **Galego**, **Euskara** y **Aranès (Occità)** desde **Personalizar → Preferencias** mediante traducciones embebidas 100 % locales (sin conexión ni dependencias externas), con cambio en vivo, traducción automática de diálogos emergentes y persistencia en `config.json` (`language`).
- **Importación y exportación de proveedores desde/hacia Excel (`.xlsx`)**: botones **Importar Excel...** y **Exportar / Plantilla...** en **Personalizar**, con 3 modos de carga (**Incremental / Acumulativo**, **Solo añadir nuevos** y **Carga completa**), cabeceras estándar `NIF / CIF` y `Vendor Account` con alias tolerantes, y compatibilidad con ficheros que incluyan columnas adicionales de la versión 5.1.

### Cambiado y corregido

- Se mantiene íntegramente la interfaz clásica de 5 pestañas (**Convertir**, **Visualizar**, **Editar**, **Personalizar**, **Ayuda**) con estilo sobrio `clam` (`#f0f0f0`) y funcionamiento 100 % local sin opciones ni dependencias de IA.
- **Ciclo de vida limpio de editores de celda**: corregida la superposición visual de widgets `Entry` / `Combobox` al cambiar de celda, desplazarse o abrir listas desplegables (`_close_all_cell_editors`).
- **Visibilidad permanente de «3. Archivo de salida» en modo Lote**: corregido el peso vertical de la rejilla en **Convertir** (`rowconfigure(2, weight=1)` para la lista del lote y `rowconfigure(3, weight=0)` para el archivo de salida) para evitar que el bloque de destino se oculte al activar el modo lote.
- Corregida la normalización de CIF/NIF (`normalize_tax_id`) para no recortar las letras iniciales de los CIF al recordar o consultar proveedores.

### Seguridad y calidad

- Procesamiento XML con `defusedxml`, sanitización frente a inyección de fórmulas en Excel, ejecución externa con `shell=False` y `shutil.which()`, guardado transaccional del XLSX, auditoría limpia con **Bandit (`0 issues`)** y **Ruff (`0 issues` en reglas `F,E9,B`)**, y verificación mediante pruebas GUI/funcionales automatizadas.

## Versiones anteriores

### [4.0] - 2026-10-19

- Visor FACe integrado en el archivo principal y pestaña **Visualizar** con vista de factura y XML.
- Rueda del ratón en las vistas de factura, XML y edición.
- Pestaña **Editar** entre **Visualizar** y **Personalizar** con presentación de la cabecera `INV` y las líneas `IL` del Excel de Alma.
- Edición de `Line type`, `PO Line`, identificadores, título, reporting codes, fechas de suscripción, notas y fondos.
- Sugerencias contextuales para reporting codes y `Fund and percent`.
- Selección múltiple de líneas, aplicación masiva de valores y menú contextual.
- Opciones **Original**, **Prefijo** y **Sufijo** para el número de factura con propuesta dinámica del año en curso entre corchetes.
- Confirmación de edición mediante `Intro` y `Tabulador`.
- Botón **Guardar cambios y actualizar Excel** con guardado transaccional y hoja de auditoría **Cambios FacturaeToAlma**.

### [3.2_dev]

- Mejora del emparejamiento por título.
- Inclusión segura de palabras completas.
- Salvaguardas frente a falsos positivos.
- Detección de títulos ambiguos.

### [3.1]

- Correcciones derivadas de Bandit.
- Uso de `subprocess.run()` y `shell=False`.
- Resolución de ejecutables mediante `shutil.which()`.

### [3.0]

- Interfaz inicial en pestañas.
- Logo institucional y preferencias persistentes.
- Conversión en segundo plano.

### [2.0]

- Estabilización de los modos individual y lote.
- XML seguro mediante `defusedxml`.
- Protección frente a fórmulas de Excel.
- Control con `Quantity for Pricing`.

### [1.0]

- Conversión inicial de Facturae a Excel de Alma.
- Plantilla dinámica.
- Emparejamiento de PO Lines.
- Reporting codes, fondos y fechas de suscripción.
- Informe de validación.
