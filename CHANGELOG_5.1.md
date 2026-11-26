# Changelog

## [5.1] - 2026-11-25

Versión 5.1 estable, 100 % determinista y sin conexión, creada a partir de `facturae_to_alma_5_0.py` para garantizar el cuadre exacto al pasar a «Línea exclusiva» en Alma (adoptando la lógica exacta de importes e IVA de la versión 4.0) e incorporar soporte multilingüe completo en las lenguas oficiales y cooficiales de España junto con importación/exportación Excel de proveedores en dos columnas.

*(Nota: el selector multilingüe completo y la importación/exportación Excel de proveedores en 2 columnas se han retroportado también a `facturae_to_alma_5_0.py`).*

### Añadido

- **Selector de lenguas oficiales y cooficiales de España (`es`, `ca`, `va`, `gl`, `eu`, `oc`) en Personalizar → Preferencias**:
  - Soporte completo para Castellano/Español (`es`), Català (`ca`), Valencià (`va`), Galego (`gl`), Euskara (`eu`) y Aranés/Occitan (`oc`), con persistencia en `config.json` y actualización instantánea en caliente sin reiniciar la aplicación.
  - Cobertura integral de toda la interfaz gráfica: pestañas principales, etiquetas de directorios habituales, botones y notas explicativas en **Personalizar**, desplegables de estado y operadores de filtro en **Editar** (manteniendo activos los filtros aplicados al cambiar de lengua), menús contextuales, barra inferior de estado, barra de diagnóstico en vivo y sumatorio en **Editar**, pestaña **Ayuda**, visor **Visualizar** (subpestañas, tarjeta de factura estilo PDF, roles DIR3 y totales) y todos los diálogos `messagebox` y `filedialog`.
- **Importación y exportación Excel de proveedores en Personalizar (exclusivamente 2 columnas)**:
  - Exportación transaccional a `.xlsx` únicamente con las 2 columnas **`NIF / CIF`** y **`Vendor Account`** (`export_vendor_map_to_excel`), con sanitización frente a fórmulas de Excel.
  - Importación desde `.xlsx` (`import_vendor_map_from_excel`) validando que el archivo contenga únicamente las 2 columnas **`NIF / CIF`** y **`Vendor Account`** y rechazando hojas con campos de línea de pedido (`PO Line`, `Reporting Code`, `Fund and percent`, etc.).
  - Selector de tres modos de importación:
    1. **Incremental / acumulativo**: añade nuevos proveedores y actualiza los existentes, conservando el resto.
    2. **Solo añadir nuevos**: incorpora únicamente los NIF/CIF que no existían previamente sin sobrescribir los actuales.
    3. **Carga completa**: reemplaza toda la tabla guardada por el contenido del archivo Excel (previa confirmación).

### Cambiado y corregido

- **Gestión exacta de importes e IVA de la versión 4.0 para cuadre garantizado en «Línea exclusiva» de Alma**:
  - En `FacturaeParser`, `InvoiceData`, `InvoiceLineData`, `InvoiceDataExtractor`, `AlmaExcelExporter` y `convert_facturae_to_alma_excel` se restablece la lógica exacta de `facturae_to_alma_4_0.py` para la lectura y escritura de importes (`InvoiceTotal`, `GrossAmount` con respaldo a `UnitPriceWithoutTax`) e IVA (`TaxesOutputs/Tax/TaxAmount/TotalAmount` y `TaxesOutputs/Tax/TaxRate`, manteniendo `VAT Percentage` a `None` en la fila `INV`).
  - Elimina la herencia sintética de `TaxRate` de cabecera a línea y el recálculo redondeado línea a línea introducidos en 5.0, de modo que al conmutar la factura a **«Línea exclusiva»** en Alma el total cuadra siempre exactamente igual que en la versión 4.0.
- Protección en el cierre/redibujado diferido del marco de selección tipo Excel (`_update_excel_selection_overlay`) cuando el widget `Treeview` ya ha sido destruido.

### Seguridad y calidad

- Procesamiento XML con `defusedxml`, sanitización frente a inyección de fórmulas en Excel, ejecución externa con `shell=False` y `shutil.which()`, escritura atómica en exportaciones Excel, auditoría limpia con **Bandit (`0 issues`)** y **Ruff (`0 issues` en reglas `F,E9,B`)**, y validación automatizada funcional y de interfaz gráfica (`Tkinter` / `Xvfb`) en las 6 lenguas.

## Versiones anteriores

### [5.0] - 2026-11-19

- Lectura rápida en streaming y caché inteligente en memoria del Excel de Alma Analytics (`get_cached_alma_lookup`).
- Barra de diagnóstico en vivo, sumatorio de importes visibles/seleccionados y guardado/carga de borradores de sesión (`.ftalma.json`).
- Gestor visual de proveedores recordados (`NIF / CIF → Vendor Account`), actualizado también con selector multilingüe (`es`, `ca`, `va`, `gl`, `eu`, `oc`) e importación/exportación Excel en 2 columnas.
- Edición estilo Excel con marco activo, tirador de relleno (`■`), `Deshacer`/`Rehacer`, búsqueda rápida (`Ctrl+F`), autofiltro multicriterio y buscador interactivo «Buscar PO Line en Alma...».
- Conversión bidireccional ISBN-10 ↔ ISBN-13, detección de `PO Line` en XML (`POL-...`) y emparejamiento bibliográfico por título en 4 niveles.

### [4.0] - 2026-10-19

- Integración del visor FACe (**Visualizar**) y del editor previo del Excel de Alma (**Editar**) con guardado transaccional y hoja de auditoría **Cambios FacturaeToAlma**.
