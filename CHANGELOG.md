# Changelog

## [4.0] - 2026-10-19

Versión estable que integra conversión, visualización y edición del Excel de Alma en una única aplicación.

### Añadido

- Visor FACe integrado en el archivo principal.
- Pestaña **Visualizar** con vista de factura y XML.
- Rueda del ratón en las vistas de factura, XML y edición.
- Pestaña **Editar** entre **Visualizar** y **Personalizar**.
- Presentación de la cabecera `INV` y las líneas `IL` del Excel de Alma.
- Edición de `Line type`, `PO Line`, identificadores, título, reporting codes, fechas de suscripción, notas y fondos.
- Selector de `Line type` con los valores admitidos por Alma.
- Sugerencias contextuales para reporting codes y `Fund and percent`.
- Selección múltiple de líneas y aplicación masiva de valores.
- Acciones para rellenar celdas vacías seleccionadas o inferiores.
- Menú contextual en la aplicación.
- Opciones **Original**, **Prefijo** y **Sufijo** para el número de factura.
- Propuesta dinámica del año en curso entre corchetes.
- Confirmación de edición mediante `Intro` y `Tabulador`.
- Botón **Guardar cambios y actualizar Excel**.
- Guardado transaccional mediante un archivo temporal.
- Verificación de las celdas editadas antes de sobrescribir el Excel de salida.
- Hoja de auditoría **Cambios FacturaeToAlma**.

### Cambiado

- La interfaz queda organizada en:
  - **Convertir**;
  - **Visualizar**;
  - **Editar**;
  - **Personalizar**;
  - **Ayuda**.
- El visor deja de distribuirse como script independiente.
- La edición se aplica únicamente al Excel de Alma; el XML/XSIG original permanece intacto.
- La fecha de factura se presenta en formato `MM/DD/AAAA` y queda protegida.
- El número de factura solo puede conservarse o modificarse mediante prefijo o sufijo.
- Se protegen los importes, cantidades, importes de IVA y porcentajes de IVA.
- El guardado actualiza inmediatamente el Excel indicado en **Convertir**.
- Las operaciones masivas actualizan primero el modelo y después redibujan la tabla.
- Se elimina la confirmación duplicada causada por eventos simultáneos de teclado y pérdida de foco.

### Conservado

- Conversión individual y por lotes.
- Restricción de los lotes a un mismo proveedor y biblioteca.
- Plantilla dinámica de Alma.
- Emparejamiento por ISBN, ISSN y título.
- Coincidencia exacta normalizada, inclusión segura y similitud textual.
- Comparación de `Quantity` con `Quantity for Pricing`.
- Informe de validación `_validacion.xlsx`.
- Preferencias persistentes y logo institucional opcional.
- Licencia GPL-3.0-only.

### Seguridad y calidad

- XML procesado con `defusedxml`.
- Sanitización frente a fórmulas de Excel.
- Ejecución externa con `shell=False`.
- Resolución de ejecutables mediante `shutil.which()`.
- Validación de rutas y protección de los archivos de entrada.
- Conversión en segundo plano.
- Comprobación con `py_compile`.
- Auditoría con Bandit.
- Corrección del aviso de depreciación relativo a la evaluación booleana de elementos XML.

### Limitaciones conocidas

La plantilla Excel actual de Alma no permite:

- indicar `Line Exclusive` o «Línea exclusiva»;
- indicar expresamente si una línea queda parcial o completamente facturada.

## Versiones anteriores

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
