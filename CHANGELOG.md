# Changelog

Todas las modificaciones relevantes de **FacturaeToAlma** se documentan en este archivo.

---

## [3.2_dev] - En desarrollo

La versión `3.2_dev` mantiene las medidas de seguridad y las funcionalidades de `3.1`, y mejora el emparejamiento por título para bases de datos, plataformas, paquetes electrónicos y descripciones comerciales extensas.

### Añadido

- Coincidencia exacta de títulos después de normalizar mayúsculas, minúsculas, tildes, puntuación y stopwords.
- Coincidencia por **inclusión segura de secuencias completas de palabras** antes de aplicar el algoritmo histórico de similitud.
- Reconocimiento de títulos breves y distintivos incluidos dentro de descripciones comerciales más largas.
- Identificación interna del método utilizado para comparar títulos:
  - `EXACT`;
  - `CONTAINMENT`;
  - `SIMILARITY`.
- Inclusión del método de comparación en el mensaje de validación cuando existen varias PO Lines candidatas por título.
- Lista de términos genéricos que no pueden provocar por sí solos una coincidencia automática por inclusión.
- Pruebas específicas para títulos de bases de datos y para evitar falsos positivos evidentes.

### Orden de comparación por título

Cuando no se ha localizado previamente una PO Line por ISBN o ISSN, la aplicación utiliza este orden:

1. Coincidencia exacta del título normalizado.
2. Inclusión segura de una secuencia completa de palabras.
3. Algoritmo histórico de similitud.
4. Ausencia de coincidencia o coincidencia ambigua.

La prioridad general continúa siendo:

1. ISBN.
2. ISSN.
3. Título.

Por tanto, la nueva lógica de inclusión no interviene en los libros o revistas cuando ya existe una coincidencia válida por ISBN o ISSN.

### Casos resueltos

La inclusión segura permite relacionar, entre otros, ejemplos como:

```text
MATHSCINET - ONLINE PACKAGE
MathSciNet
```

Y también una descripción comercial extensa que contenga el nombre distintivo:

```text
Suscripción AENORmás 2026 Precios especiales...
AENORMás
```

### Salvaguardas frente a falsos positivos

- La comparación se realiza con palabras completas, no con fragmentos de caracteres.
- Un término como `art` no coincide con `artificial`.
- Un título de una sola palabra debe tener al menos cinco caracteres.
- Un título de una sola palabra no puede pertenecer a la lista de términos genéricos.
- Entre los términos comerciales genéricos excluidos se encuentran:
  - `online`;
  - `package`;
  - `database`;
  - `subscription`;
  - `platform`;
  - `service`;
  - y sus equivalentes en español incluidos en el código.
- Si la inclusión devuelve varias PO Lines diferentes, no se asigna ninguna automáticamente y se genera `TITULO AMBIGUO`.
- El algoritmo tradicional de similitud se conserva como último recurso.

### Pruebas realizadas

- Coincidencia sin distinguir mayúsculas y minúsculas.
- `MATHSCINET - ONLINE PACKAGE` frente a `MathSciNet`.
- Descripción comercial extensa frente a `AENORMás`.
- Rechazo de `art` dentro de `artificial`.
- Rechazo de una coincidencia basada únicamente en el término genérico `Online`.
- Detección de ambigüedad cuando dos PO Lines contienen el mismo título incluido.

### Conservado sin cambios funcionales

- Medidas de seguridad y correcciones de Bandit incorporadas en `3.1`.
- Conversión individual y por lotes.
- Gestión visual y editable de lotes.
- Detección de rutas y números de factura duplicados.
- Restricción de los lotes a un mismo proveedor y una misma biblioteca.
- Uso de un único fichero Excel de Alma Analytics por conversión.
- Preferencias persistentes, institución y logo personalizables.
- Conversión en segundo plano.
- Comparación entre `Quantity` y `Quantity for Pricing`.
- Aviso `REVISAR CANTIDAD`.
- Generación del Excel de Alma y del informe de validación.
- Procesamiento XML mediante `defusedxml`.
- Sanitización de textos antes de escribirlos en Excel.

### Pendiente de evaluación

- Revisión de coincidencias por inclusión durante pruebas con facturas reales.
- Especial atención a libros sin ISBN o ISSN y títulos de una sola palabra.
- Posible incorporación futura del método de coincidencia como columna específica del informe de validación.

---

## [3.1] - Versión estable

- Corrección de avisos de bajo impacto detectados mediante Bandit.
- Sustitución de `subprocess.Popen()` por `subprocess.run()`.
- Uso explícito de `shell=False`.
- Resolución de `open` y `xdg-open` mediante `shutil.which()`.
- Normalización y validación reforzada de rutas.
- Corrección del bloque silencioso `except Exception: pass`.
- Captura específica de `tk.TclError` y registro mediante `logging.warning()`.
- Anotaciones `# nosec` específicas y justificadas para `B404`, `B606` y `B603`.

---

## [3.0] - Versión estable

- Interfaz organizada en las pestañas **Convertir**, **Personalizar** y **Ayuda**.
- Logo institucional opcional en PNG, JPG o JPEG.
- Redimensionado proporcional del logo hasta 100 píxeles de altura.
- Institución y directorios personalizables.
- Preferencias persistentes en `%APPDATA%`.
- Gestión visual y editable de facturas en modo lote.
- Conversión en segundo plano.
- Nombre de salida del lote con el patrón `Lote_[primera factura]_Alma.xlsx`.
- Licencia `GPL-3.0-only`.

---

## [2.5_dev] - Versión de desarrollo

- Introdujo la interfaz por pestañas.
- Añadió preferencias persistentes.
- Añadió la lista visual de facturas de lote.
- Incorporó detección de rutas y facturas duplicadas.
- Incorporó conversión en segundo plano.
- Añadió mensajes finales diferenciados según existieran incidencias.

---

## [2.0] - Versión estable

- Estabilización del modo individual y del modo lote.
- Lotes restringidos a un mismo proveedor y una misma biblioteca.
- Uso de un único fichero Excel de Alma Analytics por conversión.
- Extracción robusta de proveedores empresa y personas físicas.
- Procesamiento XML seguro mediante `defusedxml`.
- Protección frente a fórmulas de Excel.
- Lectura de `Quantity for Pricing`.
- Comparación con `Quantity` y generación de `REVISAR CANTIDAD`.

---

## [1.3_dev] - Versión de desarrollo

- Incorporó la lectura correcta de `Quantity for Pricing`.
- Añadió la comparación con `Quantity` de la factura.
- Añadió ambas cantidades al informe de validación.

---

## [1.2_dev] - Versión de desarrollo

- Sustituyó el parser XML estándar por `defusedxml`.
- Añadió sanitización de textos escritos en Excel.
- Reforzó la comprobación de rutas.

---

## [1.1_dev] - Versión de desarrollo

- Añadió el modo lote.
- Mejoró la identificación de proveedores.
- Añadió el número de factura al informe de validación.

---

## [1.0] - Versión estable inicial

- Conversión individual de XSIG/XML/TXT a Excel compatible con Alma.
- Uso dinámico de la plantilla de Alma.
- Localización de PO Lines por ISBN, ISSN o título.
- Incorporación de reporting codes, fondos y fechas de suscripción.
- Desplegable de `Line type`.
- Generación del informe de validación.

---

## Limitaciones conocidas de Alma

Soporte de Ex Libris confirmó que la plantilla Excel actual de Alma no permite:

- indicar `Line Exclusive` o «Línea exclusiva»;
- indicar explícitamente si una línea queda parcial o completamente facturada.

La comparación con `Quantity for Pricing` es una ayuda para la revisión y no un estado transmitido a Alma.

## Estado actual

```text
Versión estable: 3.1
Versión de desarrollo: 3.2_dev
Rama estable: stable
Rama de desarrollo: dev
```
