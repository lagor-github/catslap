![catslap](https://raw.githubusercontent.com/lagor-github/catslap/main/logo/catslap256.png)
# catslap
Última versión: `1.4.6`

`catslap` es una librería Python para la generación automática de documentos a partir de datos estructurados en JSON y plantillas parametrizables. Permite producir documentos finales en múltiples formatos evaluando directrices incrustadas directamente en las plantillas.

## Características principales
* Generación de documentos a partir de un archivo JSON de entrada.
* Soporte para múltiples formatos de salida:
  * Word
  * PowerPoint
  * Excel
  * Texto plano
  * HTML
  * JavaScript
* Plantillas simples o múltiples:
  * Un único archivo.
  * Varios archivos empaquetados en un ZIP o contenidos en un directorio.
* Posibilidad de limitar las extensiones de archivos de plantilla a procesar.
* Evaluación de expresiones y lógica con semántica Python.
* Renderización de HTML embebido en los datos JSON para formatos enriquecidos (Word y PowerPoint).

## Concepto general
El flujo de trabajo de `catslap` es el siguiente:
1. Se proporciona un archivo JSON con los datos de entrada.
2. Se define una o varias plantillas que contienen directrices.
3. `catslap` evalúa las directrices, accede a los datos y genera los documentos resolviendo las directrices de las plantillas obteniendo los documentos finales en los formatos deseados.

## Ejemplo de uso
Ejemplo de uso directo de la clase `Catslap`:
```python
import json
from catslap.catslap import Catslap

with open("data.json", "r", encoding="utf-8") as fh:
  json_map = json.load(fh)

catslap = Catslap(json_map)
catslap.process_dir_or_file(
  template="templates/example.docx",
  output="out/example.docx",
  exts=None,
  verbose=True,
)
```

## Uso por línea de comandos (prueba rápida)
También puedes ejecutar `catslap` desde la línea de comandos para probar rápidamente plantillas. Hay ejemplos dentro del directorio `test/`.

Ejemplos usando los datos de prueba existentes:
```bash
python catslap/catslap.py test/data/Catslap_sales.json test/templates/docx/Catslap_example.docx test/out
python catslap/catslap.py test/data/Catslap_sales.json test/templates test/out -v
```

## Acceso a datos desde la plantilla
El acceso a los datos del JSON se realiza mediante expresiones delimitadas por `{{ ... }}`.
La evaluación sigue el comportamiento de Python como si el JSON fuera un `dict`, con el añadido de permitir acceso mediante el operador de punto.

### Ejemplo de JSON de entrada
```json
{
  "report_name": "My report",
  "report_data": {
    "name": "BBS Tennesy",
    "account": "0000123",
    "values": [43, 56, 991, 2]
  }
}
```

### Ejemplos de acceso a datos
```text
{{report_data.account}}
{{report_data.get('account')}}
{{report_data['account']}}
```
Hay que tener especial cuidado con usar nombres de JSON que correspondan a tokens de Python para evitar problemas de evaluación. Por ejemplo, si se usa `items` dentro de un JSON, no se podría acceder a ese elemento mediante el operador punto (por ejemplo, `data.items`), pero sí se podría acceder mediante `data['items']` o `data.get('items')`

Cuando un valor del JSON contiene código HTML (comienza por una etiqueta HTML), este será renderizado de forma enriquecida en el formato de salida, siempre que el tipo de formato de documento lo permita.

## Directrices de plantilla
Las directrices se definen usando bloques `{% ... %}` y cada directriz debe ocupar un párrafo completo dentro de la plantilla.

### Tipos de directrices soportadas
* Bucles
* Condiciones
* Asignación de variables
* Configuraciones (dependientes del formato de salida)

### Bucles
Permiten iterar sobre listas del JSON. 
La sintaxis es: 
```
{% for <name> in <list-expression> %}
...
{% endfor %}
```
Ejemplo:
```text
{% for value in report_data.values %}
  {{value}}
{% endfor %}
```

### Condiciones
Permiten la ejecución condicional de bloques de contenido. La condición se evalúa como una expresión Python.
```text
{% if report_data.account %}
  Cuenta válida
{% elif report_data.account == '0' %}
  Cuenta cero
{% else %}
  Cuenta no definida
{% endif %}
```
Se puede añadir cualquier número de ramas `{% elif <expresión> %}` tras el `{% if %}`. Se renderiza la primera rama cuya condición sea verdadera; el resto se omite. `{% else %}` es opcional y solo se renderiza si ninguna rama anterior fue verdadera.

### Asignación de variables
Asigna un valor a una variable con nombre para poder reutilizarlo en expresiones posteriores.
El valor es una expresión Python evaluada contra el contexto JSON actual.
```
{% set <name>=<expresión> %}
```
Ejemplo:
```text
{% for item in report_data.items %}
  {% set detail=products[item.id] %}
  {{item.id}}  {{detail.name}}  {{detail.price}}
{% endfor %}
```
La variable existe únicamente dentro del scope en que fue declarada:
- Declarada dentro de `{% for %}` → se elimina al final de cada iteración.
- Declarada dentro de `{% if %}` → se elimina al llegar a `{% endif %}`.
- Declarada a nivel global → disponible para el resto del documento.

## Configuraciones de estilo (Word y PowerPoint)
Para documentos Word y PowerPoint, `catslap` permite definir cómo se renderiza el HTML encontrado en los datos JSON mediante directrices de estilo.
El formato de la directriz de estilo es:
```
{% style <keyword>=<style_name> %}
```
`<keyword>` son estilos predefinidos en `catslap` correspondientes a estilos de HTML.
`<style_name>` es el nombre del estilo que se utilizará de entre los estilos definidos en el documento de plantilla de Word o PowerPoint.

### Ejemplo de configuración de estilos
```text
{% style heading=Título 1 %}
{% style table_cell=Celda normal %}
{% style table_header=Celda cabecera %}
{% style table_header_bgcolor=#FF0000 %}
{% style table_cell_bgcolor=white %}
{% style table_cell_bgcolor2=#E8E8E8 %}
{% style table_caption=Tabla título %}
{% style code=Code %}
{% style codeblock=Codeblock %}
{% style token=Token %}
{% style link_title=LinkTitle %}
{% style link_url=LinkUrl %}
{% style quote=Cita destacada %}
```

### Estilos soportados

* `heading`
  Define el estilo para títulos HTML (`<H1>` a `<H6>`). Si se define un único estilo, se generan automáticamente los estilos sucesivos prefijados con el número 2, 3, 4, 5 y 6. Por defecto, ya está definido con los estilos: "Título1", ..., "Título6"

* `paragraph`
  Define el estilo para párrafos HTML `<P>`. Por defecto se usa el estilo "Normal"

* `list_bullet`
  Define el estilo para listas HTML `<UL>`. Si se define un único estilo, se generan automáticamente los estilos sucesivos prefijados con el número 2, 3, 4, 5 y 6 para las sucesivas identaciones de lista. Por defecto, ya está definido con los estilos: "Lista con viñetas1", ..., "Lista con viñetas6" 

* `list_number`
  Define el estilo para listas HTML `<OL>`. Si se define un único estilo, se generan automáticamente los estilos sucesivos prefijados con el número 2, 3, 4, 5 y 6 para las sucesivas identaciones de lista. Por defecto, ya está definido con los estilos: "Lista con números1", ..., "Lista con números6"

* `table_cell`
  Estilo de los párrafos dentro de `<TD>`. Se preserva el formato de caracteres definido por el estilo Word referenciado al renderizar el contenido de la celda.

* `table_header`
  Estilo de los párrafos dentro de `<TH>`. Se preserva el formato de caracteres definido por el estilo Word referenciado al renderizar el contenido de la cabecera.

* `table_header_bgcolor`
  Color de fondo por defecto de las cabeceras de tabla.

* `table_cell_bgcolor`
  Color de fondo por defecto de las celdas de tabla.

* `table_cell_bgcolor2`
  Color de fondo alternativo para filas impares (opcional).

* `table_caption`
  Estilo del párrafo para `<CAPTION>`.  

* `code`
  Estilo de carácter para contenido dentro de `<code>`.

* `codeblock`
  Estilo de párrafo para bloques `<pre>`.

* `token`
  Estilo de párrafo para `<div class="token">`.

* `link_title`
  Estilo de párrafo para el texto de los enlaces.

* `link_url`
  Estilo de párrafo para la URL de los enlaces.

* `quote`
  Estilo de párrafo para bloques de cita destacados.

## Renderización de HTML (Word y PowerPoint)

`catslap` soporta la interpretación de un subconjunto de HTML para generar documentos enriquecidos.

### Etiquetas soportadas

* `<P>`: Párrafos, con soporte de CSS:

  * `text-align`
  * `color`
  * `font-weight`
  * `font-style`
  * `text-decoration`

* `<H1>` a `<H6>`: Títulos de capítulo.

* `<OL>`, `<UL>`, `<LI>`: Listas ordenadas y desordenadas.

* `<PRE>`: Bloques de código.

* `<BLOCKQUOTE>`: Citas destacadas.

* `<CODE>`: Código en línea.

* `<EM>`, `<I>`: Itálica.

* `<STRONG>`, `<B>`: Negrita.

* `<U>`: Subrayado.

* `<STROKE>`: Texto tachado.

* `<FONT color="">`: Color de texto (también mediante CSS `color`).

* `<TABLE>`, `<TR>`, `<TD>`, `<TH>`, `<CAPTION>`, `<THEAD>`, `<TBODY>`: Definición de tablas.

* `<IMG>`: Imágenes.

* `<A href="">...</A>`. Enlaces.

* `<DIV class="<style>">`: Aplicación de estilos de bloque predefinidos (`token`, `table_cell`, `codeblock`, etc.).

* `<SPAN class="<style>">`: Aplicación de estilos a nivel de caracteres (solo `code`).

## Licencia

Licencia MIT

## ChangeLog

**1.4.6**
- Mejorada la actualización de tablas de contenido y exportación mediante LibreOffice, con limpieza más segura de archivos temporales y comportamiento de error más claro.
- Corregidos detalles de renderizado DOCX en runs generados y fragmentos XML, evitando espaciado no deseado en código y manteniendo salida XML compacta cuando es necesario.

**1.4.4**
- Solucionados muchos problemas de renderizado en DOCX, mejorando la estabilidad y fidelidad al generar documentos Word desde plantillas.

**1.4.1**
- Corregida la generación de tablas HTML que contienen filas vacías para que no rompan la estructura final de la tabla en Word.
- Se preservan los estilos de carácter definidos en los estilos Word de celda y cabecera al renderizar contenido dentro de celdas de tabla.

**1.4.0**
- Nueva directiva `{% elif <expresión> %}` para ramas condicionales. Se pueden añadir tantas ramas `elif` como se necesite tras un `{% if %}`; solo se renderiza la primera rama verdadera.
- Soporte completo de directivas en plantillas PowerPoint: `{% for %}`, `{% if %}`, `{% elif %}`, `{% else %}`, `{% set %}`, `{% colormap %}` funcionan ahora en archivos `.pptx` con la misma semántica que en plantillas Word.

**1.3.0**
- Nueva directiva `{% set name=value %}` para asignación de variables con scope. Las variables declaradas dentro de un bloque `for` o `if` se eliminan automáticamente al salir del bloque; las declaraciones globales persisten para todo el documento.
- Funciones nativas de Python disponibles en expresiones: `len`, `str`, `int`, `float`, `bool`, `list`, `dict`, `sum`, `min`, `max`, `abs`, `round`, `sorted`, `any`, `all`, `isinstance`, etc.
- Las variables de bucle declaradas con `for` y `set` son ahora accesibles junto a los datos del JSON en la misma expresión (p.ej. `dependencies[name]`).
- Las imágenes HTML con esquema de URL `blob:` (`blob:https://host/uuid`) se descargan del servidor y se incrustan directamente en el documento Word.
- Mejora en el procesamiento del tamaño de imágenes HTML: las dimensiones especificadas como valores CSS (p.ej. `width: 300px`) se interpretan correctamente; las imágenes se escalan automáticamente al ancho máximo de página cuando lo superan.

**1.2.1**
- Corregido error al establecer estilos numerados (Título1, Título2, ...)

**1.2.0**
- Soporta todos las formas de colores CSS
- Soporta estilos de caracteres en tablas HTML y anchos proporcionales

**1.1.1**
- Corregido un problema en las tablas HTML

**1.1.0**
- Nueva característica: soporta resaltados de texto
- Nueva directiva: 'colormap' para mapear un color de HTML a otro en tiempo de generación de documento
- Corregidos bugs de estilo

**1.0.8**
- Resubido por cambios imcompletos

**1.0.7**
- Párrafo por defecto, justificado
- Corregido problema de mayúsculas en tags HTML
- Se respetan las líneas en blanco en los <PRE>

**1.0.6**
- Word style bug fixed

**1.0.5**
- Corregidos algunos errores

**1.0.4**
- Nuevo método para `class Catslap`:
  `process(self, template_file: str) -> bytes`
  Procesa la plantilla y obtiene los bytes del documento de salida

**1.0.3**
- Release inicial
