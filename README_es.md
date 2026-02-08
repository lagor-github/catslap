![catslap](https://raw.githubusercontent.com/lagor-github/catslap/main/logo/catslap256.png)
# catslap
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
{% else %}
  Cuenta no definida
{% endif %}
```

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

* `style_paragraph`
  Define el estilo para párrafos HTML `<P>`. Por defecto se usa el estilo "Normal"

* `style_list_bullet`
  Define el estilo para listas HTML `<UL>`. Si se define un único estilo, se generan automáticamente los estilos sucesivos prefijados con el número 2, 3, 4, 5 y 6 para las sucesivas identaciones de lista. Por defecto, ya está definido con los estilos: "Lista con viñetas1", ..., "Lista con viñetas6" 

* `style_list_number`
  Define el estilo para listas HTML `<OL>`. Si se define un único estilo, se generan automáticamente los estilos sucesivos prefijados con el número 2, 3, 4, 5 y 6 para las sucesivas identaciones de lista. Por defecto, ya está definido con los estilos: "Lista con números1", ..., "Lista con números6"

* `table_cell`
  Estilo de los párrafos dentro de `<TD>`.

* `table_header`
  Estilo de los párrafos dentro de `<TH>`.

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

Pendiente de definir.
