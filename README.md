![catslap](https://raw.githubusercontent.com/lagor-github/catslap/main/logo/catslap256.png)
# catslap
`catslap` is a Python library for automatic document generation from structured JSON data and parameterized templates. It can produce final documents in multiple formats by evaluating directives embedded directly in the templates.

## Key features
* Document generation from an input JSON file.
* Support for multiple output formats:
  * Word
  * PowerPoint
  * Excel
  * Plain text
  * HTML
  * JavaScript
* Single or multiple templates:
  * A single file.
  * Multiple files packaged in a ZIP or contained in a directory.
* Ability to limit which template file extensions are processed.
* Expression and logic evaluation with Python semantics.
* Rendering of HTML embedded in JSON data for rich formats (Word and PowerPoint).

## General concept
The `catslap` workflow is:
1. Provide a JSON file with input data.
2. Define one or more templates containing directives.
3. `catslap` evaluates the directives, accesses the data, and generates the final documents in the desired formats.

## Accessing data from templates
JSON data is accessed through expressions delimited by `{{ ... }}`.
Evaluation follows Python behavior as if the JSON were a `dict`, with the addition of dot-operator access.

### Input JSON example
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

### Data access examples
```text
{{report_data.account}}
{{report_data.get('account')}}
{{report_data['account']}}
```
Be careful when using JSON names that match Python tokens to avoid evaluation issues. For example, if you use `items` inside a JSON object, you cannot access it with dot notation (e.g., `data.items`), but you can access it with `data['items']` or `data.get('items')`.

When a JSON value contains HTML code (starting with an HTML tag), it will be rendered in a rich format when the output document type supports it.

## Template directives
Directives are defined using `{% ... %}` blocks, and each directive must occupy a full paragraph in the template.

### Supported directive types
* Loops
* Conditions
* Configurations (format-dependent)

### Loops
Allow iterating over JSON lists.
Syntax:
```
{% for <name> in <list-expression> %}
...
{% endfor %}
```
Example:
```text
{% for value in report_data.values %}
  {{value}}
{% endfor %}
```

### Conditions
Allow conditional execution of content blocks. The condition is evaluated as a Python expression.
```text
{% if report_data.account %}
  Valid account
{% else %}
  Undefined account
{% endif %}
```

## Style configurations (Word and PowerPoint)
For Word and PowerPoint documents, `catslap` allows defining how HTML found in JSON data is rendered through style directives.
The style directive format is:
```
{% style <keyword>=<style_name> %}
```
`<keyword>` are predefined `catslap` styles corresponding to HTML styles.
`<style_name>` is the name of the style to use from the styles defined in the Word or PowerPoint template document.

### Style configuration example
```text
{% style heading=Heading 1 %}
{% style table_cell=Normal cell %}
{% style table_header=Header cell %}
{% style table_header_bgcolor=#FF0000 %}
{% style table_cell_bgcolor=white %}
{% style table_cell_bgcolor2=#E8E8E8 %}
{% style table_caption=Table caption %}
{% style code=Code %}
{% style codeblock=Codeblock %}
{% style token=Token %}
{% style link_title=LinkTitle %}
{% style link_url=LinkUrl %}
{% style quote=Highlighted quote %}
```

### Supported styles

* `heading`
  Defines the style for HTML headings (`<H1>` to `<H6>`). If a single style is defined, successive styles are automatically generated with prefixes 2, 3, 4, 5, and 6. By default, the styles are already defined as "Heading1", ..., "Heading6".

* `style_paragraph`
  Defines the style for HTML paragraphs `<P>`. By default, the "Normal" style is used.

* `style_list_bullet`
  Defines the style for HTML unordered lists `<UL>`. If a single style is defined, successive styles are automatically generated with prefixes 2, 3, 4, 5, and 6 for successive list indentations. By default, the styles are already defined as "Bullet List1", ..., "Bullet List6".

* `style_list_number`
  Defines the style for HTML ordered lists `<OL>`. If a single style is defined, successive styles are automatically generated with prefixes 2, 3, 4, 5, and 6 for successive list indentations. By default, the styles are already defined as "Numbered List1", ..., "Numbered List6".

* `table_cell`
  Paragraph style inside `<TD>`.

* `table_header`
  Paragraph style inside `<TH>`.

* `table_header_bgcolor`
  Default background color for table headers.

* `table_cell_bgcolor`
  Default background color for table cells.

* `table_cell_bgcolor2`
  Alternate background color for odd rows (optional).

* `table_caption`
  Paragraph style for `<CAPTION>`.

* `code`
  Character style for content inside `<code>`.

* `codeblock`
  Paragraph style for `<pre>` blocks.

* `token`
  Paragraph style for `<div class="token">`.

* `link_title`
  Paragraph style for link text.

* `link_url`
  Paragraph style for link URLs.

* `quote`
  Paragraph style for highlighted quote blocks.

## HTML rendering (Word and PowerPoint)

`catslap` supports interpretation of a subset of HTML to generate rich documents.

### Supported tags

* `<P>`: Paragraphs, with CSS support for:

  * `text-align`
  * `color`
  * `font-weight`
  * `font-style`
  * `text-decoration`

* `<H1>` to `<H6>`: Chapter headings.

* `<OL>`, `<UL>`, `<LI>`: Ordered and unordered lists.

* `<PRE>`: Code blocks.

* `<BLOCKQUOTE>`: Highlighted quotes.

* `<CODE>`: Inline code.

* `<EM>`, `<I>`: Italic.

* `<STRONG>`, `<B>`: Bold.

* `<U>`: Underline.

* `<STROKE>`: Strikethrough.

* `<FONT color="">`: Text color (also via CSS `color`).

* `<TABLE>`, `<TR>`, `<TD>`, `<TH>`, `<CAPTION>`, `<THEAD>`, `<TBODY>`: Table definition.

* `<IMG>`: Images.

* `<A href="">...</A>`: Links.

* `<DIV class="<style>">`: Apply predefined block styles (`token`, `table_cell`, `codeblock`, etc.).

* `<SPAN class="<style>">`: Apply character-level styles (only `code`).

## License

To be defined.
