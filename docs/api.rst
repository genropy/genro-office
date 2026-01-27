API Reference
=============

WordApp
-------

.. autoclass:: genro_office.WordApp
   :members:
   :undoc-members:
   :show-inheritance:

WordBuilder
-----------

.. autoclass:: genro_office.builders.word_builder.WordBuilder
   :members:
   :undoc-members:
   :show-inheritance:

ExcelApp
--------

.. autoclass:: genro_office.ExcelApp
   :members:
   :undoc-members:
   :show-inheritance:

ExcelBuilder
------------

.. autoclass:: genro_office.builders.excel_builder.ExcelBuilder
   :members:
   :undoc-members:
   :show-inheritance:

Element Reference
-----------------

Word Elements
~~~~~~~~~~~~~

document
^^^^^^^^

Root element for Word documents.

**Attributes:**

- ``title`` (str): Document title (added as H0 heading)
- ``orientation`` (str): "portrait" or "landscape"
- ``margin_top`` (float): Top margin in cm
- ``margin_bottom`` (float): Bottom margin in cm
- ``margin_left`` (float): Left margin in cm
- ``margin_right`` (float): Right margin in cm

**Children:** heading, paragraph, table, image, pagebreak, itemlist, header, footer, run

heading
^^^^^^^

Heading element (H1-H9).

**Attributes:**

- ``content`` (str): Heading text
- ``level`` (int): Heading level 1-9
- ``bold`` (bool): Bold override
- ``italic`` (bool): Italic override
- ``color`` (str): Text color (hex)

**Children:** run

paragraph
^^^^^^^^^

Text paragraph with formatting.

**Attributes:**

- ``content`` (str): Paragraph text
- ``style`` (str): Word style name
- ``bold`` (bool): Bold text
- ``italic`` (bool): Italic text
- ``underline`` (bool): Underlined text
- ``font_size`` (int): Font size in points
- ``font_name`` (str): Font family name
- ``color`` (str): Text color (hex)
- ``align`` (str): "left", "center", "right", "justify"
- ``space_before`` (float): Space before in points
- ``space_after`` (float): Space after in points
- ``line_spacing`` (float): Line spacing multiplier

**Children:** run

run
^^^

Inline text with formatting.

**Attributes:**

- ``content`` (str): Text content
- ``bold`` (bool): Bold text
- ``italic`` (bool): Italic text
- ``underline`` (bool): Underlined text
- ``strike`` (bool): Strikethrough
- ``font_size`` (int): Font size in points
- ``font_name`` (str): Font family name
- ``color`` (str): Text color (hex)
- ``highlight`` (str): Highlight color name

**Children:** none

itemlist
^^^^^^^^

List container (bulleted or numbered).

**Attributes:**

- ``type`` (str): "bullet" or "number"

**Children:** item

item
^^^^

List item.

**Attributes:**

- ``content`` (str): Item text

**Children:** paragraph, run

table
^^^^^

Table element.

**Attributes:**

- ``style`` (str): Word table style name
- ``align`` (str): "left", "center", "right"
- ``autofit`` (bool): Auto-fit table to content

**Children:** row

row (Word)
^^^^^^^^^^

Table row.

**Attributes:**

- ``height`` (float): Row height in cm

**Children:** cell

cell (Word)
^^^^^^^^^^^

Table cell.

**Attributes:**

- ``content`` (str): Cell text
- ``width`` (float): Column width in cm
- ``bold`` (bool): Bold text
- ``bg_color`` (str): Background color (hex)
- ``align`` (str): "left", "center", "right"
- ``valign`` (str): "top", "center", "bottom"

**Children:** paragraph, run

image
^^^^^

Image element.

**Attributes:**

- ``path`` (str): Image file path
- ``width`` (float): Width in inches
- ``height`` (float): Height in inches
- ``align`` (str): "left", "center", "right"

**Children:** none

pagebreak
^^^^^^^^^

Page break element.

**Attributes:** none

**Children:** none

header
^^^^^^

Document header.

**Attributes:** none

**Children:** paragraph, run

footer
^^^^^^

Document footer.

**Attributes:** none

**Children:** paragraph, run

Excel Elements
~~~~~~~~~~~~~~

workbook
^^^^^^^^

Root element for Excel workbooks.

**Attributes:** none

**Children:** sheet

sheet
^^^^^

Worksheet element.

**Attributes:**

- ``name`` (str): Sheet name
- ``freeze_panes`` (str): Cell reference to freeze at (e.g., "A2")
- ``autofilter`` (str): Range for autofilter (e.g., "A1:D10")

**Children:** row, merge, chart

row (Excel)
^^^^^^^^^^^

Spreadsheet row.

**Attributes:**

- ``height`` (float): Row height in points
- ``hidden`` (bool): Whether row is hidden

**Children:** cell

cell (Excel)
^^^^^^^^^^^^

Spreadsheet cell.

**Attributes:**

- ``content`` (Any): Cell value
- ``formula`` (str): Excel formula
- ``width`` (float): Column width
- ``bold`` (bool): Bold font
- ``italic`` (bool): Italic font
- ``underline`` (bool): Underlined font
- ``font_size`` (int): Font size in points
- ``font_color`` (str): Font color (hex)
- ``bg_color`` (str): Background color (hex)
- ``align`` (str): "left", "center", "right"
- ``valign`` (str): "top", "center", "bottom"
- ``wrap_text`` (bool): Wrap text in cell
- ``border`` (str): "thin", "medium", "thick", "double"
- ``border_color`` (str): Border color (hex)
- ``number_format`` (str): Excel number format

**Children:** none

merge
^^^^^

Merged cell range.

**Attributes:**

- ``range`` (str): Cell range (e.g., "A1:D1")

**Children:** none

chart
^^^^^

Chart element.

**Attributes:**

- ``type`` (str): "bar", "line", "pie"
- ``title`` (str): Chart title
- ``data_range`` (str): Data range (e.g., "B2:B10")
- ``categories_range`` (str): Categories range (e.g., "A2:A10")
- ``position`` (str): Cell position (e.g., "E1")
- ``width`` (float): Width in cm
- ``height`` (float): Height in cm

**Children:** none
