Word Documents
==============

The ``WordApp`` class generates Microsoft Word documents (.docx) using python-docx.

Basic Structure
---------------

.. code-block:: python

   from genro_office import WordApp


   class MyDocument(WordApp):
       def main(self, source):
           doc = source.document(title="Document Title")
           # Add elements here

Document Settings
-----------------

The ``document`` element accepts page setup options:

.. code-block:: python

   doc = source.document(
       title="Report",              # Title (added as H0)
       orientation="landscape",     # portrait (default) or landscape
       margin_top=2.5,              # Top margin in cm
       margin_bottom=2.5,           # Bottom margin in cm
       margin_left=2.0,             # Left margin in cm
       margin_right=2.0,            # Right margin in cm
   )

Headings
--------

Headings support levels 1-9:

.. code-block:: python

   doc.heading(content="Main Title", level=1)
   doc.heading(content="Subtitle", level=2)
   doc.heading(content="Section", level=3)

   # With formatting
   doc.heading(
       content="Colored Heading",
       level=1,
       bold=True,
       italic=False,
       color="FF0000",  # Hex color
   )

Paragraphs
----------

Paragraphs support rich formatting:

.. code-block:: python

   # Simple paragraph
   doc.paragraph(content="Hello World!")

   # Formatted paragraph
   doc.paragraph(
       content="Important text",
       bold=True,
       italic=False,
       underline=False,
       font_size=14,
       font_name="Arial",
       color="0000FF",
       align="center",        # left, center, right, justify
       space_before=12,       # Points before paragraph
       space_after=6,         # Points after paragraph
       line_spacing=1.5,      # Line spacing multiplier
   )

Inline Runs
-----------

For mixed formatting within a paragraph, use runs:

.. code-block:: python

   para = doc.paragraph(content="This is ")
   para.run(content="bold", bold=True)
   para.run(content=" and ")
   para.run(content="italic", italic=True)
   para.run(content=" and ")
   para.run(content="red", color="FF0000")
   para.run(content=" text.")

Run attributes:

- ``content``: Text content
- ``bold``, ``italic``, ``underline``, ``strike``: Font styles
- ``font_size``: Size in points
- ``font_name``: Font family name
- ``color``: Hex color (e.g., "FF0000")
- ``highlight``: Highlight color ("yellow", "green", "cyan", "magenta", "blue", "red", "gray")

Lists
-----

Bulleted and numbered lists use ``itemlist`` and ``item``:

.. code-block:: python

   # Bullet list
   bullets = doc.itemlist(type="bullet")
   bullets.item(content="First item")
   bullets.item(content="Second item")
   bullets.item(content="Third item")

   # Numbered list
   numbers = doc.itemlist(type="number")
   numbers.item(content="Step one")
   numbers.item(content="Step two")
   numbers.item(content="Step three")

Tables
------

Tables are created with ``table``, ``row``, and ``cell``:

.. code-block:: python

   table = doc.table(
       style="Table Grid",    # Word table style
       align="center",        # Table alignment: left, center, right
   )

   # Header row
   header = table.row(height=1.5)  # Height in cm
   header.cell(content="Name", bold=True, bg_color="4472C4")
   header.cell(content="Value", bold=True, bg_color="4472C4")

   # Data rows
   row1 = table.row()
   row1.cell(content="Item A")
   row1.cell(content="100", align="right")

Cell attributes:

- ``content``: Cell text
- ``width``: Column width in cm
- ``bold``: Bold text
- ``bg_color``: Background color (hex)
- ``align``: Horizontal alignment (left, center, right)
- ``valign``: Vertical alignment (top, center, bottom)

Headers and Footers
-------------------

.. code-block:: python

   # Header
   header = doc.header()
   header.paragraph(content="Company Name", align="left")
   header.paragraph(content="Confidential", align="right", italic=True)

   # Footer
   footer = doc.footer()
   footer.paragraph(content="Page 1", align="center", font_size=10)

Images
------

.. code-block:: python

   doc.image(
       path="logo.png",
       width=2,           # Width in inches
       height=1,          # Height in inches (optional)
       align="center",    # left, center, right
   )

Page Breaks
-----------

.. code-block:: python

   doc.paragraph(content="End of page 1")
   doc.pagebreak()
   doc.paragraph(content="Start of page 2")

Complete Example
----------------

.. code-block:: python

   from genro_office import WordApp


   class QuarterlyReport(WordApp):
       def main(self, source):
           doc = source.document(
               title="Q4 2024 Report",
               margin_top=2.5,
               margin_bottom=2.5,
           )

           # Header
           header = doc.header()
           header.paragraph(content="Acme Corp", align="right", italic=True)

           # Content
           doc.heading(content="Executive Summary", level=1)
           doc.paragraph(
               content="Revenue grew 15% year-over-year.",
               bold=True,
           )

           doc.heading(content="Financial Results", level=2)

           table = doc.table(style="Table Grid")
           header_row = table.row()
           header_row.cell(content="Metric", bold=True, bg_color="4472C4")
           header_row.cell(content="Q4 2024", bold=True, bg_color="4472C4")
           header_row.cell(content="Q4 2023", bold=True, bg_color="4472C4")

           data = [
               ("Revenue", "$1.2M", "$1.04M"),
               ("Expenses", "$800K", "$750K"),
               ("Profit", "$400K", "$290K"),
           ]

           for metric, q4_24, q4_23 in data:
               row = table.row()
               row.cell(content=metric)
               row.cell(content=q4_24, align="right")
               row.cell(content=q4_23, align="right")

           # Footer
           footer = doc.footer()
           footer.paragraph(content="Confidential", align="center", font_size=9)


   if __name__ == "__main__":
       report = QuarterlyReport()
       report.save("q4_report.docx")
