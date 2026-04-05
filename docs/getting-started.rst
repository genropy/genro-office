Getting Started
===============

Installation
------------

Install genro-office with pip:

.. code-block:: bash

   # All features
   pip install genro-office[all]

   # Word support only
   pip install genro-office[word]

   # Excel support only
   pip install genro-office[excel]

The App Pattern
---------------

genro-office uses the **App pattern** from genro-builders. You create a class that inherits
from ``WordApp`` or ``ExcelApp`` and override the ``main`` method to define your document
structure.

.. code-block:: python

   from genro_office import WordApp


   class MyDocument(WordApp):
       def main(self, source):
           # Define document structure here
           doc = source.document(title="My Document")
           doc.paragraph(content="Hello!")


   # Create and save
   document = MyDocument()
   document.save("output.docx")

Key Concepts
------------

Builder
~~~~~~~

The builder (``WordBuilder`` or ``ExcelBuilder``) defines the available elements and their
attributes. It uses the ``@element`` decorator to define which elements can contain which
child elements.

App
~~~

The app (``WordApp`` or ``ExcelApp``) provides the high-level interface:

- ``main(source)``: Override to define document structure
- ``build()``: Run the full pipeline (store → main → build → render)
- ``render()``: Returns document as bytes
- ``save(filepath)``: Saves document to file
- ``data``: The shared reactive data store for data binding

Elements
~~~~~~~~

Elements are the building blocks of documents:

**Word elements:**

- ``document``: Root element with page settings
- ``heading``: Headings (H1-H9)
- ``paragraph``: Text paragraphs with formatting
- ``run``: Inline formatted text within paragraphs
- ``itemlist``: Bulleted or numbered lists
- ``item``: List items
- ``table``, ``row``, ``cell``: Tables
- ``image``: Images
- ``header``, ``footer``: Page headers and footers
- ``pagebreak``: Page breaks

**Excel elements:**

- ``workbook``: Root element
- ``sheet``: Worksheets
- ``row``: Table rows
- ``cell``: Table cells with formatting
- ``merge``: Merged cell ranges
- ``chart``: Charts (bar, line, pie)

Your First Word Document
------------------------

.. code-block:: python

   from genro_office import WordApp


   class HelloWorld(WordApp):
       def main(self, source):
           doc = source.document(title="Hello World")

           doc.heading(content="Welcome", level=1)
           doc.paragraph(content="This is my first document.")

           doc.heading(content="Features", level=2)

           # Bullet list
           items = doc.itemlist(type="bullet")
           items.item(content="Easy to use")
           items.item(content="Declarative syntax")
           items.item(content="Full formatting support")


   if __name__ == "__main__":
       doc = HelloWorld()
       doc.save("hello.docx")

Your First Excel Spreadsheet
----------------------------

.. code-block:: python

   from genro_office import ExcelApp


   class HelloExcel(ExcelApp):
       def main(self, source):
           wb = source.workbook()

           sheet = wb.sheet(name="Products")

           # Header row
           header = sheet.row()
           header.cell(content="Product", bold=True, width=20)
           header.cell(content="Price", bold=True, width=10)

           # Data rows
           products = [
               ("Widget A", 29.99),
               ("Widget B", 49.99),
               ("Widget C", 19.99),
           ]

           for name, price in products:
               row = sheet.row()
               row.cell(content=name)
               row.cell(content=price, number_format="$#,##0.00")

           # Total
           total = sheet.row()
           total.cell(content="Total", bold=True)
           total.cell(formula="=SUM(B2:B4)", bold=True, number_format="$#,##0.00")


   if __name__ == "__main__":
       sheet = HelloExcel()
       sheet.save("hello.xlsx")

Reusable Components
-------------------

For repeated document structures, use ``@component`` on a custom builder subclass.
Components are parameterized blocks that expand into elements at build time.

.. code-block:: python

   from genro_builders.builder import component, element

   from genro_office.builders.word_builder import WordBuilder


   class LetterBuilder(WordBuilder):
       # Extend document sub_tags to include the component
       @element(sub_tags="heading,paragraph,table,image,pagebreak,itemlist,header,footer,run,address_block")
       def document(self, title="", orientation=None, margin_top=None,
                    margin_bottom=None, margin_left=None, margin_right=None): ...

       @component(sub_tags="", parent_tags="document")
       def address_block(self, comp, prefix="sender", **kwargs):
           comp.paragraph(content=f"^{prefix}?name", bold=True)
           comp.paragraph(content=f"^{prefix}?street")
           comp.paragraph(content=f"^{prefix}?city")

The same component is reused with different ``prefix`` values — pointer paths are relocatable:

.. code-block:: python

   doc.address_block(prefix="sender")     # resolves ^sender?name, etc.
   doc.address_block(prefix="recipient")  # resolves ^recipient?name, etc.

Components also require a thin compiler subclass for transparency (walking into component
children). See ``examples/word/components/`` and ``examples/excel/components/`` for complete
working examples.
