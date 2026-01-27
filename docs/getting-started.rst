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

genro-office uses the **App pattern** from genro-print. You create a class that inherits
from ``WordApp`` or ``ExcelApp`` and override the ``recipe`` method to define your document
structure.

.. code-block:: python

   from genro_office import WordApp


   class MyDocument(WordApp):
       def recipe(self, root):
           # Define document structure here
           doc = root.document(title="My Document")
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

- ``recipe(root)``: Override to define document structure
- ``render()``: Returns document as bytes
- ``save(filepath)``: Saves document to file
- ``page``: Access to the underlying Bag structure
- ``data``: A separate Bag for data binding (future use)

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
       def recipe(self, root):
           doc = root.document(title="Hello World")

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
       def recipe(self, root):
           wb = root.workbook()

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
