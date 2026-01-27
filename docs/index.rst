genro-office Documentation
==========================

Office document generation for Genropy - Word and Excel builders using the App pattern.

.. toctree::
   :maxdepth: 2
   :caption: Contents:

   getting-started
   word
   excel
   api

Quick Example
-------------

Word Document
~~~~~~~~~~~~~

.. code-block:: python

   from genro_office import WordApp


   class MyReport(WordApp):
       def recipe(self, root):
           doc = root.document(title="Report")
           doc.heading(content="Introduction", level=1)
           doc.paragraph(content="Hello World!")


   report = MyReport()
   report.save("report.docx")

Excel Spreadsheet
~~~~~~~~~~~~~~~~~

.. code-block:: python

   from genro_office import ExcelApp


   class MySheet(ExcelApp):
       def recipe(self, root):
           wb = root.workbook()
           sheet = wb.sheet(name="Data")
           sheet.row().cell(content="Name").cell(content="Value")
           sheet.row().cell(content="Item").cell(content=100)


   sheet = MySheet()
   sheet.save("data.xlsx")

Installation
------------

.. code-block:: bash

   pip install genro-office[all]

Or install specific features:

.. code-block:: bash

   pip install genro-office[word]   # Word support only
   pip install genro-office[excel]  # Excel support only

Indices and tables
==================

* :ref:`genindex`
* :ref:`modindex`
* :ref:`search`
