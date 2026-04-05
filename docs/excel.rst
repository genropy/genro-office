Excel Spreadsheets
==================

The ``ExcelApp`` class generates Microsoft Excel spreadsheets (.xlsx) using openpyxl.

Basic Structure
---------------

.. code-block:: python

   from genro_office import ExcelApp


   class MySpreadsheet(ExcelApp):
       def main(self, source):
           wb = source.workbook()
           sheet = wb.sheet(name="Data")
           # Add rows and cells here

Worksheets
----------

Create worksheets with the ``sheet`` element:

.. code-block:: python

   wb = source.workbook()

   # Basic sheet
   sheet1 = wb.sheet(name="Data")

   # Sheet with freeze panes (keeps headers visible)
   sheet2 = wb.sheet(name="Sales", freeze_panes="A2")

   # Sheet with autofilter
   sheet3 = wb.sheet(name="Products", autofilter="A1:D100")

   # Combined
   sheet4 = wb.sheet(
       name="Report",
       freeze_panes="B3",
       autofilter="A1:E50",
   )

Rows
----

Create rows with optional height and visibility:

.. code-block:: python

   # Basic row
   row = sheet.row()

   # Row with height
   header = sheet.row(height=25)  # Height in points

   # Hidden row
   hidden = sheet.row(hidden=True)

Cells
-----

Cells support rich formatting:

.. code-block:: python

   row.cell(
       content="Value",           # Cell value
       formula="=SUM(A1:A10)",    # Formula (overrides content)
       width=15,                  # Column width (applies to whole column)

       # Font
       bold=True,
       italic=False,
       font_size=12,
       font_color="FF0000",       # Hex color

       # Fill
       bg_color="FFFF00",         # Background color

       # Alignment
       align="center",            # left, center, right
       valign="center",           # top, center, bottom
       wrap_text=True,            # Wrap long text

       # Border
       border="thin",             # thin, medium, thick, double
       border_color="000000",     # Border color

       # Number format
       number_format="#,##0.00",  # Excel number format
   )

Common Number Formats
~~~~~~~~~~~~~~~~~~~~~

- ``"#,##0"``: Number with thousands separator
- ``"#,##0.00"``: Number with 2 decimal places
- ``"$#,##0.00"``: Currency
- ``"0%"``: Percentage
- ``"yyyy-mm-dd"``: Date
- ``"hh:mm:ss"``: Time

Formulas
--------

Use Excel formulas in cells:

.. code-block:: python

   # Simple formula
   row.cell(formula="=A1+B1")

   # SUM
   row.cell(formula="=SUM(A1:A10)")

   # Cross-sheet reference
   row.cell(formula="=SUM(Sales!B2:B100)")

   # Complex formula
   row.cell(formula="=IF(A1>100, \"High\", \"Low\")")

Merged Cells
------------

Merge cells after creating the rows:

.. code-block:: python

   # Create header row
   header = sheet.row(height=30)
   header.cell(content="Report Title", bold=True, font_size=16)
   header.cell(content="")  # Empty cells for merge
   header.cell(content="")
   header.cell(content="")

   # Merge A1:D1
   sheet.merge(range="A1:D1")

   # Vertical merge
   sheet.merge(range="A2:A5")

Charts
------

Add charts to worksheets:

.. code-block:: python

   # Bar chart
   sheet.chart(
       type="bar",
       title="Sales by Product",
       data_range="B1:B10",         # Data including header
       categories_range="A2:A10",   # Category labels
       position="E2",               # Chart position
       width=15,                    # Width in cm
       height=10,                   # Height in cm
   )

   # Line chart
   sheet.chart(
       type="line",
       title="Monthly Trend",
       data_range="B1:C12",         # Multiple series
       categories_range="A2:A12",
       position="E15",
   )

   # Pie chart
   sheet.chart(
       type="pie",
       title="Market Share",
       data_range="B1:B5",
       categories_range="A2:A5",
       position="E2",
   )

Styling Examples
----------------

Header Row
~~~~~~~~~~

.. code-block:: python

   header = sheet.row(height=25)
   header.cell(
       content="Product",
       bold=True,
       bg_color="4472C4",
       font_color="FFFFFF",
       align="center",
       border="thin",
   )

Alternating Row Colors
~~~~~~~~~~~~~~~~~~~~~~

.. code-block:: python

   for i, (name, value) in enumerate(data):
       row = sheet.row()
       bg = "E8E8E8" if i % 2 == 0 else None
       row.cell(content=name, bg_color=bg)
       row.cell(content=value, bg_color=bg, align="right")

Currency Formatting
~~~~~~~~~~~~~~~~~~~

.. code-block:: python

   row.cell(
       content=1234.56,
       number_format="$#,##0.00",
       align="right",
   )

Complete Example
----------------

.. code-block:: python

   from genro_office import ExcelApp


   class SalesReport(ExcelApp):
       def main(self, source):
           wb = source.workbook()

           # Sales data sheet
           sales = wb.sheet(name="Sales", freeze_panes="A2", autofilter="A1:E9")

           # Title row
           title = sales.row(height=30)
           title.cell(
               content="Monthly Sales Report",
               bold=True,
               font_size=16,
               bg_color="1F4E79",
               font_color="FFFFFF",
           )
           for _ in range(4):
               title.cell(content="", bg_color="1F4E79")
           sales.merge(range="A1:E1")

           # Header row
           header = sales.row(height=20)
           headers = ["Month", "North", "South", "East", "West"]
           for h in headers:
               header.cell(
                   content=h,
                   bold=True,
                   bg_color="4472C4",
                   font_color="FFFFFF",
                   align="center",
                   border="thin",
                   width=12,
               )

           # Data
           data = [
               ("Jan", 12000, 9500, 15000, 11000),
               ("Feb", 13500, 10200, 16000, 12500),
               ("Mar", 14000, 10800, 17000, 13000),
               ("Apr", 13200, 11500, 16500, 12800),
               ("May", 15000, 12000, 18000, 14000),
               ("Jun", 16000, 12800, 19000, 15000),
           ]

           for month, n, s, e, w in data:
               row = sales.row()
               row.cell(content=month, border="thin")
               row.cell(content=n, number_format="#,##0", border="thin", align="right")
               row.cell(content=s, number_format="#,##0", border="thin", align="right")
               row.cell(content=e, number_format="#,##0", border="thin", align="right")
               row.cell(content=w, number_format="#,##0", border="thin", align="right")

           # Total row
           total = sales.row(height=22)
           total.cell(content="Total", bold=True, bg_color="FFC000", border="medium")
           for col in ["B", "C", "D", "E"]:
               total.cell(
                   formula=f"=SUM({col}3:{col}8)",
                   bold=True,
                   bg_color="FFC000",
                   border="medium",
                   number_format="#,##0",
               )

           # Chart
           sales.chart(
               type="bar",
               title="Sales by Region",
               data_range="B2:E8",
               categories_range="A3:A8",
               position="G3",
               width=14,
               height=10,
           )


   if __name__ == "__main__":
       report = SalesReport()
       report.save("sales_report.xlsx")
