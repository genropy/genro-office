# genro-office

Office document generation for Genropy - Word and Excel builders.

## Installation

```bash
pip install genro-office

# With Word support
pip install genro-office[word]

# With Excel support
pip install genro-office[excel]

# With all features
pip install genro-office[all]
```

## Quick Start

### Word Documents

```python
from genro_bag import Bag
from genro_office import WordBuilder

doc = Bag(builder=WordBuilder)
doc.document(title="My Report")
doc.heading(level=1, content="Introduction")
doc.paragraph(content="Hello World!")

WordBuilder.render(doc, "report.docx")
```

### Excel Spreadsheets

```python
from genro_bag import Bag
from genro_office import ExcelBuilder

doc = Bag(builder=ExcelBuilder)
sheet = doc.sheet(name="Data")
sheet.row().cell(content="Name").cell(content="Value")
sheet.row().cell(content="Alpha").cell(content=100)

ExcelBuilder.render(doc, "data.xlsx")
```

## Development Status

**Pre-Alpha** - This project is in early development.

## License

Apache License 2.0 - See [LICENSE](LICENSE) for details.
