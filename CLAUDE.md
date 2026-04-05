# Claude Code Instructions - genro-office

**Parent Document**: This project follows all policies from the central [meta-genro-modules CLAUDE.md](https://github.com/softwellsrl/meta-genro-modules/blob/main/CLAUDE.md)

## Project-Specific Context

### Current Status

- Development Status: Alpha
- Has Implementation: Yes

### Project Description

Reactive office document generation for Genropy using genro-builders architecture.
Word and Excel builders with `^pointer` data binding — `main()` defines template, data provides content.

### Architecture

```text
WordApp / ExcelApp (BuilderManager)
    ├── builder_class = WordBuilder / ExcelBuilder   (@element schema)
    ├── compiler_class = WordCompiler / ExcelCompiler (Bag → bytes)
    ├── main(source)    ← template with ^pointers
    ├── data            ← content, styles, metadata (alias for reactive_store)
    ├── build()         ← store → main → build → render
    └── save(filepath)  ← write bytes
```

### Key Components

- **WordBuilder / ExcelBuilder**: Define document schema with `@element` decorators. No compile logic.
- **WordCompiler / ExcelCompiler**: Extend `BagCompilerBase`, produce bytes (docx/xlsx). Maintain live Document/Workbook for incremental updates. Support `register_handler()` for custom tag extensibility.
- **WordApp / ExcelApp**: Extend `BuilderManager`, provide `data` property (alias for `reactive_store`), `save()`.

### Reusable Components (`@component`)

Custom builder subclasses can define `@component` methods — parameterized blocks that expand
into `@element` structures at build time. Pattern:

1. Subclass `WordBuilder`/`ExcelBuilder`, redefine the root element with extended `sub_tags`
2. Define `@component` methods that populate the internal Bag with elements
3. Subclass `WordCompiler`/`ExcelCompiler` with component-transparent node walking
4. Link compiler to builder via `_compiler_class`

Component handlers MUST accept `**kwargs` (the resolver passes all node attributes).
Use `f'^{prefix}?attr'` for relocatable pointer paths.

See `examples/word/components/` and `examples/excel/components/` for reference.

### Dependencies

- `genro-builders>=0.10.0` (includes genro-bag transitively)
- Optional: `python-docx>=1.0.0` for Word support
- Optional: `openpyxl>=3.1.0` for Excel support

### Project-Specific Guidelines

- Base builder classes contain ONLY `@element` definitions — no compile logic
- Custom builder subclasses may add `@component` definitions for reusable blocks
- Compile logic lives in compiler classes (`compilers/` directory)
- `_compiler_class` is linked at module bottom of compiler files (after class definition)
- `float` types are strictly validated: use `20.0` not `20` for float parameters (height, width)
- Apps use `build()` before `save()` — `BuilderManager.build()` runs the full pipeline

---

**All general policies are inherited from the parent document.**
