# Claude Code Instructions - genro-office

**Parent Document**: This project follows all policies from the central [meta-genro-modules CLAUDE.md](https://github.com/softwellsrl/meta-genro-modules/blob/main/CLAUDE.md)

## Project-Specific Context

### Current Status

- Development Status: Alpha
- Has Implementation: Yes

### Project Description

Reactive office document generation for Genropy using genro-builders architecture.
Word and Excel builders with `^pointer` data binding — recipe defines template, data provides content.

### Architecture

```text
WordApp / ExcelApp (BagAppBase)
    ├── builder_class = WordBuilder / ExcelBuilder   (@element schema)
    ├── compiler_class = WordCompiler / ExcelCompiler (Bag → bytes)
    ├── recipe(store)   ← template with ^pointers
    ├── data            ← content, styles, metadata
    ├── setup()         ← preprocess → bind → compile
    └── save(filepath)  ← write bytes
```

### Key Components

- **WordBuilder / ExcelBuilder**: Define document schema with `@element` decorators. No compile logic.
- **WordCompiler / ExcelCompiler**: Extend `BagCompilerBase`, produce bytes (docx/xlsx). Maintain live Document/Workbook for incremental updates.
- **WordApp / ExcelApp**: Extend `BagAppBase`, override `compile()` for bytes output. Provide `render()` and `save()`.

### Dependencies

- `genro-builders>=0.2.0` (includes genro-bag transitively)
- Optional: `python-docx>=1.0.0` for Word support
- Optional: `openpyxl>=3.1.0` for Excel support

### Project-Specific Guidelines

- Builder classes contain ONLY `@element` definitions — no compile logic
- Compile logic lives in compiler classes (`compilers/` directory)
- `compiler_class` is linked at module bottom of compiler files (after class definition)
- `float` types are strictly validated: use `20.0` not `20` for float parameters (height, width)
- Apps require `setup()` before `render()` or `save()` — `BagAppBase.__init__` does NOT call recipe

---

**All general policies are inherited from the parent document.**
