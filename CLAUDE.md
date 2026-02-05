# PptxGenJS Fork

A TypeScript library for generating PowerPoint (.pptx) files programmatically. This fork extends the original PptxGenJS with a compositional API, modern tooling, and cleaner architecture.

## Philosophy

- **Breaking changes are welcome** - Prioritize API ergonomics over backward compatibility
- **No mutation** - XML generation must not mutate input options (use pure functions)
- **Composition over configuration** - High-level components built from primitives
- **Verbose naming** - Function/variable names should be self-documenting

## Architecture Overview

```
User Code
    │
    ▼
Compositional API (src/components/, src/layout/, src/styles/)
    │
    ▼
Primitive API (slide.addText, addShape, addImage in src/slide.ts)
    │
    ▼
Object Definitions (src/gen-objects.ts)
    │
    ▼
XML Generation (src/gen-xml.ts, gen-xml-text.ts, gen-charts.ts, gen-tables.ts)
    │
    ▼
JSZip → .pptx file
```

## Key Files

```
src/
├── pptxgen.ts          # Main class, presentation-level API
├── slide.ts            # Slide class - addText, addShape, addCard, etc.
├── core-interfaces.ts  # All TypeScript types (LARGE)
├── core-enums.ts       # Constants, ShapeType, ChartType (LARGE)
│
├── gen-objects.ts      # Convert API options → internal slide objects
├── gen-xml.ts          # Generate slide/shape XML (uses xmlbuilder2)
├── gen-xml-text.ts     # Generate text body XML
├── gen-charts.ts       # Generate chart XML + Excel data
├── gen-tables.ts       # Generate table XML, auto-paging
├── gen-media.ts        # Handle images, encode base64
├── gen-utils.ts        # Utility functions (inch2Emu, valToPts, etc.)
│
├── xml-namespaces.ts   # OOXML namespace constants
├── xml-builder.ts      # Legacy fluent XML builder (being phased out)
│
├── styles/             # Shadow presets ('sm', 'md', 'lg')
├── layout/             # Grid, Stack layout calculations
└── components/         # High-level components (Card, etc.)

lab/                    # Experiment scripts
tests/                  # Unit tests (Node.js test runner)
```

## Quick Start

```typescript
const pptx = new PptxGenJS()
const slide = pptx.addSlide()

// Primitives
slide.addText('Hello', { x: 1, y: 1, w: 4, fontSize: 24, color: 'FF0000' })
slide.addShape(pptx.ShapeType.roundRect, { x: 1, y: 2, w: 3, h: 2, fill: { color: 'F0F0F0' } })

// Compositional API
slide.addCard({
  x: 1, y: 1, w: 4, h: 2,
  heading: 'Title',
  body: 'Description',
  shadow: 'sm',  // Preset: 'sm' | 'md' | 'lg' | 'none'
})

await pptx.writeFile({ fileName: 'out.pptx' })
```

## Units

- **Inches** - API (x, y, w, h)
- **EMU** (914400/inch) - Internal XML
- **Points** (72/inch) - Fonts, line widths
- Converters: `inch2Emu()`, `valToPts()`

## XML Generation

Uses **xmlbuilder2** for structured XML generation. Key functions in `gen-xml.ts`:
- `makeXmlPresentation`, `makeXmlSlide`, `makeXmlContTypes` - use xmlbuilder2
- `slideObjectToXml`, `makeXmlTransition` - still use template literals (complex conditionals)

Namespace constants in `src/xml-namespaces.ts` (e.g., `NS_A`, `NS_P`, `REL_TYPE_SLIDE`).

## Development

```bash
pnpm install
pnpm run build    # Build with tsup (~50ms)
pnpm run dev      # Watch mode
pnpm test         # Run tests
```

## Known Issues

1. **`tests/deprecation-warnings.test.ts`** - Fails due to removed `DEPRECATED_PROPERTY_MAP`. Needs removal.
2. **`gen-utils.ts:correctShadowOptions()`** - Still mutates input.
3. **Theme system** - Not yet implemented (plan in `API-REDESIGN-PLAN.md`).

## Output

```
dist/
├── pptxgen.es.js       # ESM
├── pptxgen.cjs.js      # CommonJS
├── pptxgen.bundle.js   # IIFE (browser)
└── pptxgen.d.ts        # TypeScript declarations
```
