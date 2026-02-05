# PptxGenJS API Redesign Plan

> **Goal**: Transform the low-level, imperative API into a component-based, compositional API that makes creating modern presentations dramatically easier.

## Current Pain Points

### 1. No Grouping/Composition Primitives
Currently, creating a "card" requires manually placing 3+ separate objects with absolute coordinates. Moving a card means recalculating all positions.

### 2. Pure Absolute Positioning
Every element requires manual x/y calculation. No layout system for common patterns like grids or stacks.

### 3. No Style Inheritance/Defaults
Font settings, colors, and spacing are repeated everywhere. No cascade or theming.

### 4. No Relative Positioning
Can't express "place this below the previous element" or "center this in the remaining space."

### 5. Verbose Configuration
Shadows require 6 properties. Common patterns have no presets.

---

## Proposed API Improvements

### Phase 1: Style System & Presets

**Style presets for common properties:**
```typescript
// Shadow presets
shadow: 'sm' | 'md' | 'lg' | 'none' | ShadowProps

// Typography presets (can be customized via theme)
const headingStyle = pptx.styles.heading1
const bodyStyle = pptx.styles.body

// Slide-level defaults
slide.defaults = { fontFace: 'Arial', fontSize: 14 }
```

**Implementation:**
- Add `StylePresets` module with shadow, typography presets
- Add `slide.defaults` that cascade to all children
- Support string shorthand that resolves to full config

---

### Phase 2: Layout Primitives

**Grid Layout:**
```typescript
slide.addGrid({
    x: 0.5, y: 1.0,
    cols: 2, rows: 2,
    gap: 0.3,           // or { x: 0.3, y: 0.2 }
    cellWidth: 4,       // optional, auto-calculated if omitted
    cellHeight: 1.5,
    children: [elem1, elem2, elem3, elem4]
})
```

**Stack Layout (vertical/horizontal):**
```typescript
slide.addStack({
    x: 0.5, y: 1.0,
    direction: 'vertical',  // or 'horizontal'
    gap: 0.2,
    children: [elem1, elem2, elem3]
})
```

**Flex Layout:**
```typescript
slide.addFlex({
    x: 0.5, y: 1.0,
    w: 9, h: 5,
    direction: 'row',
    wrap: true,
    gap: 0.3,
    justify: 'space-between',
    align: 'center',
    children: [...]
})
```

**Implementation:**
- Create `LayoutEngine` that calculates positions from declarative specs
- Each layout returns positioned children with computed x/y/w/h
- Layouts can be nested

---

### Phase 3: Composition & Groups

**Group primitive:**
```typescript
slide.addGroup({
    x: 1, y: 2,
    w: 4, h: 3,
}, (group) => {
    // Children positioned relative to group origin (0,0)
    group.addShape('rect', { x: 0, y: 0, w: '100%', h: '100%', fill: 'F5F5F5' })
    group.addText('Title', { x: 0.2, y: 0.2 })
    group.addText('Body', { x: 0.2, y: 0.6 })
})
```

**Or declarative children array:**
```typescript
slide.addGroup({
    x: 1, y: 2, w: 4, h: 3,
    children: [
        { type: 'shape', shape: 'rect', ... },
        { type: 'text', text: 'Title', ... },
    ]
})
```

**Implementation:**
- Groups translate child coordinates by group origin
- Support both callback and declarative children syntax
- Groups can have their own defaults that cascade

---

### Phase 4: High-Level Components

**Card Component:**
```typescript
slide.addCard({
    x: 0.5, y: 1.0,
    w: 4, h: 2,
    background: 'F5F5F5',
    borderRadius: 0.1,
    shadow: 'sm',
    padding: 0.2,        // or { top, right, bottom, left }
    children: [
        { type: 'heading', text: '1. LEARNING', color: 'C5A636' },
        { type: 'body', text: 'How machines acquire knowledge.' },
    ]
})
```

**List Component:**
```typescript
slide.addList({
    x: 0.5, y: 1.0,
    items: ['First', 'Second', 'Third'],
    bullet: 'number',    // or 'disc', 'check', custom
    gap: 0.1,
})
```

**Implementation:**
- Components are sugar over Groups + Layout + Styles
- Each component has sensible defaults
- Components can be extended/customized

---

### Phase 5: Declarative Slide Builder

**Ultimate API for the example slide:**
```typescript
pptx.addSlide({
    background: 'FFFFFF',
    children: [
        {
            type: 'text',
            text: 'Four Open Problems in AI Design',
            x: 0.5, y: 0.4,
            style: 'title',
            color: '2E9B7B',
        },
        {
            type: 'grid',
            x: 0.6, y: 1.3,
            cols: 2, rows: 2,
            gap: 0.3,
            children: [
                { type: 'card', heading: '1. LEARNING', ... },
                { type: 'card', heading: '2. REASONING', ... },
                { type: 'card', heading: '3. SAFETY/CONTROL', ... },
                { type: 'card', heading: '4. ALIGNMENT', ... },
            ]
        }
    ]
})
```

---

## Implementation Order

1. **Style Presets** - Foundation for everything else
   - Shadow presets (`sm`, `md`, `lg`)
   - Add `slide.defaults` cascade

2. **Layout Engine Core** - Calculate positions from specs
   - Grid layout
   - Stack layout

3. **Group Primitive** - Relative positioning within containers
   - Coordinate translation
   - Defaults cascade

4. **Card Component** - Most common high-level pattern
   - Built on Group + Styles

5. **Declarative Children API** - Optional declarative syntax
   - Slide accepts `children` array
   - Recursive rendering

---

## Backward Compatibility

Since breaking changes are acceptable:
- Old imperative API (`slide.addText()`, `slide.addShape()`) remains functional
- New compositional API is additive
- Can deprecate verbose options in favor of presets over time

---

## File Structure

```
src/
├── components/
│   ├── card.ts
│   ├── list.ts
│   └── index.ts
├── layout/
│   ├── layout-engine.ts
│   ├── grid.ts
│   ├── stack.ts
│   ├── flex.ts
│   └── index.ts
├── styles/
│   ├── presets.ts
│   ├── shadow-presets.ts
│   ├── typography-presets.ts
│   └── index.ts
├── group.ts
└── ... (existing files)
```

---

## Success Metrics

The original 100-line card example should become ~20 lines:

```typescript
const slide = pptx.addSlide({ background: 'FFFFFF' })

slide.addText('Four Open Problems in AI Design', {
    x: 0.5, y: 0.4,
    style: { fontSize: 32, bold: true, color: '2E9B7B' }
})

slide.addGrid({
    x: 0.6, y: 1.3,
    cols: 2, gap: 0.3,
    children: [
        slide.card({ heading: '1. LEARNING', headingColor: GOLD, body: '...' }),
        slide.card({ heading: '2. REASONING', headingColor: TEAL, body: '...' }),
        slide.card({ heading: '3. SAFETY/CONTROL', headingColor: GOLD, body: '...' }),
        slide.card({ heading: '4. ALIGNMENT', headingColor: TEAL, body: '...', background: BLUE }),
    ]
})
```
