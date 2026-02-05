/**
 * Test the new compositional API
 *
 * Compare this to test-replicate-slide.mjs - same output, much less code!
 */

import PptxGenJS from './dist/pptxgen.es.js'

const pptx = new PptxGenJS()
pptx.layout = 'LAYOUT_16x9'

const slide = pptx.addSlide()
slide.background = { color: 'FFFFFF' }

// Colors
const TITLE_GREEN = '2E9B7B'
const HEADING_GOLD = 'C5A636'
const HEADING_TEAL = '2E9B7B'
const CARD_BG_BLUE = 'E8F4FC'

// Title
slide.addText('Four Open Problems in AI Design', {
    x: 0.5,
    y: 0.4,
    w: 9,
    h: 0.7,
    fontSize: 32,
    bold: true,
    color: TITLE_GREEN,
    fontFace: 'Arial',
})

// Grid of cards - using the new compositional API!
slide.addCardGrid({
    x: 0.6,
    y: 1.3,
    cols: 2,
    gap: 0.3,
    cellWidth: 4.2,
    cellHeight: 1.5,
    cards: [
        {
            heading: '1. LEARNING',
            headingColor: HEADING_GOLD,
            body: 'How machines acquire knowledge from data.',
        },
        {
            heading: '2. REASONING',
            headingColor: HEADING_TEAL,
            body: 'How machines process information logically.',
        },
        {
            heading: '3. SAFETY/CONTROL',
            headingColor: HEADING_GOLD,
            body: 'Ensuring machines operate within bounds.',
        },
        {
            heading: '4. ALIGNMENT',
            headingColor: HEADING_TEAL,
            body: 'Ensuring machine goals match human values.',
            background: CARD_BG_BLUE,
        },
    ],
})

await pptx.writeFile({ fileName: 'new-api-slide.pptx' })
console.log('Created: new-api-slide.pptx')

// ============================================================================
// CODE COMPARISON:
// ============================================================================
//
// OLD API (test-replicate-slide.mjs):     ~100 lines
// NEW API (this file):                     ~55 lines (45% reduction!)
//
// And the new API is:
// - More readable (declarative structure)
// - Less error-prone (no manual coordinate calculations)
// - Easier to modify (change one card's position? change grid layout)
// - Self-documenting (structure reflects visual layout)
