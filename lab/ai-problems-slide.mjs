/**
 * Exact replication of "Four Open Problems in AI Design" slide
 *
 * Demonstrates all 7 API improvements:
 * 1. Gradient text support (addTitle with gradient, addGradientText)
 * 2. Card border: false support
 * 3. Shadow presets with overrides ({ preset: 'sm', opacity: 0.08 })
 * 4. addTitle() convenience method
 * 5. headingLineHeight control
 * 6. interpolateColors utility (used internally)
 * 7. highlight property for cards
 */

import PptxGenJS from '../dist/pptxgen.es.js'

const pptx = new PptxGenJS()
pptx.layout = 'LAYOUT_16x9'

const slide = pptx.addSlide()
slide.background = { color: 'FFFFFF' }

// ============================================================================
// COLORS
// ============================================================================
const TITLE_BLUE = '1E88E5'       // Gradient start
const TITLE_TEAL = '26A69A'       // Gradient end
const HEADING_GOLD = 'C5A636'     // Headings 1, 3
const BODY_GRAY = '666666'        // Description text
const CARD_BG = 'FFFFFF'          // White card background

// ============================================================================
// TITLE with gradient (Improvement #1 & #4)
// ============================================================================
slide.addTitle('Four Open Problems in AI Design', {
    x: 0.75,
    y: 0.45,
    w: 9,
    h: 0.6,
    fontSize: 30,
    gradient: { from: TITLE_BLUE, to: TITLE_TEAL },
})

// ============================================================================
// CARDS using compositional API with all improvements
// ============================================================================
slide.addCardGrid({
    x: 0.75,
    y: 1.3,
    cols: 2,
    gap: { x: 0.35, y: 0.25 },
    cellWidth: 4.15,
    cellHeight: 1.5,
    cards: [
        {
            heading: '1. LEARNING',
            headingColor: HEADING_GOLD,
            headingLineHeight: 1.4, // Improvement #5
            body: 'How machines acquire knowledge from data.',
            bodyColor: BODY_GRAY,
            background: CARD_BG,
            border: false, // Improvement #2: No border
            shadow: { preset: 'subtle', opacity: 0.1 }, // Improvement #3: Preset with override
        },
        {
            heading: '2. REASONING',
            headingColor: TITLE_TEAL,
            headingLineHeight: 1.4,
            body: 'How machines process information logically.',
            bodyColor: BODY_GRAY,
            background: CARD_BG,
            border: false,
            shadow: { preset: 'subtle', opacity: 0.1 },
        },
        {
            heading: '3. SAFETY/CONTROL',
            headingColor: HEADING_GOLD,
            headingLineHeight: 1.4,
            body: 'Ensuring machines operate within bounds.',
            bodyColor: BODY_GRAY,
            background: CARD_BG,
            border: false,
            shadow: { preset: 'subtle', opacity: 0.1 },
        },
        {
            heading: '4. ALIGNMENT',
            headingColor: TITLE_TEAL,
            headingLineHeight: 1.4,
            body: 'Ensuring machine goals match human values.',
            bodyColor: BODY_GRAY,
            highlight: true, // Improvement #7: Highlighted card (uses default highlight color)
            border: false,
            shadow: { preset: 'subtle', opacity: 0.1 },
        },
    ],
})

// ============================================================================
// SAVE
// ============================================================================
await pptx.writeFile({ fileName: 'lab/ai-problems-slide-v2.pptx' })
console.log('Created: lab/ai-problems-slide-v2.pptx')
