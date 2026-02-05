/**
 * Exact reproduction of the AI Design slide
 * Aiming for precise color and layout matching
 */

import PptxGenJS from '../dist/pptxgen.es.js'

const pptx = new PptxGenJS()
pptx.layout = 'LAYOUT_16x9'

const slide = pptx.addSlide()
slide.background = { color: 'FFFFFF' }

// ============================================================================
// EXACT COLORS (sampled from image)
// ============================================================================
const TITLE_TEAL = '2A9D8F'       // Title and headings 2, 4
const HEADING_GOLD = 'D4A72C'     // Headings 1, 3 (warmer yellow)
const BODY_GRAY = '5F6368'        // Description text (Google-style gray)
const CARD_BG_DEFAULT = 'F8F9FA'  // Very light gray (cards 1-3)
const CARD_BG_BLUE = 'E8F4FD'     // Light blue (card 4)
const CARD_BORDER = 'E8EAED'      // Subtle border

// ============================================================================
// LAYOUT MEASUREMENTS (inches, estimated from image proportions)
// ============================================================================
const SLIDE_PADDING_X = 0.8
const TITLE_Y = 0.5
const CARDS_START_Y = 1.4
const CARD_WIDTH = 4.0
const CARD_HEIGHT = 1.4
const CARD_GAP_X = 0.4
const CARD_GAP_Y = 0.35
const CARD_PADDING = 0.25
const CARD_RADIUS = 0.08

// ============================================================================
// TITLE
// ============================================================================
slide.addText('Four Open Problems in AI Design', {
    x: SLIDE_PADDING_X,
    y: TITLE_Y,
    w: 8,
    h: 0.6,
    fontSize: 28,
    bold: true,
    color: TITLE_TEAL,
    fontFace: 'Arial',
})

// ============================================================================
// CARD HELPER
// ============================================================================
function addCard(x, y, headingNumber, headingText, headingColor, bodyText, bgColor) {
    // Card background
    slide.addShape(pptx.ShapeType.roundRect, {
        x,
        y,
        w: CARD_WIDTH,
        h: CARD_HEIGHT,
        fill: { color: bgColor },
        line: { color: CARD_BORDER, width: 0.5 },
        rectRadius: CARD_RADIUS,
        shadow: {
            type: 'outer',
            blur: 4,
            offset: 1,
            angle: 90,
            color: '000000',
            opacity: 0.08,
        },
    })

    // Heading (e.g., "1. LEARNING")
    slide.addText(`${headingNumber}. ${headingText}`, {
        x: x + CARD_PADDING,
        y: y + CARD_PADDING,
        w: CARD_WIDTH - CARD_PADDING * 2,
        h: 0.35,
        fontSize: 14,
        bold: true,
        color: headingColor,
        fontFace: 'Arial',
    })

    // Body text
    slide.addText(bodyText, {
        x: x + CARD_PADDING,
        y: y + CARD_PADDING + 0.45,
        w: CARD_WIDTH - CARD_PADDING * 2,
        h: 0.7,
        fontSize: 12,
        color: BODY_GRAY,
        fontFace: 'Arial',
        valign: 'top',
    })
}

// ============================================================================
// CARDS (2x2 grid)
// ============================================================================
const col1X = SLIDE_PADDING_X
const col2X = SLIDE_PADDING_X + CARD_WIDTH + CARD_GAP_X
const row1Y = CARDS_START_Y
const row2Y = CARDS_START_Y + CARD_HEIGHT + CARD_GAP_Y

// Row 1
addCard(col1X, row1Y, '1', 'LEARNING', HEADING_GOLD,
    'How machines acquire knowledge from data.', CARD_BG_DEFAULT)

addCard(col2X, row1Y, '2', 'REASONING', TITLE_TEAL,
    'How machines process information logically.', CARD_BG_DEFAULT)

// Row 2
addCard(col1X, row2Y, '3', 'SAFETY/CONTROL', HEADING_GOLD,
    'Ensuring machines operate within bounds.', CARD_BG_DEFAULT)

addCard(col2X, row2Y, '4', 'ALIGNMENT', TITLE_TEAL,
    'Ensuring machine goals match human values.', CARD_BG_BLUE)

// ============================================================================
// SAVE
// ============================================================================
await pptx.writeFile({ fileName: 'lab/exact-slide.pptx' })
console.log('Created: lab/exact-slide.pptx')
