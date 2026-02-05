/**
 * Exact replication of "Four Open Problems in AI Design" slide
 *
 * Note: The original title has a gradient effect (blue → teal).
 * Since PptxGenJS doesn't support text gradients natively,
 * we simulate it by splitting the title into segments with transitioning colors.
 */

import PptxGenJS from '../dist/pptxgen.es.js'

const pptx = new PptxGenJS()
pptx.layout = 'LAYOUT_16x9'

const slide = pptx.addSlide()
slide.background = { color: 'FFFFFF' }

// ============================================================================
// EXACT COLORS (sampled from image)
// ============================================================================
// Title gradient: blue on left → teal on right
const TITLE_BLUE = '1E88E5'       // Left side of title (more blue)
const TITLE_TEAL = '26A69A'       // Right side of title (teal/green)
const HEADING_GOLD = 'C5A636'     // Headings 1, 3 (gold/yellow)
const BODY_GRAY = '666666'        // Description text (neutral gray)
const CARD_BG_WHITE = 'FFFFFF'    // Cards 1-3 background (white)
const CARD_BG_BLUE = 'E3F2FD'     // Card 4 background (light blue highlight)
const CARD_BORDER = 'E0E0E0'      // Subtle gray border

// ============================================================================
// LAYOUT MEASUREMENTS (inches)
// ============================================================================
const SLIDE_PADDING_X = 0.75
const TITLE_Y = 0.45
const CARDS_START_Y = 1.3
const CARD_WIDTH = 4.15
const CARD_HEIGHT = 1.5
const CARD_GAP_X = 0.35
const CARD_GAP_Y = 0.25
const CARD_PADDING = 0.25
const CARD_RADIUS = 0.06

// ============================================================================
// TITLE with simulated gradient (split into color segments)
// ============================================================================
// Colors interpolated from blue (#1E88E5) to teal (#26A69A)
const titleSegments = [
    { text: 'Four ', options: { color: '1E88E5', bold: true, fontSize: 30, fontFace: 'Arial' } },
    { text: 'Open ', options: { color: '219F94', bold: true, fontSize: 30, fontFace: 'Arial' } },
    { text: 'Problems ', options: { color: '23A08F', bold: true, fontSize: 30, fontFace: 'Arial' } },
    { text: 'in ', options: { color: '25A298', bold: true, fontSize: 30, fontFace: 'Arial' } },
    { text: 'AI ', options: { color: '26A499', bold: true, fontSize: 30, fontFace: 'Arial' } },
    { text: 'Design', options: { color: '26A69A', bold: true, fontSize: 30, fontFace: 'Arial' } },
]

slide.addText(titleSegments, {
    x: SLIDE_PADDING_X,
    y: TITLE_Y,
    w: 9,
    h: 0.6,
})

// ============================================================================
// CARD HELPER
// ============================================================================
function addCard(x, y, headingNumber, headingText, headingColor, bodyText, bgColor) {
    // Card background (rounded rectangle with subtle shadow)
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
            blur: 3,
            offset: 1.5,
            angle: 90,
            color: '000000',
            opacity: 0.1,
        },
    })

    // Heading (e.g., "1. LEARNING")
    slide.addText(`${headingNumber}. ${headingText}`, {
        x: x + CARD_PADDING,
        y: y + CARD_PADDING,
        w: CARD_WIDTH - CARD_PADDING * 2,
        h: 0.4,
        fontSize: 15,
        bold: true,
        color: headingColor,
        fontFace: 'Arial',
    })

    // Body text
    slide.addText(bodyText, {
        x: x + CARD_PADDING,
        y: y + CARD_PADDING + 0.55,
        w: CARD_WIDTH - CARD_PADDING * 2,
        h: 0.7,
        fontSize: 13,
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
    'How machines acquire knowledge from data.', CARD_BG_WHITE)

addCard(col2X, row1Y, '2', 'REASONING', TITLE_TEAL,
    'How machines process information logically.', CARD_BG_WHITE)

// Row 2
addCard(col1X, row2Y, '3', 'SAFETY/CONTROL', HEADING_GOLD,
    'Ensuring machines operate within bounds.', CARD_BG_WHITE)

addCard(col2X, row2Y, '4', 'ALIGNMENT', TITLE_TEAL,
    'Ensuring machine goals match human values.', CARD_BG_BLUE)

// ============================================================================
// SAVE
// ============================================================================
await pptx.writeFile({ fileName: 'lab/ai-problems-slide.pptx' })
console.log('Created: lab/ai-problems-slide.pptx')
