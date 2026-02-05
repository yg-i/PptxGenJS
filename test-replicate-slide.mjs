import PptxGenJS from './dist/pptxgen.es.js'

const pptx = new PptxGenJS()
pptx.layout = 'LAYOUT_16x9'

const slide = pptx.addSlide()
slide.background = { color: 'FFFFFF' }

// Colors from the image
const TITLE_GREEN = '2E9B7B'      // Title "Four Open Problems in AI Design"
const HEADING_GOLD = 'C5A636'     // 1. LEARNING, 3. SAFETY/CONTROL
const HEADING_TEAL = '2E9B7B'     // 2. REASONING, 4. ALIGNMENT
const DESCRIPTION_GRAY = '555555' // Description text
const CARD_BG_LIGHT = 'F5F5F5'    // Light gray card background
const CARD_BG_BLUE = 'E8F4FC'     // Card 4 has light blue background
const CARD_BORDER = 'E0E0E0'      // Light border

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

// Card dimensions
const cardWidth = 4.2
const cardHeight = 1.5
const cardGap = 0.3
const startX = 0.6
const startY = 1.3
const col2X = startX + cardWidth + cardGap
const row2Y = startY + cardHeight + cardGap

// Helper to create a card
function addCard(slide, x, y, headingNum, headingText, headingColor, description, bgColor) {
    // Card background (rounded rectangle with shadow)
    slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
        x,
        y,
        w: cardWidth,
        h: cardHeight,
        fill: { color: bgColor },
        line: { color: CARD_BORDER, width: 1 },
        rectRadius: 0.1,
        shadow: {
            type: 'outer',
            blur: 4,
            offset: 2,
            angle: 45,
            color: '000000',
            opacity: 0.15,
        },
    })

    // Heading text (e.g., "1. LEARNING")
    slide.addText(`${headingNum}. ${headingText}`, {
        x: x + 0.2,
        y: y + 0.15,
        w: cardWidth - 0.4,
        h: 0.45,
        fontSize: 16,
        bold: true,
        color: headingColor,
        fontFace: 'Arial',
    })

    // Description text
    slide.addText(description, {
        x: x + 0.2,
        y: y + 0.6,
        w: cardWidth - 0.4,
        h: 0.75,
        fontSize: 13,
        color: DESCRIPTION_GRAY,
        fontFace: 'Arial',
    })
}

// Card 1: Learning (top-left)
addCard(
    slide,
    startX,
    startY,
    '1',
    'LEARNING',
    HEADING_GOLD,
    'How machines acquire knowledge from data.',
    CARD_BG_LIGHT
)

// Card 2: Reasoning (top-right)
addCard(
    slide,
    col2X,
    startY,
    '2',
    'REASONING',
    HEADING_TEAL,
    'How machines process information logically.',
    CARD_BG_LIGHT
)

// Card 3: Safety/Control (bottom-left)
addCard(
    slide,
    startX,
    row2Y,
    '3',
    'SAFETY/CONTROL',
    HEADING_GOLD,
    'Ensuring machines operate within bounds.',
    CARD_BG_LIGHT
)

// Card 4: Alignment (bottom-right) - has blue background
addCard(
    slide,
    col2X,
    row2Y,
    '4',
    'ALIGNMENT',
    HEADING_TEAL,
    'Ensuring machine goals match human values.',
    CARD_BG_BLUE
)

// Save the file
await pptx.writeFile({ fileName: 'replicated-slide.pptx' })
console.log('Created: replicated-slide.pptx')
