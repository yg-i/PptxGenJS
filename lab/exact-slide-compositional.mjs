/**
 * Same slide using the compositional API
 */

import PptxGenJS from '../dist/pptxgen.es.js'

const pptx = new PptxGenJS()
pptx.layout = 'LAYOUT_16x9'

const slide = pptx.addSlide()
slide.background = { color: 'FFFFFF' }

// Colors
const TEAL = '2A9D8F'
const GOLD = 'D4A72C'
const BODY_GRAY = '5F6368'
const CARD_BG = 'F8F9FA'
const CARD_BLUE = 'E8F4FD'

// Title
slide.addText('Four Open Problems in AI Design', {
    x: 0.8, y: 0.5, w: 8, h: 0.6,
    fontSize: 28, bold: true, color: TEAL, fontFace: 'Arial',
})

// Cards using compositional API
slide.addCardGrid({
    x: 0.8,
    y: 1.4,
    cols: 2,
    gap: { x: 0.4, y: 0.35 },
    cellWidth: 4.0,
    cellHeight: 1.4,
    cards: [
        { heading: '1. LEARNING', headingColor: GOLD, body: 'How machines acquire knowledge from data.', background: CARD_BG, bodyColor: BODY_GRAY },
        { heading: '2. REASONING', headingColor: TEAL, body: 'How machines process information logically.', background: CARD_BG, bodyColor: BODY_GRAY },
        { heading: '3. SAFETY/CONTROL', headingColor: GOLD, body: 'Ensuring machines operate within bounds.', background: CARD_BG, bodyColor: BODY_GRAY },
        { heading: '4. ALIGNMENT', headingColor: TEAL, body: 'Ensuring machine goals match human values.', background: CARD_BLUE, bodyColor: BODY_GRAY },
    ],
})

await pptx.writeFile({ fileName: 'lab/exact-slide-compositional.pptx' })
console.log('Created: lab/exact-slide-compositional.pptx')
