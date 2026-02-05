/**
 * "Upshots for the Future" slide - using addCardGrid with rich text body
 */

import PptxGenJS, { textStyle } from '../dist/pptxgen.es.js'

const pptx = new PptxGenJS()
pptx.layout = 'LAYOUT_16x9'

const slide = pptx.addSlide()
slide.background = { color: 'FFFFFF' }

// ============================================================================
// COLORS & STYLES
// ============================================================================
const TEAL = '1ABC9C'
const DARK_TEXT = '2C3E50'
const GRAY_BG = 'F8F9FA'
const FONT = 'Arial'

const italic = textStyle({ italic: true })

// ============================================================================
// TITLE
// ============================================================================
slide.addText('Upshots for the Future', {
    x: 0.6,
    y: 0.4,
    w: 9,
    h: 0.8,
    fontSize: 44,
    fontFace: FONT,
    color: TEAL,
    bold: true,
})

// ============================================================================
// CARDS - Now using addCardGrid with rich text support
// ============================================================================
slide.addCardGrid({
    x: 0.6,
    y: 1.4,
    cols: 2,
    gap: { x: 0.4, y: 0.5 },
    cellWidth: 4.4,
    cellHeight: 1.8,
    cards: [
        {
            title: 'Alignment is Structural',
            titleColor: TEAL,
            titleFontSize: 14,
            body: 'It\'s not about "more data," but about the logic of value itself.',
            bodyColor: DARK_TEXT,
            bodyFontSize: 20,
            background: GRAY_BG,
            accentLine: TEAL,
            border: false,
            shadow: 'none',
            padding: 0.25,
        },
        {
            title: 'Value-Based Data',
            titleColor: TEAL,
            titleFontSize: 14,
            body: 'Machines must process evaluative data directly, not just proxies.',
            bodyColor: DARK_TEXT,
            bodyFontSize: 20,
            background: GRAY_BG,
            accentLine: TEAL,
            border: false,
            shadow: 'none',
            padding: 0.25,
        },
        {
            title: 'Meaning in Commitment',
            titleColor: TEAL,
            titleFontSize: 14,
            // Rich text body with italics!
            body: [
                { text: 'The human role shifts from ', options: { color: DARK_TEXT, fontSize: 20, fontFace: FONT, bold: true } },
                { text: 'calculating', options: { color: DARK_TEXT, fontSize: 20, fontFace: FONT, bold: true, italic: true } },
                { text: ' to ', options: { color: DARK_TEXT, fontSize: 20, fontFace: FONT, bold: true } },
                { text: 'creating', options: { color: DARK_TEXT, fontSize: 20, fontFace: FONT, bold: true, italic: true } },
                { text: ' value through choice.', options: { color: DARK_TEXT, fontSize: 20, fontFace: FONT, bold: true } },
            ],
            background: GRAY_BG,
            accentLine: TEAL,
            border: false,
            shadow: 'none',
            padding: 0.25,
        },
        {
            title: 'Carving at the Joints',
            titleColor: TEAL,
            titleFontSize: 14,
            body: 'A design that respects parity is a design that respects the human condition.',
            bodyColor: DARK_TEXT,
            bodyFontSize: 20,
            background: GRAY_BG,
            accentLine: TEAL,
            border: false,
            shadow: 'none',
            padding: 0.25,
        },
    ],
})

// ============================================================================
// SAVE
// ============================================================================
await pptx.writeFile({ fileName: 'lab/upshots-slide.pptx' })
console.log('Created: lab/upshots-slide.pptx')
