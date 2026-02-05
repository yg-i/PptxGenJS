/**
 * "Hybrid Voluntarism" slide - numbered list with highlighted keywords
 */

import PptxGenJS from '../dist/pptxgen.es.js'

const pptx = new PptxGenJS()
pptx.layout = 'LAYOUT_16x9'

const slide = pptx.addSlide()

// ============================================================================
// BACKGROUND - Dark gradient
// ============================================================================
slide.background = {
    type: 'gradient',
    gradient: {
        type: 'linear',
        angle: 180,
        stops: [
            { position: 0, color: '0A0F1A' },
            { position: 60, color: '0D1B2A' },
            { position: 100, color: '152238' },
        ],
    },
}

// ============================================================================
// COLORS
// ============================================================================
const CYAN = '4FC3F7'
const WHITE = 'FFFFFF'
const FONT = 'Arial'

// ============================================================================
// TITLE
// ============================================================================
slide.addText('Hybrid Voluntarism', {
    x: 0.6,
    y: 0.7,
    w: 9,
    h: 0.8,
    fontSize: 36,
    fontFace: FONT,
    color: WHITE,
    bold: true,
})

// ============================================================================
// NUMBERED LIST
// ============================================================================
const listX = 0.6
const listW = 8.8
const fontSize = 18
const lineHeight = 22  // approximate line height in points

// Item 1
slide.addText([
    { text: '1.', options: { color: CYAN, bold: true, fontSize } },
    { text: '    Normative reasons can be either ', options: { color: WHITE, fontSize } },
    { text: "'given'", options: { color: CYAN, bold: true, fontSize } },
    { text: ' or ', options: { color: WHITE, fontSize } },
    { text: "'created'", options: { color: CYAN, bold: true, fontSize } },
    { text: ' reasons.', options: { color: WHITE, fontSize } },
], {
    x: listX,
    y: 1.6,
    w: listW,
    h: 0.5,
    fontFace: FONT,
    valign: 'top',
})

// Item 2
slide.addText([
    { text: '2.', options: { color: CYAN, bold: true, fontSize } },
    { text: '    We create reasons by ', options: { color: WHITE, fontSize } },
    { text: 'willing', options: { color: CYAN, bold: true, fontSize } },
    { text: ', under the right conditions, that something is a reason.', options: { color: WHITE, fontSize } },
], {
    x: listX,
    y: 2.2,
    w: listW,
    h: 0.8,
    fontFace: FONT,
    valign: 'top',
})

// Item 3
slide.addText([
    { text: '3.', options: { color: CYAN, bold: true, fontSize } },
    { text: '    Thus we have ', options: { color: WHITE, fontSize } },
    { text: 'robust normative powers', options: { color: CYAN, bold: true, fontSize } },
    { text: ' \u2014 the power to will something to be a reason.', options: { color: WHITE, fontSize } },
], {
    x: listX,
    y: 2.9,
    w: listW,
    h: 0.8,
    fontFace: FONT,
    valign: 'top',
})

// ============================================================================
// SAVE
// ============================================================================
await pptx.writeFile({ fileName: 'lab/hybrid-voluntarism-slide.pptx' })
console.log('Created: lab/hybrid-voluntarism-slide.pptx')
