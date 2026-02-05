/**
 * "Vis-à-vis Our Puzzle" slide - semi-transparent cards with left accent lines
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
            { position: 0, color: '0D1B2A' },
            { position: 100, color: '1B263B' },
        ],
    },
}

// ============================================================================
// COLORS
// ============================================================================
const GREEN = '4ECDC4'
const PINK = 'FF8A8A'
const CARD_BG = '1A2332'
const WHITE = 'FFFFFF'
const GRAY = 'B0B8C4'
const FONT = 'Georgia'

// ============================================================================
// TITLE
// ============================================================================
slide.addText('Vis-à-vis Our Puzzle', {
    x: 0.5,
    y: 0.5,
    w: 9,
    h: 0.9,
    fontSize: 44,
    fontFace: FONT,
    color: WHITE,
    bold: true,
    italic: true,
    align: 'center',
})

// ============================================================================
// LEFT CARD - Green accent
// ============================================================================
slide.addCard({
    x: 0.6,
    y: 1.7,
    w: 4.3,
    h: 3.0,
    background: { color: CARD_BG, transparency: 40 },
    border: { color: '3A4A5A', width: 1 },
    borderRadius: 0.15,
    accentLine: { color: GREEN, position: 'left', thickness: 0.06 },
    shadow: 'none',
    padding: 0.35,
    align: 'center',
    // Title with checkmark
    title: '✓     No puzzle about \'getting into\' a choice situation',
    titleColor: GREEN,
    titleFontSize: 18,
    titleFontFace: FONT,
    // Body
    body: 'We get into a choice situation randomly as it is one among a number of rationally eligible choice situations.',
    bodyColor: GRAY,
    bodyFontSize: 16,
    bodyFontFace: FONT,
})

// ============================================================================
// RIGHT CARD - Pink accent
// ============================================================================
slide.addCard({
    x: 5.1,
    y: 1.7,
    w: 4.3,
    h: 3.0,
    background: { color: CARD_BG, transparency: 40 },
    border: { color: '3A4A5A', width: 1 },
    borderRadius: 0.15,
    accentLine: { color: PINK, position: 'left', thickness: 0.06 },
    shadow: 'none',
    padding: 0.35,
    align: 'center',
    // Title with checkmark
    title: '✓     No puzzle about justification',
    titleColor: PINK,
    titleFontSize: 18,
    titleFontFace: FONT,
    // Body
    body: 'We are justified in being in it because we have sufficient reason to be in it.',
    bodyColor: GRAY,
    bodyFontSize: 16,
    bodyFontFace: FONT,
})

// ============================================================================
// SAVE
// ============================================================================
await pptx.writeFile({ fileName: 'lab/puzzle-slide.pptx' })
console.log('Created: lab/puzzle-slide.pptx')
