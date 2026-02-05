/**
 * "Vis-à-vis Our Puzzle" slide - curved accent line using overlay technique
 */

import PptxGenJS from '../dist/pptxgen.es.js'

const pptx = new PptxGenJS()
pptx.layout = 'LAYOUT_16x9'

const slide = pptx.addSlide()

// ============================================================================
// BACKGROUND
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
// HELPER: Card with curved accent line (overlay technique)
// ============================================================================
function addCardWithCurvedAccent(slide, options) {
    const {
        x, y, w, h,
        accentColor,
        accentWidth = 0.06,
        backgroundColor,
        backgroundTransparency = 40,
        borderRadius = 0.15,
        borderColor = '3A4A5A',
        borderWidth = 1,
    } = options

    // 1. Draw outer rounded rect (accent color) - this creates the curved left edge
    slide.addShape(pptx.ShapeType.roundRect, {
        x,
        y,
        w,
        h,
        fill: { color: accentColor },
        line: { color: borderColor, width: borderWidth },
        rectRadius: borderRadius,
    })

    // 2. Draw inner rounded rect (card background) - inset from left
    slide.addShape(pptx.ShapeType.roundRect, {
        x: x + accentWidth,
        y,
        w: w - accentWidth,
        h,
        fill: { color: backgroundColor, transparency: backgroundTransparency },
        line: { color: borderColor, width: borderWidth },
        rectRadius: borderRadius,
    })

    // Return content area bounds for adding text
    return {
        contentX: x + accentWidth + 0.3,
        contentY: y + 0.3,
        contentW: w - accentWidth - 0.6,
        contentH: h - 0.6,
    }
}

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
// LEFT CARD
// ============================================================================
const leftCard = addCardWithCurvedAccent(slide, {
    x: 0.6,
    y: 1.7,
    w: 4.3,
    h: 3.0,
    accentColor: GREEN,
    backgroundColor: CARD_BG,
})

// Left card title
slide.addText('✓', {
    x: leftCard.contentX,
    y: leftCard.contentY,
    w: 0.4,
    h: 0.5,
    fontSize: 20,
    color: GREEN,
    align: 'center',
})

slide.addText("No puzzle about 'getting into' a choice situation", {
    x: leftCard.contentX + 0.4,
    y: leftCard.contentY,
    w: leftCard.contentW - 0.4,
    h: 0.7,
    fontSize: 18,
    fontFace: FONT,
    color: GREEN,
    align: 'center',
    italic: true,
})

// Left card body
slide.addText('We get into a choice situation randomly as it is one among a number of rationally eligible choice situations.', {
    x: leftCard.contentX,
    y: leftCard.contentY + 0.9,
    w: leftCard.contentW,
    h: 1.8,
    fontSize: 16,
    fontFace: FONT,
    color: GRAY,
    align: 'center',
})

// ============================================================================
// RIGHT CARD
// ============================================================================
const rightCard = addCardWithCurvedAccent(slide, {
    x: 5.1,
    y: 1.7,
    w: 4.3,
    h: 3.0,
    accentColor: PINK,
    backgroundColor: CARD_BG,
})

// Right card title
slide.addText('✓', {
    x: rightCard.contentX,
    y: rightCard.contentY,
    w: 0.4,
    h: 0.5,
    fontSize: 20,
    color: PINK,
    align: 'center',
})

slide.addText('No puzzle about justification', {
    x: rightCard.contentX + 0.4,
    y: rightCard.contentY,
    w: rightCard.contentW - 0.4,
    h: 0.7,
    fontSize: 18,
    fontFace: FONT,
    color: PINK,
    align: 'center',
    italic: true,
})

// Right card body
slide.addText('We are justified in being in it because we have sufficient reason to be in it.', {
    x: rightCard.contentX,
    y: rightCard.contentY + 0.9,
    w: rightCard.contentW,
    h: 1.8,
    fontSize: 16,
    fontFace: FONT,
    color: GRAY,
    align: 'center',
})

// ============================================================================
// SAVE
// ============================================================================
await pptx.writeFile({ fileName: 'lab/puzzle-slide-v2.pptx' })
console.log('Created: lab/puzzle-slide-v2.pptx')
