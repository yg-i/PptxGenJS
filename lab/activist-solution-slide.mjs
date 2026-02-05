/**
 * Exact replication of "The Activist Solution (cont.)" slide
 * Using addStack for automatic vertical positioning
 */

import PptxGenJS, { textStyle } from '../dist/pptxgen.es.js'

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
        angle: 135,
        stops: [
            { position: 0, color: '0A1628' },
            { position: 100, color: '1A3A4A' },
        ],
    },
}

// ============================================================================
// STYLES
// ============================================================================
const keyword = textStyle({ bold: true, color: '5DADE2' })
const FONT = 'Outfit'

// ============================================================================
// CONTENT - Using addStack for automatic Y positioning
// ============================================================================
slide.addStack({ x: 0.65, y: 0.5, w: 9, gap: 0.25 }, (add) => {
    // Title
    add.text('The Activist Solution (cont.)', {
        h: 0.7,
        fontSize: 40,
        fontFace: FONT,
        color: 'FFFFFF',
        bold: true,
    })

    add.space(0.1) // extra gap after title

    // Paragraph 1 (white text)
    add.richText({ h: 0.85, fontSize: 22, fontFace: FONT, color: 'FFFFFF' })`By ${keyword('committing')} to a choice situation, the agent is ${keyword('justified')} in being in it.`

    // Paragraph 2 (gray text)
    add.richText({ h: 0.85, fontSize: 22, fontFace: FONT, color: '9EAAB8' })`She ${keyword('makes')} being in that choice situation better for her with respect to having a meaningful life.`

    // Paragraph 3 (gray text)
    add.richText({ h: 1.2, fontSize: 22, fontFace: FONT, color: '9EAAB8' })`She thus has meaning in her life in virtue of her capacity to ${keyword('make it true')} through creating reasons that she has most reason to be in one choice situation rather than another.`
})

// ============================================================================
// SAVE
// ============================================================================
await pptx.writeFile({ fileName: 'lab/activist-solution-slide.pptx' })
console.log('Created: lab/activist-solution-slide.pptx')
