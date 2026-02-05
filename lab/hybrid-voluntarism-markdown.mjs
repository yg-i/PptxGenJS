/**
 * "Hybrid Voluntarism" slide - using addMarkdownText for even simpler syntax
 *
 * This demonstrates the markdown-lite API where **bold** text automatically
 * gets styled with a different color.
 */

import PptxGenJS from '../dist/pptxgen.es.js'

const pptx = new PptxGenJS()
pptx.layout = 'LAYOUT_16x9'

const slide = pptx.addSlide()

// Background
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

const CYAN = '4FC3F7'
const WHITE = 'FFFFFF'

// Title
slide.addText('Hybrid Voluntarism', {
    x: 0.6, y: 0.7, w: 9, h: 0.8,
    fontSize: 36, fontFace: 'Arial', color: WHITE, bold: true,
})

// Using addMarkdownText - **bold** and 'quoted' text get colored automatically!
const listY = 1.6
const listGap = 0.7

slide.addMarkdownText("1.    Normative reasons can be either **'given'** or **'created'** reasons.", {
    x: 0.6, y: listY, w: 8.8, h: 0.5,
    fontSize: 18, fontFace: 'Arial', color: WHITE, boldColor: CYAN,
})

slide.addMarkdownText("2.    We create reasons by **willing**, under the right conditions, that something is a reason.", {
    x: 0.6, y: listY + listGap, w: 8.8, h: 0.8,
    fontSize: 18, fontFace: 'Arial', color: WHITE, boldColor: CYAN,
})

slide.addMarkdownText("3.    Thus we have **robust normative powers** \u2014 the power to will something to be a reason.", {
    x: 0.6, y: listY + listGap * 2, w: 8.8, h: 0.8,
    fontSize: 18, fontFace: 'Arial', color: WHITE, boldColor: CYAN,
})

await pptx.writeFile({ fileName: 'lab/hybrid-voluntarism-markdown.pptx' })
console.log('Created: lab/hybrid-voluntarism-markdown.pptx')
