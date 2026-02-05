/**
 * "Hybrid Voluntarism" slide - CLEAN API with proper font weights
 */

import PptxGenJS from '../dist/pptxgen.es.js'

const pptx = new PptxGenJS()
pptx.layout = 'LAYOUT_16x9'

const slide = pptx.addSlide()

// Set defaults
slide.fontFace = 'Outfit'
slide.color = 'FFFFFF'
slide.accentColor = '4FC3F7'

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

// Title - use ExtraBold weight for heavier appearance
slide.addText('Hybrid Voluntarism', {
    x: 0.6, y: 0.7, w: 9, h: 0.8,
    fontSize: 36,
    fontFace: 'Outfit ExtraBold',
    bold: true,
})

// Numbered list - use SemiBold for body text
slide.addNumberedList({
    x: 0.6,
    y: 1.6,
    w: 8.8,
    fontSize: 18,
    fontFace: 'Outfit SemiBold',
    itemGap: 0.4,
    items: [
        "Normative reasons can be either **'given'** or **'created'** reasons.",
        "We create reasons by **willing**, under the right conditions, that something is a reason.",
        "Thus we have **robust normative powers** — the power to will something to be a reason.",
    ],
})

await pptx.writeFile({ fileName: 'lab/hybrid-voluntarism-slide.pptx' })
console.log('Created: lab/hybrid-voluntarism-slide.pptx')
