/**
 * "From Trichotomy to Quadchotomy" slide - using addStack + addPill with click animations
 */

import PptxGenJS from '../dist/pptxgen.es.js'

const pptx = new PptxGenJS()
pptx.layout = 'LAYOUT_16x9'

const slide = pptx.addSlide()
slide.fontFace = 'Outfit SemiBold'
slide.background = { color: 'FFFFFF' }

// Colors
const DARK_TEXT = '2D3436'
const GREEN = '2ECC71'
const BLUE = '3498DB'
const ORANGE = 'E67E22'
const TEAL = '1ABC9C'
const YELLOW_ORANGE = 'F5A623'
const PINK = 'E84A5F'

// Animation preset - fade in on click
const fadeOnClick = { type: 'fade', trigger: 'onClick', durationMs: 400 }

// Title
slide.addText('From Trichotomy to Quadchotomy', {
    x: 0.5, y: 0.4, w: 9, h: 0.7,
    fontSize: 32,
    fontFace: 'Outfit ExtraBold',
    color: DARK_TEXT,
})

// Two-column layout with stacked pills
slide.addTwoColumn({
    x: 0.5, y: 1.3, w: 9, h: 3.5,
    gap: 0.4,
    left: { ratio: 0.5 },

    renderLeft: (bounds) => {
        // Header
        slide.addText('Traditional View', {
            x: bounds.x, y: bounds.y, w: bounds.w, h: 0.5,
            fontSize: 18, color: DARK_TEXT, align: 'center',
        })

        // Pills using addStack - each pill fades in on click
        slide.addStack({ x: bounds.x, y: bounds.y + 0.6, w: bounds.w, gap: 0.15 }, (add) => {
            add.pill({ h: 0.65, text: 'Equally good', fill: GREEN, animation: fadeOnClick })
            add.pill({ h: 0.65, text: 'Better than', fill: BLUE, animation: fadeOnClick })
            add.pill({ h: 0.65, text: 'Worse than', fill: ORANGE, animation: fadeOnClick })
        })
    },

    renderRight: (bounds) => {
        // Header
        slide.addText('Parity View', {
            x: bounds.x, y: bounds.y, w: bounds.w, h: 0.5,
            fontSize: 18, color: TEAL, align: 'center',
        })

        // Pills using addStack - each pill fades in on click
        slide.addStack({ x: bounds.x, y: bounds.y + 0.6, w: bounds.w, gap: 0.15 }, (add) => {
            add.pill({ h: 0.65, text: 'Equally good', fill: TEAL, animation: fadeOnClick })
            add.pill({ h: 0.65, text: 'Better than', fill: TEAL, animation: fadeOnClick })
            add.pill({ h: 0.65, text: 'On a par', fill: YELLOW_ORANGE, animation: fadeOnClick })
            add.pill({ h: 0.65, text: 'Worse than', fill: PINK, animation: fadeOnClick })
        })
    },
})

await pptx.writeFile({ fileName: 'lab/trichotomy-slide.pptx' })
console.log('Created: lab/trichotomy-slide.pptx')
