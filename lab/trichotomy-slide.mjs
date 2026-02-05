/**
 * Exact replication of "The Trichotomy Trap" slide
 *
 * Demonstrates all 4 new API improvements:
 * 1. addBulletList() - for lists with colored badges
 * 2. addBadge() - colored circles with text (used internally by addBulletList)
 * 3. Card title/align/bodyItalic - multi-section centered cards
 * 4. addTwoColumn() - two-column layout helper
 */

import PptxGenJS from '../dist/pptxgen.es.js'

const pptx = new PptxGenJS()
pptx.layout = 'LAYOUT_16x9'

const slide = pptx.addSlide()
slide.background = { color: 'FFFFFF' }

// ============================================================================
// COLORS (sampled from image)
// ============================================================================
const TITLE_BLUE = '1E88E5'
const TITLE_TEAL = '26A69A'
const BODY_GRAY = '555555'
const BADGE_BLUE = '29B6F6'    // Number 1 badge
const BADGE_RED = 'EF5350'     // Number 2 badge
const BADGE_GREEN = '66BB6A'   // Number 3 badge
const CARD_BORDER = '4DD0E1'   // Cyan border
const CARD_BG = 'F0FDFF'       // Very light cyan
const PARITY_COLOR = '26A69A'  // Teal for "PARITY"

// ============================================================================
// TITLE with gradient (Improvement from previous session)
// ============================================================================
slide.addTitle('The "Trichotomy" Trap', {
    x: 0.5,
    y: 0.3,
    w: 9,
    h: 0.7,
    fontSize: 36,
    gradient: { from: TITLE_BLUE, to: TITLE_TEAL },
})

// ============================================================================
// TWO-COLUMN LAYOUT (Improvement #4)
// ============================================================================
slide.addTwoColumn({
    x: 0.5,
    y: 1.1,
    w: 9.0,
    h: 3.2,
    gap: 0.5,
    left: { ratio: 0.48 },

    // LEFT COLUMN: Intro paragraph + bullet list
    renderLeft: (bounds) => {
        // Intro paragraph
        slide.addText('Standard logic assumes only 3 relations between two options A and B:', {
            x: bounds.x,
            y: bounds.y,
            w: bounds.w,
            h: 0.9,
            fontSize: 16,
            color: BODY_GRAY,
            fontFace: 'Arial',
            valign: 'top',
        })

        // Bullet list with colored badges (Improvement #1)
        slide.addBulletList({
            x: bounds.x,
            y: bounds.y + 1.0,
            w: bounds.w,
            itemHeight: 0.5,
            fontSize: 16,
            color: BODY_GRAY,
            badgeSize: 0.28,
            items: [
                { badge: { text: '1', color: BADGE_BLUE }, text: 'A is Better than B' },
                { badge: { text: '2', color: BADGE_RED }, text: 'A is Worse than B' },
                { badge: { text: '3', color: BADGE_GREEN }, text: 'A is Equally Good as B' },
            ],
        })
    },

    // RIGHT COLUMN: Card with PARITY (Improvement #3)
    renderRight: (bounds) => {
        slide.addCard({
            x: bounds.x,
            y: bounds.y,
            w: bounds.w,
            h: bounds.h,
            background: CARD_BG,
            border: { color: CARD_BORDER, width: 2 },
            shadow: 'none',
            padding: 0.3,
            align: 'center',  // All text centered

            // Title (small text above heading)
            title: 'The Missing Fourth Relation:',
            titleFontSize: 14,
            titleColor: BODY_GRAY,

            // Heading (large prominent text)
            heading: 'PARITY',
            headingFontSize: 32,
            headingColor: PARITY_COLOR,
            headingBold: true,
            headingLineHeight: 1.8,

            // Body (italic description)
            body: 'Options are comparable but qualitatively different. Neither is better, nor are they equal.',
            bodyFontSize: 14,
            bodyColor: BODY_GRAY,
            bodyItalic: true,
        })
    },
})

// ============================================================================
// SAVE
// ============================================================================
await pptx.writeFile({ fileName: 'lab/trichotomy-slide-v2.pptx' })
console.log('Created: lab/trichotomy-slide-v2.pptx')
