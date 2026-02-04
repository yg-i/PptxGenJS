/**
 * Advanced Morph Transition Tests
 *
 * Morph has three modes controlled by the 'morphOption' attribute:
 * - byObject (default): Morphs matching objects between slides
 * - byWord: Morphs text word-by-word (great for text transformations)
 * - byChar: Morphs text character-by-character (cinematic effect)
 *
 * Objects are matched by their 'objectName' property - identical names = morph together
 * Special prefix: !! forces objects to morph even if they're different types
 */
import PptxGenJS from './src/bld/pptxgen.es.js'

const pptx = new PptxGenJS()
pptx.title = 'Advanced Morph Transitions'

// =============================================================================
// DEMO 1: Morph by Object - Shape Position/Size Animation
// =============================================================================

// Slide 1A: Shape in top-left, small
const slide1a = pptx.addSlide()
slide1a.transition = { type: 'fade' } // First slide uses fade
slide1a.addText('Morph Demo: Position & Size', { x: 0.5, y: 0.3, w: 9, h: 0.5, fontSize: 24, bold: true })
slide1a.addText('Click to see the shape morph to a new position and size', { x: 0.5, y: 0.8, w: 9, h: 0.4, fontSize: 14, color: '666666' })
// Named shape - will morph with same-named shape on next slide
slide1a.addShape('rect', {
	x: 1, y: 2, w: 1.5, h: 1.5,
	fill: { color: '4472C4' },
	objectName: 'MorphBox'  // This name links it to the next slide's shape
})

// Slide 1B: Same shape, different position/size - MORPH!
const slide1b = pptx.addSlide()
slide1b.transition = { type: 'morph', durationMs: 1500 }
slide1b.addText('Morph Demo: Position & Size', { x: 0.5, y: 0.3, w: 9, h: 0.5, fontSize: 24, bold: true })
slide1b.addText('The shape smoothly animated to its new position!', { x: 0.5, y: 0.8, w: 9, h: 0.4, fontSize: 14, color: '666666' })
// Same name = morphs from previous slide
slide1b.addShape('rect', {
	x: 6, y: 2.5, w: 3, h: 2,
	fill: { color: '4472C4' },
	objectName: 'MorphBox'
})

// =============================================================================
// DEMO 2: Morph by Word - Text Transformation
// This is HARD to do in PowerPoint UI - requires XML control!
// =============================================================================

const slide2a = pptx.addSlide()
slide2a.transition = { type: 'morph', morphOption: 'byWord', durationMs: 1500 }
slide2a.addText('Morph by Word', { x: 0.5, y: 0.3, w: 9, h: 0.5, fontSize: 24, bold: true })
slide2a.addText('Each word morphs independently!', { x: 0.5, y: 0.8, w: 9, h: 0.4, fontSize: 14, color: '666666' })
// Text in one arrangement
slide2a.addText('Hello World Today', {
	x: 1, y: 2, w: 8, h: 1.5,
	fontSize: 48, bold: true, color: '2E7D32',
	objectName: 'MorphText'
})

const slide2b = pptx.addSlide()
slide2b.transition = { type: 'morph', morphOption: 'byWord', durationMs: 1500 }
slide2b.addText('Morph by Word', { x: 0.5, y: 0.3, w: 9, h: 0.5, fontSize: 24, bold: true })
slide2b.addText('Words rearranged and restyled!', { x: 0.5, y: 0.8, w: 9, h: 0.4, fontSize: 14, color: '666666' })
// Same text, different arrangement - words will morph to new positions
slide2b.addText('Today World Hello', {
	x: 1, y: 2.5, w: 8, h: 1.5,
	fontSize: 48, bold: true, color: 'C62828',
	objectName: 'MorphText'
})

// =============================================================================
// DEMO 3: Morph by Character - Cinematic Text Effect
// This creates a dramatic letter-by-letter animation!
// =============================================================================

const slide3a = pptx.addSlide()
slide3a.transition = { type: 'morph', morphOption: 'byChar', durationMs: 2000 }
slide3a.addText('Morph by Character', { x: 0.5, y: 0.3, w: 9, h: 0.5, fontSize: 24, bold: true })
slide3a.addText('Watch each letter morph individually!', { x: 0.5, y: 0.8, w: 9, h: 0.4, fontSize: 14, color: '666666' })
slide3a.addText('ABCDEF', {
	x: 2, y: 2, w: 6, h: 1.5,
	fontSize: 72, bold: true, color: '1565C0',
	align: 'center',
	objectName: 'CharMorph'
})

const slide3b = pptx.addSlide()
slide3b.transition = { type: 'morph', morphOption: 'byChar', durationMs: 2000 }
slide3b.addText('Morph by Character', { x: 0.5, y: 0.3, w: 9, h: 0.5, fontSize: 24, bold: true })
slide3b.addText('Letters shuffled and transformed!', { x: 0.5, y: 0.8, w: 9, h: 0.4, fontSize: 14, color: '666666' })
slide3b.addText('FEDCBA', {
	x: 2, y: 2, w: 6, h: 1.5,
	fontSize: 72, bold: true, color: 'D84315',
	align: 'center',
	objectName: 'CharMorph'
})

// =============================================================================
// DEMO 4: Morph by Object - Multiple Objects Choreographed
// =============================================================================

const slide4a = pptx.addSlide()
slide4a.transition = { type: 'morph', durationMs: 1200 }
slide4a.addText('Morph Demo: Multiple Objects', { x: 0.5, y: 0.3, w: 9, h: 0.5, fontSize: 24, bold: true })
// Three shapes in a row
slide4a.addShape('rect', { x: 1, y: 2.5, w: 1.5, h: 1.5, fill: { color: 'E74C3C' }, objectName: 'Box1' })
slide4a.addShape('rect', { x: 4, y: 2.5, w: 1.5, h: 1.5, fill: { color: '3498DB' }, objectName: 'Box2' })
slide4a.addShape('rect', { x: 7, y: 2.5, w: 1.5, h: 1.5, fill: { color: '2ECC71' }, objectName: 'Box3' })

const slide4b = pptx.addSlide()
slide4b.transition = { type: 'morph', durationMs: 1200 }
slide4b.addText('Morph Demo: Multiple Objects', { x: 0.5, y: 0.3, w: 9, h: 0.5, fontSize: 24, bold: true })
// Rearranged in a triangle
slide4b.addShape('rect', { x: 4, y: 1.5, w: 1.5, h: 1.5, fill: { color: 'E74C3C' }, objectName: 'Box1' })
slide4b.addShape('rect', { x: 2, y: 3.5, w: 1.5, h: 1.5, fill: { color: '3498DB' }, objectName: 'Box2' })
slide4b.addShape('rect', { x: 6, y: 3.5, w: 1.5, h: 1.5, fill: { color: '2ECC71' }, objectName: 'Box3' })

// =============================================================================
// DEMO 5: Morph Shape Type Change (using !! prefix convention)
// =============================================================================

const slide5a = pptx.addSlide()
slide5a.transition = { type: 'morph', durationMs: 1500 }
slide5a.addText('Morph Demo: Shape Transformation', { x: 0.5, y: 0.3, w: 9, h: 0.5, fontSize: 24, bold: true })
slide5a.addText('Watch the circle become a star!', { x: 0.5, y: 0.8, w: 9, h: 0.4, fontSize: 14, color: '666666' })
slide5a.addShape('ellipse', { x: 3.5, y: 2, w: 3, h: 3, fill: { color: 'F39C12' }, objectName: '!!Transform' })

const slide5b = pptx.addSlide()
slide5b.transition = { type: 'morph', durationMs: 1500 }
slide5b.addText('Morph Demo: Shape Transformation', { x: 0.5, y: 0.3, w: 9, h: 0.5, fontSize: 24, bold: true })
slide5b.addText('The circle morphed into a star!', { x: 0.5, y: 0.8, w: 9, h: 0.4, fontSize: 14, color: '666666' })
slide5b.addShape('star5', { x: 3.5, y: 2, w: 3, h: 3, fill: { color: 'F39C12' }, objectName: '!!Transform' })

// =============================================================================
// DEMO 6: Morph by Character - Word Reveal Effect
// This creates a "typing" or "unscrambling" effect - very cinematic!
// =============================================================================

const slide6a = pptx.addSlide()
slide6a.transition = { type: 'morph', morphOption: 'byChar', durationMs: 2500 }
slide6a.addText('Character Reveal Effect', { x: 0.5, y: 0.3, w: 9, h: 0.5, fontSize: 24, bold: true })
slide6a.addText('Scrambled letters morph into readable text', { x: 0.5, y: 0.8, w: 9, h: 0.4, fontSize: 14, color: '666666' })
slide6a.addText('XYZQKP', {
	x: 1, y: 2, w: 8, h: 2,
	fontSize: 96, bold: true, color: '7B1FA2',
	align: 'center',
	objectName: 'RevealText'
})

const slide6b = pptx.addSlide()
slide6b.transition = { type: 'morph', morphOption: 'byChar', durationMs: 2500 }
slide6b.addText('Character Reveal Effect', { x: 0.5, y: 0.3, w: 9, h: 0.5, fontSize: 24, bold: true })
slide6b.addText('The letters morphed into "MORPH"!', { x: 0.5, y: 0.8, w: 9, h: 0.4, fontSize: 14, color: '666666' })
slide6b.addText('MORPH!', {
	x: 1, y: 2, w: 8, h: 2,
	fontSize: 96, bold: true, color: '00897B',
	align: 'center',
	objectName: 'RevealText'
})

// Save
pptx.writeFile({ fileName: 'test-morph.pptx' })
	.then(() => console.log('Created: test-morph.pptx'))
	.catch(err => console.error('Error:', err))
