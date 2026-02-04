/**
 * Minimal Morph Test - 2 slides only
 */
import PptxGenJS from './src/bld/pptxgen.es.js'

const pptx = new PptxGenJS()
pptx.title = 'Simple Morph Test'

// Slide 1: Shape on the left
const slide1 = pptx.addSlide()
slide1.addText('Slide 1 - Shape on LEFT', { x: 0.5, y: 0.3, w: 9, h: 0.5, fontSize: 24 })
slide1.addShape('rect', {
	x: 1, y: 2, w: 2, h: 2,
	fill: { color: '4472C4' },
	objectName: 'TestShape'
})

// Slide 2: Same shape on the right - should morph!
const slide2 = pptx.addSlide()
slide2.transition = { type: 'morph', durationMs: 2000 }
slide2.addText('Slide 2 - Shape on RIGHT (should have morphed)', { x: 0.5, y: 0.3, w: 9, h: 0.5, fontSize: 24 })
slide2.addShape('rect', {
	x: 6, y: 2, w: 2, h: 2,
	fill: { color: '4472C4' },
	objectName: 'TestShape'
})

pptx.writeFile({ fileName: 'test-morph-simple.pptx' })
	.then(() => console.log('Created: test-morph-simple.pptx'))
	.catch(err => console.error('Error:', err))
