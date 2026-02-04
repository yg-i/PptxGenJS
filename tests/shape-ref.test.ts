/**
 * Tests for ShapeRef-based animation API
 * Verifies that addShape/addText/addImage return ShapeRef for animation targeting
 */
import { describe, it } from 'node:test'
import assert from 'node:assert/strict'

import PptxGenJS from '../src/pptxgen'

describe('ShapeRef API', () => {
	it('addShape returns a ShapeRef with correct index', () => {
		const pptx = new PptxGenJS()
		const slide = pptx.addSlide()

		const shape1 = slide.addShape('rect', { x: 1, y: 1, w: 2, h: 2 })
		assert.equal(shape1._shapeIndex, 0, 'First shape should have index 0')

		const shape2 = slide.addShape('ellipse', { x: 3, y: 1, w: 2, h: 2 })
		assert.equal(shape2._shapeIndex, 1, 'Second shape should have index 1')
	})

	it('addText returns a ShapeRef with correct index', () => {
		const pptx = new PptxGenJS()
		const slide = pptx.addSlide()

		const text1 = slide.addText('Hello', { x: 1, y: 1, w: 2, h: 1 })
		assert.equal(text1._shapeIndex, 0, 'First text should have index 0')

		const text2 = slide.addText('World', { x: 1, y: 2, w: 2, h: 1 })
		assert.equal(text2._shapeIndex, 1, 'Second text should have index 1')
	})

	it('addImage returns a ShapeRef with correct index', () => {
		const pptx = new PptxGenJS()
		const slide = pptx.addSlide()

		// Use a base64 encoded 1x1 transparent PNG
		const transparentPng = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII='

		const image1 = slide.addImage({ data: transparentPng, x: 1, y: 1, w: 2, h: 2 })
		assert.equal(image1._shapeIndex, 0, 'First image should have index 0')

		const image2 = slide.addImage({ data: transparentPng, x: 3, y: 1, w: 2, h: 2 })
		assert.equal(image2._shapeIndex, 1, 'Second image should have index 1')
	})

	it('addAnimation accepts ShapeRef', () => {
		const pptx = new PptxGenJS()
		const slide = pptx.addSlide()

		const shape = slide.addShape('rect', { x: 1, y: 1, w: 2, h: 2 })
		slide.addAnimation(shape, { type: 'fade' })

		assert.equal(slide._animations.length, 1, 'Animation should be added')
		assert.equal(slide._animations[0].shapeIndex, 0, 'Animation should target shape index 0')
	})

	it('addAnimation still accepts numeric index for backward compatibility', () => {
		const pptx = new PptxGenJS()
		const slide = pptx.addSlide()

		slide.addShape('rect', { x: 1, y: 1, w: 2, h: 2 })
		slide.addAnimation(0, { type: 'fade' })

		assert.equal(slide._animations.length, 1, 'Animation should be added')
		assert.equal(slide._animations[0].shapeIndex, 0, 'Animation should target shape index 0')
	})

	it('addAnimation with ShapeRef works correctly with multiple shapes', () => {
		const pptx = new PptxGenJS()
		const slide = pptx.addSlide()

		const rect = slide.addShape('rect', { x: 1, y: 1, w: 2, h: 2 })
		const text = slide.addText('Hello', { x: 1, y: 3, w: 2, h: 1 })
		const ellipse = slide.addShape('ellipse', { x: 3, y: 1, w: 2, h: 2 })

		// Animate in non-sequential order
		slide.addAnimation(ellipse, { type: 'fly-in', direction: 'from-bottom' })
		slide.addAnimation(rect, { type: 'fade' })
		slide.addAnimation(text, { type: 'appear' })

		assert.equal(slide._animations.length, 3, 'All animations should be added')
		assert.equal(slide._animations[0].shapeIndex, 2, 'First animation targets ellipse (index 2)')
		assert.equal(slide._animations[1].shapeIndex, 0, 'Second animation targets rect (index 0)')
		assert.equal(slide._animations[2].shapeIndex, 1, 'Third animation targets text (index 1)')
	})

	it('ShapeRef stores reference to parent slide', () => {
		const pptx = new PptxGenJS()
		const slide = pptx.addSlide()

		const shape = slide.addShape('rect', { x: 1, y: 1, w: 2, h: 2 })

		// ShapeRef should have a reference back to the slide
		assert.ok(shape._slideRef, 'ShapeRef should have _slideRef')
	})

	it('addAnimation warns when ShapeRef belongs to different slide', () => {
		const pptx = new PptxGenJS()
		const slide1 = pptx.addSlide()
		const slide2 = pptx.addSlide()

		const shapeOnSlide1 = slide1.addShape('rect', { x: 1, y: 1, w: 2, h: 2 })

		// Try to use shapeOnSlide1's ShapeRef with slide2 - should warn but not crash
		slide2.addShape('ellipse', { x: 1, y: 1, w: 2, h: 2 }) // Add a shape to slide2
		slide2.addAnimation(shapeOnSlide1, { type: 'fade' }) // This should warn

		// Animation should not be added since the ShapeRef belongs to a different slide
		assert.equal(slide2._animations.length, 0, 'Animation should not be added for wrong slide')
	})
})
