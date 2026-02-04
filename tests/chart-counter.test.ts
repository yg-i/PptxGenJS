/**
 * Tests for chart counter isolation
 * Verifies that multiple PptxGenJS instances have independent chart counters
 */
import { describe, it } from 'node:test'
import assert from 'node:assert/strict'

import PptxGenJS from '../src/pptxgen'

describe('chart counter isolation', () => {
	it('initializes _chartCounter to 0 on new presentation', () => {
		const pptx = new PptxGenJS()
		assert.equal(pptx.presLayout._chartCounter, 0)
	})

	it('maintains independent counters for different presentations', () => {
		const pptx1 = new PptxGenJS()
		const pptx2 = new PptxGenJS()

		// Verify both start at 0
		assert.equal(pptx1.presLayout._chartCounter, 0)
		assert.equal(pptx2.presLayout._chartCounter, 0)

		// Simulate incrementing counter on first presentation
		pptx1.presLayout._chartCounter++
		pptx1.presLayout._chartCounter++

		// Verify counters are independent
		assert.equal(pptx1.presLayout._chartCounter, 2)
		assert.equal(pptx2.presLayout._chartCounter, 0, 'Second presentation counter should be unaffected')

		// Increment second presentation
		pptx2.presLayout._chartCounter++

		// Final verification
		assert.equal(pptx1.presLayout._chartCounter, 2)
		assert.equal(pptx2.presLayout._chartCounter, 1)
	})

	it('slides share the same chart counter with their presentation', () => {
		const pptx = new PptxGenJS()
		const slide1 = pptx.addSlide()
		const slide2 = pptx.addSlide()

		// All should share the same _presLayout object
		assert.equal(slide1._presLayout, pptx.presLayout)
		assert.equal(slide2._presLayout, pptx.presLayout)
		assert.equal(slide1._presLayout._chartCounter, slide2._presLayout._chartCounter)
	})
})
