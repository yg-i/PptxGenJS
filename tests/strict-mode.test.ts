/**
 * Tests for strict mode functionality
 */
import { describe, it } from 'node:test'
import assert from 'node:assert/strict'

import PptxGenJS from '../src/pptxgen'

describe('strictMode', () => {
	it('defaults to false', () => {
		const pptx = new PptxGenJS()
		assert.equal(pptx.strictMode, false)
	})

	it('can be set to true', () => {
		const pptx = new PptxGenJS()
		pptx.strictMode = true
		assert.equal(pptx.strictMode, true)
	})

	it('warns on invalid input when strictMode is false', () => {
		const pptx = new PptxGenJS()
		pptx.strictMode = false

		// defineLayout with invalid height type - warns but doesn't crash
		let errorThrown = false
		try {
			// @ts-expect-error - intentionally passing invalid type
			pptx.defineLayout({ name: 'Test', width: 10, height: 'invalid' })
		} catch {
			errorThrown = true
		}
		assert.equal(errorThrown, false, 'Should not throw when strictMode is false')
	})

	it('throws on invalid input when strictMode is true', () => {
		const pptx = new PptxGenJS()
		pptx.strictMode = true

		// Should throw
		assert.throws(
			() => {
				// @ts-expect-error - intentionally passing invalid input
				pptx.addSection(null)
			},
			{ message: /addSection requires an argument/ }
		)
	})

	it('throws on missing section title when strictMode is true', () => {
		const pptx = new PptxGenJS()
		pptx.strictMode = true

		assert.throws(
			() => {
				// @ts-expect-error - intentionally passing incomplete input
				pptx.addSection({})
			},
			{ message: /addSection requires a title/ }
		)
	})

	it('throws on missing defineLayout params when strictMode is true', () => {
		const pptx = new PptxGenJS()
		pptx.strictMode = true

		assert.throws(
			() => {
				// @ts-expect-error - intentionally passing invalid input
				pptx.defineLayout(null)
			},
			{ message: /defineLayout requires/ }
		)
	})
})
