/**
 * Tests for deprecation warning functionality
 * Verifies that deprecated properties emit console warnings
 */
import { describe, it, beforeEach, mock } from 'node:test'
import assert from 'node:assert/strict'

import PptxGenJS from '../src/pptxgen'
import { DEPRECATED_PROPERTY_MAP, warnDeprecatedProperty } from '../src/gen-utils'

describe('deprecation warnings', () => {
	// Store original console.warn
	let originalWarn: typeof console.warn
	let warnCalls: string[]

	beforeEach(() => {
		originalWarn = console.warn
		warnCalls = []
		// Mock console.warn to capture calls
		console.warn = (...args: unknown[]) => {
			warnCalls.push(args.map(a => String(a)).join(' '))
		}
	})

	it('DEPRECATED_PROPERTY_MAP contains expected mappings', () => {
		assert.equal(DEPRECATED_PROPERTY_MAP.alpha, 'transparency')
		assert.equal(DEPRECATED_PROPERTY_MAP.bkgd, 'background')
		assert.equal(DEPRECATED_PROPERTY_MAP.autoFit, 'fit')
		assert.equal(DEPRECATED_PROPERTY_MAP.lineSize, 'line.width')
		assert.equal(DEPRECATED_PROPERTY_MAP.newSlideStartY, 'autoPageSlideStartY')
	})

	it('warnDeprecatedProperty emits warning with replacement suggestion', () => {
		// Reset the warning tracker by using a unique context
		warnDeprecatedProperty('alpha', 'test-unique-context-1')

		assert.equal(warnCalls.length, 1)
		assert.ok(warnCalls[0].includes('alpha'))
		assert.ok(warnCalls[0].includes('DEPRECATION WARNING'))
		assert.ok(warnCalls[0].includes('transparency'))

		// Restore
		console.warn = originalWarn
	})

	it('warnDeprecatedProperty only warns once per property per context', () => {
		// Use unique context to avoid interference from other tests
		const ctx = 'test-unique-context-2'
		warnDeprecatedProperty('bkgd', ctx)
		warnDeprecatedProperty('bkgd', ctx)
		warnDeprecatedProperty('bkgd', ctx)

		// Should only have one warning despite multiple calls
		const bkgdWarnings = warnCalls.filter(w => w.includes('bkgd') && w.includes(ctx))
		assert.equal(bkgdWarnings.length, 1)

		// Restore
		console.warn = originalWarn
	})

	it('addShape warns on deprecated properties', () => {
		const pptx = new PptxGenJS()
		const slide = pptx.addSlide()

		// Use deprecated lineSize property
		// @ts-expect-error - intentionally using deprecated property
		slide.addShape('rect', { x: 1, y: 1, w: 2, h: 2, lineSize: 3 })

		const lineSizeWarning = warnCalls.find(w => w.includes('lineSize'))
		assert.ok(lineSizeWarning, 'Should warn about lineSize')
		assert.ok(lineSizeWarning.includes('line.width'), 'Should suggest replacement')

		// Restore
		console.warn = originalWarn
	})

	it('addText warns on deprecated properties', () => {
		const pptx = new PptxGenJS()
		const slide = pptx.addSlide()

		// Use deprecated autoFit property
		// @ts-expect-error - intentionally using deprecated property
		slide.addText('Hello', { x: 1, y: 1, w: 2, h: 1, autoFit: true })

		const autoFitWarning = warnCalls.find(w => w.includes('autoFit'))
		assert.ok(autoFitWarning, 'Should warn about autoFit')
		assert.ok(autoFitWarning.includes('fit'), 'Should suggest replacement')

		// Restore
		console.warn = originalWarn
	})

	it('addTable warns on deprecated properties', () => {
		const pptx = new PptxGenJS()
		const slide = pptx.addSlide()

		// Use deprecated newSlideStartY property
		// @ts-expect-error - intentionally using deprecated property
		slide.addTable([['A', 'B']], { x: 1, y: 1, w: 8, newSlideStartY: 1 })

		const newSlideStartYWarning = warnCalls.find(w => w.includes('newSlideStartY'))
		assert.ok(newSlideStartYWarning, 'Should warn about newSlideStartY')
		assert.ok(newSlideStartYWarning.includes('autoPageSlideStartY'), 'Should suggest replacement')

		// Restore
		console.warn = originalWarn
	})
})
