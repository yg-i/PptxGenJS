/**
 * Tests for gen-utils.ts utility functions
 */
import { describe, it } from 'node:test'
import assert from 'node:assert/strict'

import {
	inch2Emu,
	valToPts,
	encodeXmlEntities,
	getSmartParseNumber,
	convertRotationDegrees,
	componentToHex,
	rgbToHex,
	createColorElement,
} from '../src/gen-utils'
import { EMU, ONEPT } from '../src/core-enums'
import type { PresLayout } from '../src/core-interfaces'

describe('inch2Emu', () => {
	it('converts inches to EMU correctly', () => {
		assert.equal(inch2Emu(1), EMU)
		assert.equal(inch2Emu(2), 2 * EMU)
		assert.equal(inch2Emu(0.5), Math.round(0.5 * EMU))
	})

	it('handles string input', () => {
		assert.equal(inch2Emu('1'), EMU)
		assert.equal(inch2Emu('2in'), 2 * EMU)
		assert.equal(inch2Emu('0.5'), Math.round(0.5 * EMU))
	})

	it('returns value unchanged if already large (>100, assumed EMU)', () => {
		assert.equal(inch2Emu(914400), 914400)
		assert.equal(inch2Emu(1000000), 1000000)
	})
})

describe('valToPts', () => {
	it('converts number to points using ONEPT', () => {
		assert.equal(valToPts(1), ONEPT)
		assert.equal(valToPts(2), 2 * ONEPT)
		assert.equal(valToPts(12), 12 * ONEPT)
	})

	it('handles string input', () => {
		assert.equal(valToPts('1'), ONEPT)
		assert.equal(valToPts('10'), 10 * ONEPT)
	})

	it('returns 0 for invalid input', () => {
		assert.equal(valToPts('abc'), 0)
		assert.equal(valToPts(NaN), 0)
	})
})

describe('encodeXmlEntities', () => {
	it('encodes special XML characters', () => {
		assert.equal(encodeXmlEntities('&'), '&amp;')
		assert.equal(encodeXmlEntities('<'), '&lt;')
		assert.equal(encodeXmlEntities('>'), '&gt;')
		assert.equal(encodeXmlEntities('"'), '&quot;')
		assert.equal(encodeXmlEntities("'"), '&apos;')
	})

	it('encodes all special chars in a string', () => {
		assert.equal(
			encodeXmlEntities('<div class="test">Hello & goodbye</div>'),
			'&lt;div class=&quot;test&quot;&gt;Hello &amp; goodbye&lt;/div&gt;'
		)
	})

	it('returns empty string for undefined/null', () => {
		assert.equal(encodeXmlEntities(undefined as unknown as string), '')
		assert.equal(encodeXmlEntities(null as unknown as string), '')
	})

	it('handles empty string', () => {
		assert.equal(encodeXmlEntities(''), '')
	})

	it('handles strings without special chars', () => {
		assert.equal(encodeXmlEntities('Hello World'), 'Hello World')
	})
})

describe('convertRotationDegrees', () => {
	it('converts degrees to PowerPoint rotation value', () => {
		assert.equal(convertRotationDegrees(0), 0)
		assert.equal(convertRotationDegrees(90), 90 * 60000)
		assert.equal(convertRotationDegrees(180), 180 * 60000)
		assert.equal(convertRotationDegrees(270), 270 * 60000)
		assert.equal(convertRotationDegrees(360), 360 * 60000)
	})

	it('handles values over 360 by subtracting 360', () => {
		assert.equal(convertRotationDegrees(450), 90 * 60000) // 450 - 360 = 90
	})

	it('handles undefined/null as 0', () => {
		assert.equal(convertRotationDegrees(undefined as unknown as number), 0)
		assert.equal(convertRotationDegrees(null as unknown as number), 0)
	})
})

describe('componentToHex', () => {
	it('converts component values to hex strings', () => {
		assert.equal(componentToHex(0), '00')
		assert.equal(componentToHex(255), 'ff')
		assert.equal(componentToHex(128), '80')
		assert.equal(componentToHex(15), '0f')
		assert.equal(componentToHex(16), '10')
	})
})

describe('rgbToHex', () => {
	it('converts RGB values to uppercase hex string', () => {
		assert.equal(rgbToHex(255, 0, 0), 'FF0000') // red
		assert.equal(rgbToHex(0, 255, 0), '00FF00') // green
		assert.equal(rgbToHex(0, 0, 255), '0000FF') // blue
		assert.equal(rgbToHex(255, 255, 255), 'FFFFFF') // white
		assert.equal(rgbToHex(0, 0, 0), '000000') // black
		assert.equal(rgbToHex(128, 128, 128), '808080') // gray
	})
})

describe('getSmartParseNumber', () => {
	const mockLayout: PresLayout = {
		name: 'test',
		width: 9144000, // 10 inches in EMU
		height: 5143500, // ~5.625 inches in EMU
	}

	it('converts inches (small numbers < 100) to EMU', () => {
		assert.equal(getSmartParseNumber(1, 'X', mockLayout), EMU)
		assert.equal(getSmartParseNumber(5, 'Y', mockLayout), 5 * EMU)
	})

	it('returns large numbers unchanged (assumed EMU)', () => {
		assert.equal(getSmartParseNumber(914400, 'X', mockLayout), 914400)
		assert.equal(getSmartParseNumber(1000000, 'Y', mockLayout), 1000000)
	})

	it('handles percentage strings for X direction', () => {
		assert.equal(getSmartParseNumber('50%', 'X', mockLayout), Math.round(0.5 * mockLayout.width))
		assert.equal(getSmartParseNumber('100%', 'X', mockLayout), mockLayout.width)
	})

	it('handles percentage strings for Y direction', () => {
		assert.equal(getSmartParseNumber('50%', 'Y', mockLayout), Math.round(0.5 * mockLayout.height))
		assert.equal(getSmartParseNumber('100%', 'Y', mockLayout), mockLayout.height)
	})

	it('handles string numeric values', () => {
		assert.equal(getSmartParseNumber('5', 'X', mockLayout), 5 * EMU)
	})

	it('returns 0 for invalid input', () => {
		assert.equal(getSmartParseNumber(undefined as unknown as number, 'X', mockLayout), 0)
	})
})

describe('createColorElement', () => {
	it('creates srgbClr element for hex colors', () => {
		const result = createColorElement('FF0000')
		assert.equal(result, '<a:srgbClr val="FF0000"/>')
	})

	it('creates schemeClr element for theme colors', () => {
		const result = createColorElement('tx1')
		assert.equal(result, '<a:schemeClr val="tx1"/>')
	})

	it('strips # prefix from hex colors', () => {
		const result = createColorElement('#FF0000')
		assert.equal(result, '<a:srgbClr val="FF0000"/>')
	})

	it('adds inner elements when provided', () => {
		const result = createColorElement('FF0000', '<a:alpha val="50000"/>')
		assert.equal(result, '<a:srgbClr val="FF0000"><a:alpha val="50000"/></a:srgbClr>')
	})

	it('uppercases hex values', () => {
		const result = createColorElement('ff0000')
		assert.equal(result, '<a:srgbClr val="FF0000"/>')
	})
})
