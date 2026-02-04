/**
 * Tests for XmlBuilder utility class
 */
import { describe, it } from 'node:test'
import assert from 'node:assert/strict'

import { XmlBuilder, xml, OoxmlElements } from '../src/xml-builder'

describe('XmlBuilder', () => {
	describe('basic tag operations', () => {
		it('creates simple self-closing tag', () => {
			const result = new XmlBuilder()
				.selfClosingTag('a:off', { x: 0, y: 0 })
				.build()
			assert.equal(result, '<a:off x="0" y="0"/>')
		})

		it('creates tag with content', () => {
			const result = new XmlBuilder()
				.tag('a:t')
				.text('Hello World')
				.close()
				.build()
			assert.equal(result, '<a:t>Hello World</a:t>')
		})

		it('creates nested tags', () => {
			const result = new XmlBuilder()
				.tag('a:p')
					.tag('a:r')
						.tag('a:t').text('Content').close()
					.close()
				.close()
				.build()
			assert.equal(result, '<a:p><a:r><a:t>Content</a:t></a:r></a:p>')
		})

		it('handles attributes with various types', () => {
			const result = new XmlBuilder()
				.selfClosingTag('test', {
					str: 'value',
					num: 42,
					bool: true,
					empty: undefined,
					nullVal: null,
					falseVal: false,
				})
				.build()
			assert.equal(result, '<test str="value" num="42" bool="1"/>')
		})
	})

	describe('element helper', () => {
		it('creates element with text content', () => {
			const result = new XmlBuilder()
				.element('a:t', 'Hello', { lang: 'en-US' })
				.build()
			assert.equal(result, '<a:t lang="en-US">Hello</a:t>')
		})

		it('handles numeric content', () => {
			const result = new XmlBuilder()
				.element('val', 42)
				.build()
			assert.equal(result, '<val>42</val>')
		})
	})

	describe('attrElement helper', () => {
		it('creates attribute element with val', () => {
			const result = new XmlBuilder()
				.attrElement('a:alpha', 50000)
				.build()
			assert.equal(result, '<a:alpha val="50000"/>')
		})
	})

	describe('XML escaping', () => {
		it('escapes special XML characters in text', () => {
			const result = new XmlBuilder()
				.tag('a:t')
				.text('<script>alert("XSS")</script>')
				.close()
				.build()
			assert.equal(result, '<a:t>&lt;script&gt;alert(&quot;XSS&quot;)&lt;/script&gt;</a:t>')
		})

		it('escapes special XML characters in attributes', () => {
			const result = new XmlBuilder()
				.selfClosingTag('test', { name: 'Value with "quotes" & <brackets>' })
				.build()
			assert.equal(result, '<test name="Value with &quot;quotes&quot; &amp; &lt;brackets&gt;"/>')
		})
	})

	describe('raw content', () => {
		it('adds raw XML without escaping', () => {
			const result = new XmlBuilder()
				.tag('p:sp')
				.raw('<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>')
				.close()
				.build()
			assert.equal(result, '<p:sp><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></p:sp>')
		})
	})

	describe('conditional building', () => {
		it('when() adds content when condition is true', () => {
			const result = new XmlBuilder()
				.tag('a:rPr')
				.when(true, b => b.selfClosingTag('a:b'))
				.close()
				.build()
			assert.equal(result, '<a:rPr><a:b/></a:rPr>')
		})

		it('when() skips content when condition is false', () => {
			const result = new XmlBuilder()
				.tag('a:rPr')
				.when(false, b => b.selfClosingTag('a:b'))
				.close()
				.build()
			assert.equal(result, '<a:rPr></a:rPr>')
		})
	})

	describe('forEach iteration', () => {
		it('iterates over array items', () => {
			const items = ['A', 'B', 'C']
			const result = new XmlBuilder()
				.tag('list')
				.forEach(items, (b, item) => {
					b.element('item', item)
				})
				.close()
				.build()
			assert.equal(result, '<list><item>A</item><item>B</item><item>C</item></list>')
		})

		it('provides index to callback', () => {
			const items = ['X', 'Y']
			const result = new XmlBuilder()
				.tag('list')
				.forEach(items, (b, item, idx) => {
					b.selfClosingTag('item', { id: idx, value: item })
				})
				.close()
				.build()
			assert.equal(result, '<list><item id="0" value="X"/><item id="1" value="Y"/></list>')
		})
	})

	describe('closeMultiple and closeAll', () => {
		it('closeMultiple closes specified number of tags', () => {
			const result = new XmlBuilder()
				.tag('a').tag('b').tag('c')
				.text('X')
				.closeMultiple(3)
				.build()
			assert.equal(result, '<a><b><c>X</c></b></a>')
		})

		it('closeAll closes all open tags', () => {
			const result = new XmlBuilder()
				.tag('a').tag('b').tag('c')
				.text('X')
				.closeAll()
				.build()
			assert.equal(result, '<a><b><c>X</c></b></a>')
		})

		it('build with autoCloseAll closes remaining tags', () => {
			const result = new XmlBuilder()
				.tag('a').tag('b').tag('c')
				.text('X')
				.build(true)
			assert.equal(result, '<a><b><c>X</c></b></a>')
		})
	})

	describe('depth property', () => {
		it('tracks nesting depth', () => {
			const builder = new XmlBuilder()
			assert.equal(builder.depth, 0)

			builder.tag('a')
			assert.equal(builder.depth, 1)

			builder.tag('b')
			assert.equal(builder.depth, 2)

			builder.close()
			assert.equal(builder.depth, 1)
		})
	})

	describe('reset', () => {
		it('clears builder state for reuse', () => {
			const builder = new XmlBuilder()
			builder.tag('a').text('first').close()
			const first = builder.build()

			builder.reset()
			builder.tag('b').text('second').close()
			const second = builder.build()

			assert.equal(first, '<a>first</a>')
			assert.equal(second, '<b>second</b>')
		})
	})
})

describe('xml() factory function', () => {
	it('creates new XmlBuilder instance', () => {
		const result = xml()
			.tag('test')
			.close()
			.build()
		assert.equal(result, '<test></test>')
	})
})

describe('OoxmlElements helpers', () => {
	it('offset creates a:off element', () => {
		const result = OoxmlElements.offset(100, 200)
		assert.equal(result, '<a:off x="100" y="200"/>')
	})

	it('extent creates a:ext element', () => {
		const result = OoxmlElements.extent(500, 300)
		assert.equal(result, '<a:ext cx="500" cy="300"/>')
	})

	it('transform creates full xfrm element', () => {
		const result = OoxmlElements.transform(100, 200, 500, 300)
		assert.equal(result, '<a:xfrm><a:off x="100" y="200"/><a:ext cx="500" cy="300"/></a:xfrm>')
	})

	it('transform with attributes', () => {
		const result = OoxmlElements.transform(0, 0, 100, 100, { rot: 5400000 })
		assert.equal(result, '<a:xfrm rot="5400000"><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm>')
	})

	it('solidFill wraps color element', () => {
		const result = OoxmlElements.solidFill('<a:srgbClr val="FF0000"/>')
		assert.equal(result, '<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>')
	})

	it('noFill creates empty element', () => {
		assert.equal(OoxmlElements.noFill(), '<a:noFill/>')
	})
})
