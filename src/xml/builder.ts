/**
 * PptxGenJS: XML Builder
 * Provides a fluent API for building XML strings, replacing string concatenation
 * @since v4.2.0
 */

/**
 * Attribute value type - string, number, or boolean
 */
type AttributeValue = string | number | boolean | undefined | null

/**
 * Attributes object for XML elements
 */
type Attributes = Record<string, AttributeValue>

/**
 * Fluent XML Builder for constructing XML strings
 *
 * @example Basic usage
 * ```ts
 * const xml = new XmlBuilder()
 *   .tag('a:p')
 *     .tag('a:r')
 *       .tag('a:t').text('Hello World').close()
 *     .close()
 *   .close()
 *   .build()
 * // Result: <a:p><a:r><a:t>Hello World</a:t></a:r></a:p>
 * ```
 *
 * @example With attributes
 * ```ts
 * const xml = new XmlBuilder()
 *   .tag('a:sp', { id: 1, name: 'Shape 1' })
 *     .selfClosingTag('a:off', { x: 0, y: 0 })
 *   .close()
 *   .build()
 * // Result: <a:sp id="1" name="Shape 1"><a:off x="0" y="0"/></a:sp>
 * ```
 *
 * @example Conditional content
 * ```ts
 * const xml = new XmlBuilder()
 *   .tag('a:p')
 *     .when(hasBold, b => b.tag('a:rPr', { b: 1 }).close())
 *     .tag('a:t').text('Content').close()
 *   .close()
 *   .build()
 * ```
 */
export class XmlBuilder {
	private parts: string[] = []
	private openTags: string[] = []

	/**
	 * Opens an XML tag with optional attributes
	 * @param tagName - The tag name (e.g., 'a:p', 'p:sp')
	 * @param attrs - Optional attributes object
	 * @returns this builder for chaining
	 */
	tag(tagName: string, attrs?: Attributes): this {
		this.openTags.push(tagName)
		this.parts.push(`<${tagName}${this.formatAttributes(attrs)}>`)
		return this
	}

	/**
	 * Adds a self-closing tag with optional attributes
	 * @param tagName - The tag name
	 * @param attrs - Optional attributes object
	 * @returns this builder for chaining
	 */
	selfClosingTag(tagName: string, attrs?: Attributes): this {
		this.parts.push(`<${tagName}${this.formatAttributes(attrs)}/>`)
		return this
	}

	/**
	 * Closes the most recently opened tag
	 * @returns this builder for chaining
	 * @throws Error if no tag is open
	 */
	close(): this {
		const tagName = this.openTags.pop()
		if (!tagName) {
			throw new Error('XmlBuilder: No open tag to close')
		}
		this.parts.push(`</${tagName}>`)
		return this
	}

	/**
	 * Closes multiple tags at once
	 * @param count - Number of tags to close
	 * @returns this builder for chaining
	 */
	closeMultiple(count: number): this {
		for (let i = 0; i < count; i++) {
			this.close()
		}
		return this
	}

	/**
	 * Closes all open tags
	 * @returns this builder for chaining
	 */
	closeAll(): this {
		while (this.openTags.length > 0) {
			this.close()
		}
		return this
	}

	/**
	 * Adds text content (XML-escaped)
	 * @param content - Text content to add
	 * @returns this builder for chaining
	 */
	text(content: string | number): this {
		this.parts.push(this.escapeXml(String(content)))
		return this
	}

	/**
	 * Adds raw XML content without escaping
	 * @param xml - Raw XML string to add
	 * @returns this builder for chaining
	 */
	raw(xml: string): this {
		this.parts.push(xml)
		return this
	}

	/**
	 * Conditionally adds content to the builder
	 * @param condition - Boolean condition
	 * @param builder - Function that receives the builder and adds content
	 * @returns this builder for chaining
	 */
	when(condition: boolean, builder: (b: XmlBuilder) => void): this {
		if (condition) {
			builder(this)
		}
		return this
	}

	/**
	 * Iterates over an array and adds content for each item
	 * @param items - Array to iterate over
	 * @param builder - Function that receives the builder and current item
	 * @returns this builder for chaining
	 */
	forEach<T>(items: T[], builder: (b: XmlBuilder, item: T, index: number) => void): this {
		items.forEach((item, index) => {
			builder(this, item, index)
		})
		return this
	}

	/**
	 * Adds an element with text content in one call
	 * @param tagName - The tag name
	 * @param content - Text content
	 * @param attrs - Optional attributes
	 * @returns this builder for chaining
	 */
	element(tagName: string, content: string | number, attrs?: Attributes): this {
		this.parts.push(`<${tagName}${this.formatAttributes(attrs)}>${this.escapeXml(String(content))}</${tagName}>`)
		return this
	}

	/**
	 * Adds an attribute element (common in OOXML)
	 * Shorthand for selfClosingTag with a 'val' attribute
	 * @param tagName - The tag name
	 * @param value - The value for the 'val' attribute
	 * @returns this builder for chaining
	 */
	attrElement(tagName: string, value: AttributeValue): this {
		return this.selfClosingTag(tagName, { val: value })
	}

	/**
	 * Builds and returns the final XML string
	 * @param autoCloseAll - If true, automatically closes all open tags
	 * @returns The built XML string
	 */
	build(autoCloseAll = false): string {
		if (autoCloseAll) {
			this.closeAll()
		}
		if (this.openTags.length > 0) {
			console.warn(`XmlBuilder: ${this.openTags.length} unclosed tag(s): ${this.openTags.join(', ')}`)
		}
		return this.parts.join('')
	}

	/**
	 * Returns the current XML string without closing tags (for inspection)
	 * @returns Current XML string
	 */
	toString(): string {
		return this.parts.join('')
	}

	/**
	 * Resets the builder for reuse
	 * @returns this builder for chaining
	 */
	reset(): this {
		this.parts = []
		this.openTags = []
		return this
	}

	/**
	 * Returns the number of currently open tags
	 */
	get depth(): number {
		return this.openTags.length
	}

	/**
	 * Formats attributes object into XML attribute string
	 */
	private formatAttributes(attrs?: Attributes): string {
		if (!attrs) return ''

		const attrParts: string[] = []
		for (const [key, value] of Object.entries(attrs)) {
			// Skip undefined, null, and false boolean values
			if (value === undefined || value === null || value === false) continue
			// Boolean true becomes just the attribute name with "1" value (OOXML convention)
			if (value === true) {
				attrParts.push(`${key}="1"`)
			} else {
				attrParts.push(`${key}="${this.escapeXml(String(value))}"`)
			}
		}

		return attrParts.length > 0 ? ' ' + attrParts.join(' ') : ''
	}

	/**
	 * Escapes special XML characters
	 */
	private escapeXml(str: string): string {
		return str
			.replace(/&/g, '&amp;')
			.replace(/</g, '&lt;')
			.replace(/>/g, '&gt;')
			.replace(/"/g, '&quot;')
			.replace(/'/g, '&apos;')
	}
}

/**
 * Creates a new XmlBuilder instance
 * @returns New XmlBuilder
 */
export function xml(): XmlBuilder {
	return new XmlBuilder()
}

/**
 * Common OOXML element helpers
 */
export const OoxmlElements = {
	/**
	 * Creates <a:off x="..." y="..."/>
	 */
	offset(x: number, y: number): string {
		return `<a:off x="${x}" y="${y}"/>`
	},

	/**
	 * Creates <a:ext cx="..." cy="..."/>
	 */
	extent(cx: number, cy: number): string {
		return `<a:ext cx="${cx}" cy="${cy}"/>`
	},

	/**
	 * Creates position/size transform: <a:xfrm>...</a:xfrm>
	 */
	transform(x: number, y: number, cx: number, cy: number, attrs?: Attributes): string {
		const builder = new XmlBuilder()
		builder.tag('a:xfrm', attrs)
			.raw(OoxmlElements.offset(x, y))
			.raw(OoxmlElements.extent(cx, cy))
			.close()
		return builder.build()
	},

	/**
	 * Creates solid fill element: <a:solidFill>...</a:solidFill>
	 */
	solidFill(colorElement: string): string {
		return `<a:solidFill>${colorElement}</a:solidFill>`
	},

	/**
	 * Creates no fill element: <a:noFill/>
	 */
	noFill(): string {
		return '<a:noFill/>'
	},
}
