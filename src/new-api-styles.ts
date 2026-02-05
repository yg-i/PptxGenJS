/**
 * Chalk-like chainable text styling for the new API
 *
 * Usage:
 *   bold.cyan("text")
 *   cyan.bold("text")
 *   bold.italic.red("text")
 *   style({ color: 'FF0000', bold: true })("text")
 */

// ============================================================================
// TYPES
// ============================================================================

export interface TextStyle {
	bold?: boolean
	italic?: boolean
	underline?: boolean
	strike?: boolean
	color?: string
	fontSize?: number
	fontFace?: string
}

export interface StyledFragment {
	text: string
	style: TextStyle
}

export type StyleBuilder = {
	(text: string): StyledFragment
} & ChainableStyles

interface ChainableStyles {
	// Font styles
	bold: StyleBuilder
	italic: StyleBuilder
	underline: StyleBuilder
	strike: StyleBuilder

	// Common colors
	white: StyleBuilder
	black: StyleBuilder
	gray: StyleBuilder
	red: StyleBuilder
	orange: StyleBuilder
	yellow: StyleBuilder
	green: StyleBuilder
	cyan: StyleBuilder
	blue: StyleBuilder
	purple: StyleBuilder
	magenta: StyleBuilder
	pink: StyleBuilder

	// Custom color
	color: (hex: string) => StyleBuilder

	// Font size
	size: (pt: number) => StyleBuilder

	// Font face
	font: (name: string) => StyleBuilder
}

// ============================================================================
// COLOR PALETTE
// ============================================================================

const COLORS: Record<string, string> = {
	white: 'FFFFFF',
	black: '000000',
	gray: 'B0B8C4',
	red: 'E74C3C',
	orange: 'E67E22',
	yellow: 'F1C40F',
	green: '2ECC71',
	cyan: '4FC3F7',
	blue: '3498DB',
	purple: '9B59B6',
	magenta: 'E91E63',
	pink: 'FF69B4',
}

// ============================================================================
// STYLE BUILDER IMPLEMENTATION
// ============================================================================

function createStyleBuilder(baseStyle: TextStyle = {}): StyleBuilder {
	// The callable function that applies styles to text
	const applyStyle = (text: string): StyledFragment => ({
		text,
		style: { ...baseStyle },
	})

	// Proxy to handle property access for chaining
	return new Proxy(applyStyle as StyleBuilder, {
		get(target, prop: string) {
			// Font style properties
			if (prop === 'bold') return createStyleBuilder({ ...baseStyle, bold: true })
			if (prop === 'italic') return createStyleBuilder({ ...baseStyle, italic: true })
			if (prop === 'underline') return createStyleBuilder({ ...baseStyle, underline: true })
			if (prop === 'strike') return createStyleBuilder({ ...baseStyle, strike: true })

			// Named colors
			if (COLORS[prop]) {
				return createStyleBuilder({ ...baseStyle, color: COLORS[prop] })
			}

			// Custom color function
			if (prop === 'color') {
				return (hex: string) => createStyleBuilder({ ...baseStyle, color: hex.replace('#', '') })
			}

			// Font size function
			if (prop === 'size') {
				return (pt: number) => createStyleBuilder({ ...baseStyle, fontSize: pt })
			}

			// Font face function
			if (prop === 'font') {
				return (name: string) => createStyleBuilder({ ...baseStyle, fontFace: name })
			}

			// Default: return the original property
			// eslint-disable-next-line @typescript-eslint/no-explicit-any
			return (target as any)[prop]
		},
	})
}

// ============================================================================
// EXPORTS - Pre-built style builders
// ============================================================================

/** Base style function - start a chain or apply custom styles */
export const style = (customStyle: TextStyle): StyleBuilder => createStyleBuilder(customStyle)

// Font styles
export const bold: StyleBuilder = createStyleBuilder({ bold: true })
export const italic: StyleBuilder = createStyleBuilder({ italic: true })
export const underline: StyleBuilder = createStyleBuilder({ underline: true })
export const strike: StyleBuilder = createStyleBuilder({ strike: true })

// Colors - can be chained with font styles
export const white: StyleBuilder = createStyleBuilder({ color: COLORS.white })
export const black: StyleBuilder = createStyleBuilder({ color: COLORS.black })
export const gray: StyleBuilder = createStyleBuilder({ color: COLORS.gray })
export const red: StyleBuilder = createStyleBuilder({ color: COLORS.red })
export const orange: StyleBuilder = createStyleBuilder({ color: COLORS.orange })
export const yellow: StyleBuilder = createStyleBuilder({ color: COLORS.yellow })
export const green: StyleBuilder = createStyleBuilder({ color: COLORS.green })
export const cyan: StyleBuilder = createStyleBuilder({ color: COLORS.cyan })
export const blue: StyleBuilder = createStyleBuilder({ color: COLORS.blue })
export const purple: StyleBuilder = createStyleBuilder({ color: COLORS.purple })
export const magenta: StyleBuilder = createStyleBuilder({ color: COLORS.magenta })
export const pink: StyleBuilder = createStyleBuilder({ color: COLORS.pink })

/** Create a custom color style builder */
export const color = (hex: string): StyleBuilder => createStyleBuilder({ color: hex.replace('#', '') })

// ============================================================================
// TAGGED TEMPLATE HELPER
// ============================================================================

/**
 * Process a tagged template literal with styled fragments
 */
export function processStyledTemplate(
	strings: TemplateStringsArray,
	...values: (string | StyledFragment)[]
): Array<string | StyledFragment> {
	const result: Array<string | StyledFragment> = []

	for (let i = 0; i < strings.length; i++) {
		// Add the string part (if non-empty)
		if (strings[i]) {
			result.push(strings[i])
		}

		// Add the interpolated value (if exists)
		if (i < values.length) {
			const value = values[i]
			if (typeof value === 'string') {
				result.push(value)
			} else {
				result.push(value) // StyledFragment
			}
		}
	}

	return result
}
