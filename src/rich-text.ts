/**
 * Rich text utilities for ergonomic inline text styling.
 *
 * @example
 * ```typescript
 * import { textStyle } from 'pptxgenjs'
 *
 * const keyword = textStyle({ bold: true, color: '5DADE2' })
 * const emphasis = textStyle({ italic: true, color: 'FF6B6B' })
 *
 * slide.addRichText`By ${keyword('committing')} to a choice, the agent is ${keyword('justified')} in being in it.`
 * ```
 */

import type { TextPropsOptions, TextProps } from './core-interfaces'

/**
 * A styled text fragment returned by a style function.
 */
export interface StyledTextFragment {
	readonly __styledText: true
	readonly text: string
	readonly style: TextPropsOptions
}

/**
 * A style function that wraps text with predefined styles.
 */
export type StyleFunction = (text: string) => StyledTextFragment

/**
 * Creates a reusable style function for rich text.
 *
 * @param options - Text styling options (bold, italic, color, fontSize, etc.)
 * @returns A function that wraps text with the specified styles
 *
 * @example
 * ```typescript
 * const keyword = textStyle({ bold: true, color: '5DADE2' })
 * const warning = textStyle({ bold: true, color: 'FF0000' })
 *
 * // Use in tagged template
 * slide.addRichText`Click ${keyword('here')} to continue. ${warning('Warning')}: this is permanent.`
 * ```
 */
export function textStyle(options: TextPropsOptions): StyleFunction {
	return (text: string): StyledTextFragment => ({
		__styledText: true,
		text,
		style: options,
	})
}

/**
 * Type guard to check if a value is a StyledTextFragment.
 */
export function isStyledTextFragment(value: unknown): value is StyledTextFragment {
	return (
		typeof value === 'object' &&
		value !== null &&
		'__styledText' in value &&
		(value as StyledTextFragment).__styledText === true
	)
}

/**
 * Options for addRichText method.
 * Extends TextPropsOptions for positioning and default styling.
 */
export interface RichTextOptions extends TextPropsOptions {}

/**
 * Converts tagged template arguments to TextProps array.
 *
 * @param strings - Template literal strings
 * @param values - Interpolated values (strings or StyledTextFragments)
 * @param defaultOptions - Default text options for unstyled text
 * @returns Array of TextProps for addText
 */
export function convertRichTextToTextProps(
	strings: TemplateStringsArray | string[],
	values: (string | StyledTextFragment)[],
	defaultOptions: RichTextOptions = {}
): TextProps[] {
	const result: TextProps[] = []

	// Extract styling options (exclude positioning props)
	const { x, y, w, h, ...styleOptions } = defaultOptions

	// Clean up undefined values from style options
	const defaultStyle: TextPropsOptions = {}
	for (const [key, value] of Object.entries(styleOptions)) {
		if (value !== undefined) {
			(defaultStyle as Record<string, unknown>)[key] = value
		}
	}

	const hasDefaultStyle = Object.keys(defaultStyle).length > 0

	for (let i = 0; i < strings.length; i++) {
		// Add the string part (if not empty)
		const str = strings[i]
		if (str) {
			result.push({
				text: str,
				options: hasDefaultStyle ? { ...defaultStyle } : undefined,
			})
		}

		// Add the interpolated value (if exists)
		if (i < values.length) {
			const value = values[i]
			if (isStyledTextFragment(value)) {
				// Merge default style with fragment style (fragment wins)
				result.push({
					text: value.text,
					options: { ...defaultStyle, ...value.style },
				})
			} else if (typeof value === 'string' && value) {
				// Plain string interpolation
				result.push({
					text: value,
					options: hasDefaultStyle ? { ...defaultStyle } : undefined,
				})
			}
		}
	}

	return result
}
