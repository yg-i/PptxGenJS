/**
 * Card Component
 *
 * A high-level component for creating cards with:
 * - Rounded rectangle background
 * - Optional shadow
 * - Heading and body text
 * - Consistent padding and styling
 */

import type { ShadowProps, ShapeFillProps, HexColor } from '../core-interfaces'
import type { ShadowPresetName, ShadowWithPreset } from '../styles'
import { resolveShadowPreset, normalizeFillValue } from '../styles'

/**
 * Border can be false/'none' to disable, or an object with color/width.
 */
export type BorderValue = false | 'none' | {
	color?: HexColor
	width?: number
}

/**
 * Padding can be a single number or per-side values.
 */
export type PaddingValue = number | {
	top?: number
	right?: number
	bottom?: number
	left?: number
}

/**
 * Normalize padding to { top, right, bottom, left } format.
 */
export function normalizePaddingValue(padding: PaddingValue | undefined): {
	top: number
	right: number
	bottom: number
	left: number
} {
	const defaultPadding = 0.2
	if (padding === undefined) {
		return { top: defaultPadding, right: defaultPadding, bottom: defaultPadding, left: defaultPadding }
	}
	if (typeof padding === 'number') {
		return { top: padding, right: padding, bottom: padding, left: padding }
	}
	return {
		top: padding.top ?? defaultPadding,
		right: padding.right ?? defaultPadding,
		bottom: padding.bottom ?? defaultPadding,
		left: padding.left ?? defaultPadding,
	}
}

/**
 * Text alignment for card content.
 */
export type CardAlign = 'left' | 'center' | 'right'

/**
 * Options for creating a card.
 */
export interface CardOptions {
	/** X position (inches) */
	x: number
	/** Y position (inches) */
	y: number
	/** Width (inches) */
	w: number
	/** Height (inches) */
	h: number

	// Background & border
	/** Background color (hex string or ShapeFillProps) */
	background?: string | ShapeFillProps
	/** Border radius for rounded corners */
	borderRadius?: number
	/**
	 * Border configuration. Can be:
	 * - false or 'none' to disable border completely
	 * - { color, width } object for custom border
	 * - Omit to use default border
	 */
	border?: BorderValue
	/** @deprecated Use `border: { color }` instead */
	borderColor?: HexColor
	/** @deprecated Use `border: { width }` instead */
	borderWidth?: number

	// Shadow
	/** Shadow preset, preset with overrides, or full config */
	shadow?: ShadowPresetName | ShadowWithPreset | ShadowProps

	// Padding
	/** Padding inside the card */
	padding?: PaddingValue

	/**
	 * Text alignment for all card text content.
	 * Default: 'left'
	 */
	align?: CardAlign

	// Content - Title (small text above heading)
	/** Title text (appears above heading, smaller font) */
	title?: string
	/** Title color */
	titleColor?: HexColor
	/** Title font size (default: 14) */
	titleFontSize?: number
	/** Title font face */
	titleFontFace?: string

	// Content - Heading (main prominent text)
	/** Heading text (main prominent text) */
	heading?: string
	/** Heading color */
	headingColor?: HexColor
	/** Heading font size */
	headingFontSize?: number
	/** Heading font face */
	headingFontFace?: string
	/** Heading bold */
	headingBold?: boolean
	/**
	 * Heading line height multiplier.
	 * Default: 1.5 (heading height = fontSize * lineHeight / 72)
	 */
	headingLineHeight?: number

	// Content - Body (description text)
	/** Body text */
	body?: string
	/** Body color */
	bodyColor?: HexColor
	/** Body font size */
	bodyFontSize?: number
	/** Body font face */
	bodyFontFace?: string
	/** Body italic */
	bodyItalic?: boolean

	/** Gap between content sections */
	contentGap?: number

	/**
	 * Highlight this card with a different background color.
	 * When true, uses the highlight color. Can also be a hex color string.
	 */
	highlight?: boolean | string
}

/**
 * Default card styling.
 */
export const CARD_DEFAULTS = {
	background: 'F5F5F5',
	borderRadius: 0.1,
	borderColor: 'E0E0E0',
	borderWidth: 1,
	shadow: 'sm' as ShadowPresetName,
	padding: 0.2,
	align: 'left' as CardAlign,
	// Title defaults
	titleColor: '555555',
	titleFontSize: 14,
	titleFontFace: 'Arial',
	titleLineHeight: 1.4,
	// Heading defaults
	headingColor: '333333',
	headingFontSize: 16,
	headingFontFace: 'Arial',
	headingBold: true,
	headingLineHeight: 1.5,
	// Body defaults
	bodyColor: '555555',
	bodyFontSize: 13,
	bodyFontFace: 'Arial',
	bodyItalic: false,
	contentGap: 0.15,
	highlightColor: 'E3F2FD', // Light blue for highlighted cards
}

/**
 * Resolved card configuration with all values filled in.
 */
export interface ResolvedCardConfig {
	x: number
	y: number
	w: number
	h: number

	// Background shape
	backgroundFill: ShapeFillProps
	borderRadius: number
	/** Border color, or undefined if border is disabled */
	borderColor: HexColor | undefined
	/** Border width, or 0 if border is disabled */
	borderWidth: number
	/** Whether border is enabled */
	hasBorder: boolean
	shadow: ShadowProps | undefined

	// Padding
	padding: { top: number; right: number; bottom: number; left: number }

	// Alignment
	align: CardAlign

	// Title (small text above heading)
	title: string | undefined
	titleColor: HexColor
	titleFontSize: number
	titleFontFace: string
	titleX: number
	titleY: number
	titleW: number
	titleH: number

	// Heading
	heading: string | undefined
	headingColor: HexColor
	headingFontSize: number
	headingFontFace: string
	headingBold: boolean
	headingLineHeight: number
	headingX: number
	headingY: number
	headingW: number
	headingH: number

	// Body
	body: string | undefined
	bodyColor: HexColor
	bodyFontSize: number
	bodyFontFace: string
	bodyItalic: boolean
	bodyX: number
	bodyY: number
	bodyW: number
	bodyH: number
}

/**
 * Normalize border value to { color, width, enabled } format.
 */
function normalizeBorderValue(
	border: BorderValue | undefined,
	legacyColor: HexColor | undefined,
	legacyWidth: number | undefined
): { color: HexColor | undefined; width: number; enabled: boolean } {
	// If border is explicitly disabled
	if (border === false || border === 'none') {
		return { color: undefined, width: 0, enabled: false }
	}

	// If border is an object with color/width
	if (typeof border === 'object') {
		return {
			color: border.color ?? CARD_DEFAULTS.borderColor,
			width: border.width ?? CARD_DEFAULTS.borderWidth,
			enabled: true,
		}
	}

	// Use legacy properties or defaults
	return {
		color: legacyColor ?? CARD_DEFAULTS.borderColor,
		width: legacyWidth ?? CARD_DEFAULTS.borderWidth,
		enabled: true,
	}
}

/**
 * Resolve background color, considering highlight option.
 */
function resolveBackgroundColor(
	background: string | ShapeFillProps | undefined,
	highlight: boolean | string | undefined
): ShapeFillProps {
	// If highlight is a color string, use it
	if (typeof highlight === 'string') {
		return { color: highlight }
	}

	// If highlight is true, use default highlight color
	if (highlight === true) {
		return { color: CARD_DEFAULTS.highlightColor }
	}

	// Otherwise use the provided background or default
	return normalizeFillValue(background ?? CARD_DEFAULTS.background) ?? { color: CARD_DEFAULTS.background }
}

/**
 * Resolve card options with defaults and calculate internal positions.
 */
export function resolveCardConfig(options: CardOptions): ResolvedCardConfig {
	const padding = normalizePaddingValue(options.padding ?? CARD_DEFAULTS.padding)
	const align = options.align ?? CARD_DEFAULTS.align
	const contentGap = options.contentGap ?? CARD_DEFAULTS.contentGap

	// Font sizes
	const titleFontSize = options.titleFontSize ?? CARD_DEFAULTS.titleFontSize
	const headingFontSize = options.headingFontSize ?? CARD_DEFAULTS.headingFontSize
	const bodyFontSize = options.bodyFontSize ?? CARD_DEFAULTS.bodyFontSize

	// Line heights
	const titleLineHeight = CARD_DEFAULTS.titleLineHeight
	const headingLineHeight = options.headingLineHeight ?? CARD_DEFAULTS.headingLineHeight

	// Calculate section heights (1 point ≈ 1/72 inch)
	const titleHeight = options.title ? (titleFontSize / 72) * titleLineHeight : 0
	const headingHeight = options.heading ? (headingFontSize / 72) * headingLineHeight : 0

	// Content area calculations
	const contentX = options.x + padding.left
	const contentY = options.y + padding.top
	const contentW = options.w - padding.left - padding.right
	const contentH = options.h - padding.top - padding.bottom

	// Calculate Y positions for each section
	let currentY = contentY

	// Title position
	const titleY = currentY
	if (options.title) {
		currentY += titleHeight + contentGap
	}

	// Heading position
	const headingY = currentY
	if (options.heading) {
		currentY += headingHeight + contentGap
	}

	// Body position (remaining space)
	const bodyY = currentY
	const bodyH = Math.max(0, contentH - (currentY - contentY))

	// Resolve border
	const border = normalizeBorderValue(options.border, options.borderColor, options.borderWidth)

	return {
		x: options.x,
		y: options.y,
		w: options.w,
		h: options.h,

		backgroundFill: resolveBackgroundColor(options.background, options.highlight),
		borderRadius: options.borderRadius ?? CARD_DEFAULTS.borderRadius,
		borderColor: border.color,
		borderWidth: border.width,
		hasBorder: border.enabled,
		shadow: resolveShadowPreset(options.shadow ?? CARD_DEFAULTS.shadow),

		padding,
		align,

		// Title
		title: options.title,
		titleColor: options.titleColor ?? CARD_DEFAULTS.titleColor,
		titleFontSize,
		titleFontFace: options.titleFontFace ?? CARD_DEFAULTS.titleFontFace,
		titleX: contentX,
		titleY,
		titleW: contentW,
		titleH: titleHeight,

		// Heading
		heading: options.heading,
		headingColor: options.headingColor ?? CARD_DEFAULTS.headingColor,
		headingFontSize,
		headingFontFace: options.headingFontFace ?? CARD_DEFAULTS.headingFontFace,
		headingBold: options.headingBold ?? CARD_DEFAULTS.headingBold,
		headingLineHeight,
		headingX: contentX,
		headingY,
		headingW: contentW,
		headingH: headingHeight,

		// Body
		body: options.body,
		bodyColor: options.bodyColor ?? CARD_DEFAULTS.bodyColor,
		bodyFontSize,
		bodyFontFace: options.bodyFontFace ?? CARD_DEFAULTS.bodyFontFace,
		bodyItalic: options.bodyItalic ?? CARD_DEFAULTS.bodyItalic,
		bodyX: contentX,
		bodyY,
		bodyW: contentW,
		bodyH,
	}
}
