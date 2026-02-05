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
import type { ShadowPresetName } from '../styles'
import { resolveShadowPreset, normalizeFillValue } from '../styles'

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
	/** Border color */
	borderColor?: HexColor
	/** Border width */
	borderWidth?: number

	// Shadow
	/** Shadow preset or full config */
	shadow?: ShadowPresetName | ShadowProps

	// Padding
	/** Padding inside the card */
	padding?: PaddingValue

	// Content
	/** Heading text */
	heading?: string
	/** Heading color */
	headingColor?: HexColor
	/** Heading font size */
	headingFontSize?: number
	/** Heading font face */
	headingFontFace?: string
	/** Heading bold */
	headingBold?: boolean

	/** Body text */
	body?: string
	/** Body color */
	bodyColor?: HexColor
	/** Body font size */
	bodyFontSize?: number
	/** Body font face */
	bodyFontFace?: string

	/** Gap between heading and body */
	contentGap?: number
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
	headingColor: '333333',
	headingFontSize: 16,
	headingFontFace: 'Arial',
	headingBold: true,
	bodyColor: '555555',
	bodyFontSize: 13,
	bodyFontFace: 'Arial',
	contentGap: 0.15,
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
	borderColor: HexColor
	borderWidth: number
	shadow: ShadowProps | undefined

	// Padding
	padding: { top: number; right: number; bottom: number; left: number }

	// Heading
	heading: string | undefined
	headingColor: HexColor
	headingFontSize: number
	headingFontFace: string
	headingBold: boolean
	headingX: number
	headingY: number
	headingW: number

	// Body
	body: string | undefined
	bodyColor: HexColor
	bodyFontSize: number
	bodyFontFace: string
	bodyX: number
	bodyY: number
	bodyW: number
	bodyH: number
}

/**
 * Resolve card options with defaults and calculate internal positions.
 */
export function resolveCardConfig(options: CardOptions): ResolvedCardConfig {
	const padding = normalizePaddingValue(options.padding ?? CARD_DEFAULTS.padding)

	const headingFontSize = options.headingFontSize ?? CARD_DEFAULTS.headingFontSize
	const contentGap = options.contentGap ?? CARD_DEFAULTS.contentGap

	// Calculate heading height (approximate based on font size)
	// 1 point ≈ 1/72 inch, add some padding
	const headingHeight = options.heading ? (headingFontSize / 72) * 1.5 : 0

	// Content area calculations
	const contentX = options.x + padding.left
	const contentY = options.y + padding.top
	const contentW = options.w - padding.left - padding.right
	const contentH = options.h - padding.top - padding.bottom

	// Body position (below heading)
	const bodyY = contentY + headingHeight + (options.heading ? contentGap : 0)
	const bodyH = contentH - headingHeight - (options.heading ? contentGap : 0)

	return {
		x: options.x,
		y: options.y,
		w: options.w,
		h: options.h,

		backgroundFill: normalizeFillValue(options.background ?? CARD_DEFAULTS.background) ?? { color: CARD_DEFAULTS.background },
		borderRadius: options.borderRadius ?? CARD_DEFAULTS.borderRadius,
		borderColor: options.borderColor ?? CARD_DEFAULTS.borderColor,
		borderWidth: options.borderWidth ?? CARD_DEFAULTS.borderWidth,
		shadow: resolveShadowPreset(options.shadow ?? CARD_DEFAULTS.shadow),

		padding,

		heading: options.heading,
		headingColor: options.headingColor ?? CARD_DEFAULTS.headingColor,
		headingFontSize,
		headingFontFace: options.headingFontFace ?? CARD_DEFAULTS.headingFontFace,
		headingBold: options.headingBold ?? CARD_DEFAULTS.headingBold,
		headingX: contentX,
		headingY: contentY,
		headingW: contentW,

		body: options.body,
		bodyColor: options.bodyColor ?? CARD_DEFAULTS.bodyColor,
		bodyFontSize: options.bodyFontSize ?? CARD_DEFAULTS.bodyFontSize,
		bodyFontFace: options.bodyFontFace ?? CARD_DEFAULTS.bodyFontFace,
		bodyX: contentX,
		bodyY,
		bodyW: contentW,
		bodyH,
	}
}
