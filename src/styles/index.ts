/**
 * Styles Module
 *
 * Provides style presets and resolution utilities for the compositional API.
 */

export { SHADOW_PRESETS, resolveShadowPreset } from './shadow-presets'
export type { ShadowPresetName, ShadowWithPreset } from './shadow-presets'

import type { TextBaseProps, ShapeFillProps, ShadowProps } from '../core-interfaces'
import type { ShadowPresetName } from './shadow-presets'

/**
 * Text style that can be applied to text elements.
 * Extends TextBaseProps with convenience properties.
 */
export interface TextStyle extends TextBaseProps {
	shadow?: ShadowPresetName | ShadowProps
}

/**
 * Common defaults that can be set on a slide and cascade to children.
 */
export interface SlideDefaults {
	fontFace?: string
	fontSize?: number
	color?: string
	fill?: ShapeFillProps | string
	shadow?: ShadowPresetName | ShadowProps
}

/**
 * Merge defaults with specific options, where specific options take precedence.
 */
export function mergeWithDefaults<T extends Record<string, unknown>>(
	defaults: Partial<T> | undefined,
	specific: T
): T {
	if (!defaults) return specific
	return { ...defaults, ...specific }
}

/**
 * Normalize a fill value - accepts hex string or full ShapeFillProps
 */
export function normalizeFillValue(fill: string | ShapeFillProps | undefined): ShapeFillProps | undefined {
	if (fill === undefined) return undefined
	if (typeof fill === 'string') {
		return { color: fill }
	}
	return fill
}
