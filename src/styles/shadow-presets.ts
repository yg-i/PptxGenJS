/**
 * Shadow Presets
 *
 * Provides shorthand names for common shadow configurations.
 * Use 'xs', 'sm', 'md', 'lg', 'xl' instead of specifying 6+ properties.
 *
 * You can also extend a preset with custom overrides:
 * { preset: 'sm', opacity: 0.05 }
 */

import type { ShadowProps } from '../core-interfaces'

export type ShadowPresetName = 'none' | 'xs' | 'subtle' | 'sm' | 'md' | 'lg' | 'xl'

/**
 * Shadow with preset reference and optional overrides.
 * Allows extending a preset with custom properties.
 *
 * @example
 * { preset: 'sm', opacity: 0.05 } // Use 'sm' but with lower opacity
 * { preset: 'md', blur: 4 }       // Use 'md' but with less blur
 */
export interface ShadowWithPreset extends Partial<ShadowProps> {
	preset: ShadowPresetName
}

export const SHADOW_PRESETS: Record<ShadowPresetName, ShadowProps | undefined> = {
	none: undefined,

	/** Extra small - very subtle, barely visible */
	xs: {
		type: 'outer',
		blur: 2,
		offset: 0.5,
		angle: 90,
		color: '000000',
		opacity: 0.05,
	},

	/** Subtle - gentle shadow for light interfaces */
	subtle: {
		type: 'outer',
		blur: 3,
		offset: 1,
		angle: 90,
		color: '000000',
		opacity: 0.08,
	},

	/** Small - light shadow */
	sm: {
		type: 'outer',
		blur: 3,
		offset: 1,
		angle: 45,
		color: '000000',
		opacity: 0.1,
	},

	/** Medium - default shadow */
	md: {
		type: 'outer',
		blur: 6,
		offset: 3,
		angle: 45,
		color: '000000',
		opacity: 0.15,
	},

	/** Large - prominent shadow */
	lg: {
		type: 'outer',
		blur: 10,
		offset: 5,
		angle: 45,
		color: '000000',
		opacity: 0.2,
	},

	/** Extra large - dramatic shadow */
	xl: {
		type: 'outer',
		blur: 15,
		offset: 8,
		angle: 45,
		color: '000000',
		opacity: 0.25,
	},
}

/**
 * Check if a value is a ShadowWithPreset object.
 */
function isShadowWithPreset(value: unknown): value is ShadowWithPreset {
	return typeof value === 'object' && value !== null && 'preset' in value
}

/**
 * Resolve a shadow value - accepts preset name, preset with overrides, full config, or undefined.
 *
 * @example
 * resolveShadowPreset('sm')                    // Returns SHADOW_PRESETS.sm
 * resolveShadowPreset({ preset: 'sm', opacity: 0.05 }) // Returns sm preset with opacity override
 * resolveShadowPreset({ type: 'outer', ... })  // Returns the full config as-is
 */
export function resolveShadowPreset(
	shadowValue: ShadowPresetName | ShadowWithPreset | ShadowProps | undefined
): ShadowProps | undefined {
	if (shadowValue === undefined || shadowValue === 'none') {
		return undefined
	}

	// Handle preset name string
	if (typeof shadowValue === 'string') {
		const preset = SHADOW_PRESETS[shadowValue]
		if (!preset) {
			console.warn(`PptxGenJS: Unknown shadow preset '${shadowValue}', using 'md'`)
			return SHADOW_PRESETS.md
		}
		return { ...preset }
	}

	// Handle preset with overrides: { preset: 'sm', opacity: 0.05 }
	if (isShadowWithPreset(shadowValue)) {
		const { preset, ...overrides } = shadowValue
		const basePreset = SHADOW_PRESETS[preset]
		if (!basePreset) {
			console.warn(`PptxGenJS: Unknown shadow preset '${preset}', using 'md'`)
			return { ...SHADOW_PRESETS.md!, ...overrides } as ShadowProps
		}
		return { ...basePreset, ...overrides } as ShadowProps
	}

	// Handle full ShadowProps object
	return shadowValue
}
