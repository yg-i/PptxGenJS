/**
 * Shadow Presets
 *
 * Provides shorthand names for common shadow configurations.
 * Use 'sm', 'md', 'lg' instead of specifying 6+ properties.
 */

import type { ShadowProps } from '../core-interfaces'

export type ShadowPresetName = 'none' | 'sm' | 'md' | 'lg' | 'xl'

export const SHADOW_PRESETS: Record<ShadowPresetName, ShadowProps | undefined> = {
	none: undefined,

	sm: {
		type: 'outer',
		blur: 3,
		offset: 1,
		angle: 45,
		color: '000000',
		opacity: 0.1,
	},

	md: {
		type: 'outer',
		blur: 6,
		offset: 3,
		angle: 45,
		color: '000000',
		opacity: 0.15,
	},

	lg: {
		type: 'outer',
		blur: 10,
		offset: 5,
		angle: 45,
		color: '000000',
		opacity: 0.2,
	},

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
 * Resolve a shadow value - accepts preset name, full config, or undefined.
 */
export function resolveShadowPreset(
	shadowValue: ShadowPresetName | ShadowProps | undefined
): ShadowProps | undefined {
	if (shadowValue === undefined || shadowValue === 'none') {
		return undefined
	}

	if (typeof shadowValue === 'string') {
		const preset = SHADOW_PRESETS[shadowValue]
		if (!preset) {
			console.warn(`PptxGenJS: Unknown shadow preset '${shadowValue}', using 'md'`)
			return SHADOW_PRESETS.md
		}
		return preset
	}

	return shadowValue
}
