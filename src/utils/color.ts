/**
 * Color Utilities
 *
 * Provides color manipulation functions for gradients, interpolation, etc.
 */

/**
 * Parse a hex color string to RGB components.
 * Accepts formats: "RRGGBB", "#RRGGBB", "RGB", "#RGB"
 */
export function parseHexColorToRgb(hex: string): { r: number; g: number; b: number } {
	// Remove # prefix if present
	let cleanHex = hex.replace(/^#/, '')

	// Expand shorthand (RGB -> RRGGBB)
	if (cleanHex.length === 3) {
		cleanHex = cleanHex[0] + cleanHex[0] + cleanHex[1] + cleanHex[1] + cleanHex[2] + cleanHex[2]
	}

	const r = parseInt(cleanHex.substring(0, 2), 16)
	const g = parseInt(cleanHex.substring(2, 4), 16)
	const b = parseInt(cleanHex.substring(4, 6), 16)

	return { r, g, b }
}

/**
 * Convert RGB components to a hex color string (without #).
 */
export function rgbToHexColor(r: number, g: number, b: number): string {
	const toHex = (n: number) => Math.round(Math.max(0, Math.min(255, n))).toString(16).padStart(2, '0')
	return (toHex(r) + toHex(g) + toHex(b)).toUpperCase()
}

/**
 * Interpolate between two colors.
 * @param colorFrom - Starting color (hex string)
 * @param colorTo - Ending color (hex string)
 * @param ratio - Interpolation ratio (0 = colorFrom, 1 = colorTo)
 * @returns Interpolated hex color string
 */
export function interpolateColor(colorFrom: string, colorTo: string, ratio: number): string {
	const from = parseHexColorToRgb(colorFrom)
	const to = parseHexColorToRgb(colorTo)

	const r = from.r + (to.r - from.r) * ratio
	const g = from.g + (to.g - from.g) * ratio
	const b = from.b + (to.b - from.b) * ratio

	return rgbToHexColor(r, g, b)
}

/**
 * Generate an array of interpolated colors between two colors.
 * @param colorFrom - Starting color (hex string)
 * @param colorTo - Ending color (hex string)
 * @param steps - Number of colors to generate (including start and end)
 * @returns Array of hex color strings
 *
 * @example
 * interpolateColors('1E88E5', '26A69A', 4)
 * // Returns: ['1E88E5', '2192C0', '249D9D', '26A69A']
 */
export function interpolateColors(colorFrom: string, colorTo: string, steps: number): string[] {
	if (steps < 2) {
		return [colorFrom]
	}

	const colors: string[] = []
	for (let i = 0; i < steps; i++) {
		const ratio = i / (steps - 1)
		colors.push(interpolateColor(colorFrom, colorTo, ratio))
	}
	return colors
}

/**
 * Lighten a color by a percentage.
 * @param color - Hex color string
 * @param amount - Amount to lighten (0-1, where 1 = white)
 */
export function lightenColor(color: string, amount: number): string {
	return interpolateColor(color, 'FFFFFF', amount)
}

/**
 * Darken a color by a percentage.
 * @param color - Hex color string
 * @param amount - Amount to darken (0-1, where 1 = black)
 */
export function darkenColor(color: string, amount: number): string {
	return interpolateColor(color, '000000', amount)
}
