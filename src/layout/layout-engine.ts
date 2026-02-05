/**
 * Layout Engine
 *
 * Calculates absolute positions from declarative layout specifications.
 * Supports Grid, Stack, and Flex layouts.
 */

import type { Coord } from '../core-interfaces'

/**
 * A positioned rectangle with computed coordinates.
 */
export interface ComputedBounds {
	x: number
	y: number
	w: number
	h: number
}

/**
 * Gap can be a single number or separate x/y values.
 */
export type GapValue = number | { x: number; y: number }

/**
 * Normalize gap to { x, y } format.
 */
export function normalizeGapValue(gap: GapValue | undefined): { x: number; y: number } {
	if (gap === undefined) return { x: 0, y: 0 }
	if (typeof gap === 'number') return { x: gap, y: gap }
	return gap
}

// ============================================================================
// Grid Layout
// ============================================================================

export interface GridLayoutOptions {
	/** X position of the grid container */
	x: number
	/** Y position of the grid container */
	y: number
	/** Total width of the grid (optional - calculated from cols * cellWidth + gaps if omitted) */
	w?: number
	/** Total height of the grid (optional - calculated from rows * cellHeight + gaps if omitted) */
	h?: number
	/** Number of columns */
	cols: number
	/** Number of rows (optional - calculated from children count if omitted) */
	rows?: number
	/** Gap between cells */
	gap?: GapValue
	/** Width of each cell (optional - calculated from w/cols if w provided) */
	cellWidth?: number
	/** Height of each cell (optional - calculated from h/rows if h provided) */
	cellHeight?: number
}

/**
 * Calculate positions for children in a grid layout.
 * Returns an array of ComputedBounds for each child position.
 */
export function calculateGridLayout(
	options: GridLayoutOptions,
	childrenCount: number
): ComputedBounds[] {
	const { x, y, cols } = options
	const gap = normalizeGapValue(options.gap)

	// Calculate rows if not specified
	const rows = options.rows ?? Math.ceil(childrenCount / cols)

	// Calculate cell dimensions
	let cellWidth: number
	let cellHeight: number

	if (options.cellWidth !== undefined) {
		cellWidth = options.cellWidth
	} else if (options.w !== undefined) {
		cellWidth = (options.w - gap.x * (cols - 1)) / cols
	} else {
		throw new Error('Grid layout requires either cellWidth or w to be specified')
	}

	if (options.cellHeight !== undefined) {
		cellHeight = options.cellHeight
	} else if (options.h !== undefined) {
		cellHeight = (options.h - gap.y * (rows - 1)) / rows
	} else {
		throw new Error('Grid layout requires either cellHeight or h to be specified')
	}

	// Calculate positions for each child
	const positions: ComputedBounds[] = []
	for (let i = 0; i < childrenCount; i++) {
		const col = i % cols
		const row = Math.floor(i / cols)

		positions.push({
			x: x + col * (cellWidth + gap.x),
			y: y + row * (cellHeight + gap.y),
			w: cellWidth,
			h: cellHeight,
		})
	}

	return positions
}

// ============================================================================
// Stack Layout
// ============================================================================

export interface StackLayoutOptions {
	/** X position of the stack container */
	x: number
	/** Y position of the stack container */
	y: number
	/** Direction of stacking */
	direction: 'vertical' | 'horizontal'
	/** Gap between items */
	gap?: number
	/** Width of each item (for vertical stacks) or total width (for horizontal) */
	itemWidth?: number
	/** Height of each item (for horizontal stacks) or total height (for vertical) */
	itemHeight?: number
}

export interface StackChildSize {
	w?: number
	h?: number
}

/**
 * Calculate positions for children in a stack layout.
 * Each child can have its own size, or use the default itemWidth/itemHeight.
 */
export function calculateStackLayout(
	options: StackLayoutOptions,
	childSizes: StackChildSize[]
): ComputedBounds[] {
	const { x, y, direction, gap = 0 } = options
	const positions: ComputedBounds[] = []

	let currentOffset = 0

	for (const childSize of childSizes) {
		const w = childSize.w ?? options.itemWidth ?? 0
		const h = childSize.h ?? options.itemHeight ?? 0

		if (direction === 'vertical') {
			positions.push({
				x,
				y: y + currentOffset,
				w,
				h,
			})
			currentOffset += h + gap
		} else {
			positions.push({
				x: x + currentOffset,
				y,
				w,
				h,
			})
			currentOffset += w + gap
		}
	}

	return positions
}

// ============================================================================
// Flex Layout (simplified)
// ============================================================================

export interface FlexLayoutOptions {
	/** X position of the flex container */
	x: number
	/** Y position of the flex container */
	y: number
	/** Total width of the container */
	w: number
	/** Total height of the container */
	h: number
	/** Direction */
	direction: 'row' | 'column'
	/** Gap between items */
	gap?: number
	/** Wrap to next row/column */
	wrap?: boolean
	/** Justify content along main axis */
	justify?: 'start' | 'center' | 'end' | 'space-between' | 'space-around'
	/** Align items along cross axis */
	align?: 'start' | 'center' | 'end' | 'stretch'
}

export interface FlexChildSize {
	w: number
	h: number
	/** Flex grow factor (default 0) */
	grow?: number
}

/**
 * Calculate positions for children in a flex layout.
 * This is a simplified flexbox implementation for common use cases.
 */
export function calculateFlexLayout(
	options: FlexLayoutOptions,
	childSizes: FlexChildSize[]
): ComputedBounds[] {
	const { x, y, w, h, direction, gap = 0, justify = 'start', align = 'start' } = options
	const positions: ComputedBounds[] = []

	const isRow = direction === 'row'
	const mainAxisSize = isRow ? w : h
	const crossAxisSize = isRow ? h : w

	// Calculate total size of children along main axis
	const totalChildMainSize = childSizes.reduce(
		(sum, child) => sum + (isRow ? child.w : child.h),
		0
	)
	const totalGaps = gap * (childSizes.length - 1)
	const freeSpace = mainAxisSize - totalChildMainSize - totalGaps

	// Calculate starting offset based on justify
	let mainOffset: number
	let gapBetween = gap

	switch (justify) {
		case 'center':
			mainOffset = freeSpace / 2
			break
		case 'end':
			mainOffset = freeSpace
			break
		case 'space-between':
			mainOffset = 0
			gapBetween = childSizes.length > 1 ? (freeSpace + totalGaps) / (childSizes.length - 1) : 0
			break
		case 'space-around':
			gapBetween = (freeSpace + totalGaps) / childSizes.length
			mainOffset = gapBetween / 2
			break
		default: // 'start'
			mainOffset = 0
	}

	for (const childSize of childSizes) {
		const childMainSize = isRow ? childSize.w : childSize.h
		let childCrossSize = isRow ? childSize.h : childSize.w

		// Handle cross-axis alignment
		let crossOffset: number
		if (align === 'stretch') {
			childCrossSize = crossAxisSize
			crossOffset = 0
		} else if (align === 'center') {
			crossOffset = (crossAxisSize - childCrossSize) / 2
		} else if (align === 'end') {
			crossOffset = crossAxisSize - childCrossSize
		} else {
			crossOffset = 0
		}

		if (isRow) {
			positions.push({
				x: x + mainOffset,
				y: y + crossOffset,
				w: childSize.w,
				h: align === 'stretch' ? crossAxisSize : childSize.h,
			})
		} else {
			positions.push({
				x: x + crossOffset,
				y: y + mainOffset,
				w: align === 'stretch' ? crossAxisSize : childSize.w,
				h: childSize.h,
			})
		}

		mainOffset += childMainSize + gapBetween
	}

	return positions
}
