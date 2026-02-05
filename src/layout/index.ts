/**
 * Layout Module
 *
 * Exports layout calculation utilities for Grid, Stack, and Flex layouts.
 */

export {
	calculateGridLayout,
	calculateStackLayout,
	calculateFlexLayout,
	normalizeGapValue,
} from './layout-engine'

export type {
	ComputedBounds,
	GapValue,
	GridLayoutOptions,
	StackLayoutOptions,
	StackChildSize,
	FlexLayoutOptions,
	FlexChildSize,
} from './layout-engine'
