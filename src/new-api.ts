/**
 * PptxGenJS New API - Clean, compositional, declarative
 *
 * Elements are data. Layouts position children. Animation is a wrapper.
 * XML only generates at write() time.
 */

// ============================================================================
// TYPES
// ============================================================================

export type HexColor = string

export interface ElementNode {
	readonly _tag: string
	// eslint-disable-next-line @typescript-eslint/no-explicit-any
	readonly props: Record<string, any>
	readonly children: ElementNode[]
	readonly animation?: AnimationConfig
}

export interface AnimationConfig {
	type: 'fade' | 'fly-in' | 'zoom' | 'wipe' | 'appear'
	direction?: 'from-left' | 'from-right' | 'from-top' | 'from-bottom'
	trigger: 'onClick' | 'withPrevious' | 'afterPrevious'
	durationMs?: number
	delayMs?: number
	/** For containers: stagger children animations by this many ms */
	stagger?: number
}

export interface SlideConfig {
	background?: HexColor | GradientConfig
	defaults?: {
		fontFace?: string
		fontSize?: number
		color?: HexColor
		accentColor?: HexColor
	}
}

export interface GradientConfig {
	type?: 'linear'
	angle?: number
	stops: Array<{ position: number; color: HexColor }>
}

export interface TextConfig {
	fontSize?: number
	fontFace?: string
	color?: HexColor
	bold?: boolean
	italic?: boolean
	align?: 'left' | 'center' | 'right'
	/** Position - if omitted, parent layout handles it */
	x?: number
	y?: number
	w?: number
	h?: number
}

export interface ShapeConfig {
	type?: 'rect' | 'roundRect' | 'ellipse' | 'line'
	fill?: HexColor
	line?: { color: HexColor; width?: number }
	rectRadius?: number
	x?: number
	y?: number
	w?: number
	h?: number
}

export interface PillConfig {
	text: string
	fill: HexColor
	color?: HexColor
	fontSize?: number
	fontFace?: string
	h?: number
	rectRadius?: number
}

export interface CardConfig {
	heading?: string
	body?: string
	fill?: HexColor
	headingColor?: HexColor
	bodyColor?: HexColor
	shadow?: 'none' | 'sm' | 'md' | 'lg'
	h?: number
	w?: number
	x?: number
	y?: number
}

export interface StackConfig {
	gap?: number
	x?: number
	y?: number
	w?: number
}

export interface ColumnsConfig {
	gap?: number
	x?: number
	y?: number
	w?: number
	h?: number
	/** Ratio for each column, e.g., [1, 1] for equal, [2, 1] for 2:1 */
	ratio?: number[]
}

export interface GridConfig {
	cols: number
	gap?: number
	x?: number
	y?: number
	w?: number
	h?: number
}

export interface NumberedListConfig {
	x?: number
	y?: number
	w?: number
	fontSize?: number
	fontFace?: string
	color?: HexColor
	accentColor?: HexColor
	itemGap?: number
}

export interface ImageConfig {
	path?: string
	data?: string  // base64
	x?: number
	y?: number
	w?: number
	h?: number
	sizing?: { type: 'cover' | 'contain'; w: number; h: number }
}

// ============================================================================
// ELEMENT CONSTRUCTORS
// ============================================================================

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function createElement(tag: string, props: Record<string, any>, children: ElementNode[] = []): ElementNode {
	return { _tag: tag, props, children }
}

// Import styled fragment type
import type { StyledFragment } from './new-api-styles'
import { processStyledTemplate } from './new-api-styles'

// Re-export all style builders
export {
	style,
	bold,
	italic,
	underline,
	strike,
	white,
	black,
	gray,
	red,
	orange,
	yellow,
	green,
	cyan,
	blue,
	purple,
	magenta,
	pink,
	color,
	processStyledTemplate,
} from './new-api-styles'
export type { StyledFragment, TextStyle, StyleBuilder } from './new-api-styles'

/** Content that can be passed to Text - either plain string or styled fragments */
export type TextContent = string | Array<string | StyledFragment>

/**
 * Text element - two ways to use:
 *
 * 1. Simple string:
 *    Text("Hello world", { fontSize: 24 })
 *
 * 2. Tagged template with styled fragments:
 *    Text`Hello ${bold.cyan("world")}!`
 *
 * Chain modifiers for styling:
 *    Text`...`.config({ fontSize: 24 })
 */
export function Text(content: string, config?: TextConfig): ElementNode
export function Text(strings: TemplateStringsArray, ...values: (string | StyledFragment)[]): TextElement
export function Text(
	contentOrStrings: string | TemplateStringsArray,
	configOrFirstValue?: TextConfig | string | StyledFragment,
	...restValues: (string | StyledFragment)[]
): ElementNode | TextElement {
	// Case 1: Simple string - Text("hello", { fontSize: 24 })
	if (typeof contentOrStrings === 'string') {
		const config = (configOrFirstValue as TextConfig) || {}
		return createElement('text', { content: contentOrStrings, ...config }, [])
	}

	// Case 2: Tagged template - Text`hello ${bold("world")}`
	const strings = contentOrStrings as TemplateStringsArray
	const values = configOrFirstValue !== undefined
		? [configOrFirstValue as string | StyledFragment, ...restValues]
		: []

	const fragments = processStyledTemplate(strings, ...values)

	// Return a TextElement that can be further configured
	return createTextElement(fragments)
}

/** A Text element that supports chaining config */
export interface TextElement extends ElementNode {
	/** Apply configuration to this text element */
	config(config: TextConfig): TextElement
}

function createTextElement(fragments: Array<string | StyledFragment>, baseConfig: TextConfig = {}): TextElement {
	const node = createElement('text', { fragments, ...baseConfig }, []) as TextElement

	// Add chainable config method
	node.config = (config: TextConfig): TextElement => {
		return createTextElement(fragments, { ...baseConfig, ...config })
	}

	return node
}

/** Basic shape */
export function Shape(config: ShapeConfig): ElementNode {
	return createElement('shape', config, [])
}

/** Pill - rounded rectangle with centered text */
export function Pill(config: PillConfig): ElementNode {
	return createElement('pill', config, [])
}

/** Card - box with optional heading, body, shadow */
export function Card(config: CardConfig, children: ElementNode[] = []): ElementNode {
	return createElement('card', config, children)
}

/** Image */
export function Image(config: ImageConfig): ElementNode {
	return createElement('image', config, [])
}

/** Numbered list - items support **bold** markdown */
export function NumberedList(config: NumberedListConfig, items: string[]): ElementNode {
	return createElement('numberedList', { ...config, items }, [])
}

/** Bullet list */
export function BulletList(config: NumberedListConfig, items: string[]): ElementNode {
	return createElement('bulletList', { ...config, items }, [])
}

// ============================================================================
// LAYOUT CONTAINERS
// ============================================================================

/** Vertical stack - positions children top-to-bottom */
export function Stack(config: StackConfig, children: ElementNode[]): ElementNode {
	return createElement('stack', config, children)
}

/** Horizontal columns */
export function Columns(config: ColumnsConfig, children: ElementNode[]): ElementNode {
	return createElement('columns', config, children)
}

/** Grid layout */
export function Grid(config: GridConfig, children: ElementNode[]): ElementNode {
	return createElement('grid', config, children)
}

/** Absolute positioning escape hatch */
export function Absolute(config: { x: number; y: number; w?: number; h?: number }, child: ElementNode): ElementNode {
	return createElement('absolute', config, [child])
}

// ============================================================================
// ANIMATION WRAPPERS
// ============================================================================

type PartialAnimationConfig = Omit<AnimationConfig, 'type' | 'trigger'> & {
	onClick?: boolean
	withPrevious?: boolean
	afterPrevious?: boolean
}

function resolveAnimationTrigger(config: PartialAnimationConfig): AnimationConfig['trigger'] {
	if (config.onClick) return 'onClick'
	if (config.withPrevious) return 'withPrevious'
	if (config.afterPrevious) return 'afterPrevious'
	return 'onClick' // default
}

function wrapWithAnimation(type: AnimationConfig['type'], config: PartialAnimationConfig, child: ElementNode): ElementNode {
	const animation: AnimationConfig = {
		type,
		trigger: resolveAnimationTrigger(config),
		direction: config.direction,
		durationMs: config.durationMs,
		delayMs: config.delayMs,
		stagger: config.stagger,
	}
	return { ...child, animation }
}

/** Fade in animation */
export function FadeIn(config: PartialAnimationConfig, child: ElementNode): ElementNode {
	return wrapWithAnimation('fade', config, child)
}

/** Fly in animation */
export function FlyIn(config: PartialAnimationConfig & { direction?: AnimationConfig['direction'] }, child: ElementNode): ElementNode {
	return wrapWithAnimation('fly-in', { direction: 'from-bottom', ...config }, child)
}

/** Zoom in animation */
export function ZoomIn(config: PartialAnimationConfig, child: ElementNode): ElementNode {
	return wrapWithAnimation('zoom', config, child)
}

/** Wipe animation */
export function Wipe(config: PartialAnimationConfig & { direction?: AnimationConfig['direction'] }, child: ElementNode): ElementNode {
	return wrapWithAnimation('wipe', { direction: 'from-left', ...config }, child)
}

/** Instant appear (no visual effect, just click-to-show) */
export function Appear(config: PartialAnimationConfig, child: ElementNode): ElementNode {
	return wrapWithAnimation('appear', config, child)
}

// ============================================================================
// HELPERS
// ============================================================================

/** Create a gradient config */
export function gradient(config: { from: HexColor; to: HexColor; angle?: number }): GradientConfig {
	return {
		type: 'linear',
		angle: config.angle ?? 180,
		stops: [
			{ position: 0, color: config.from },
			{ position: 100, color: config.to },
		],
	}
}

// ============================================================================
// SLIDE
// ============================================================================

export class $Slide {
	readonly config: SlideConfig
	readonly elements: ElementNode[]

	constructor(config: SlideConfig, elements: ElementNode[]) {
		this.config = config
		this.elements = elements
	}
}

// ============================================================================
// PRESENTATION
// ============================================================================

export interface PresentationConfig {
	layout?: '16x9' | '16x10' | '4x3' | 'WIDE' | { width: number; height: number }
	title?: string
	author?: string
	subject?: string
}

export class $pptxNewAPI {
	private readonly config: PresentationConfig
	private readonly slides: $Slide[] = []

	constructor(config: PresentationConfig = {}) {
		this.config = { layout: '16x9', ...config }
	}

	/** Add a slide with elements */
	slide(configOrFirstElement: SlideConfig | ElementNode, ...elements: ElementNode[]): $Slide {
		let config: SlideConfig = {}
		let allElements: ElementNode[]

		// Check if first arg is config or element
		if (configOrFirstElement && '_tag' in configOrFirstElement) {
			// It's an element
			allElements = [configOrFirstElement, ...elements]
		} else {
			// It's config
			config = configOrFirstElement as SlideConfig
			allElements = elements
		}

		const slide = new $Slide(config, allElements)
		this.slides.push(slide)
		return slide
	}

	/** Get all slides */
	getSlides(): readonly $Slide[] {
		return this.slides
	}

	/** Render to old API and write file */
	async write(fileName: string): Promise<void> {
		const { renderPresentationToOldAPI } = await import('./new-api-render')
		await renderPresentationToOldAPI(this, fileName)
	}
}

// ============================================================================
// CONVENIENCE EXPORT
// ============================================================================

export function createPresentation(config?: PresentationConfig): $pptxNewAPI {
	return new $pptxNewAPI(config)
}
