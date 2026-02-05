/**
 * Renderer: Converts new API element tree → old PptxGenJS API calls
 */

import PptxGenJS from './pptxgen'
import type { $pptxNewAPI, $Slide, ElementNode, AnimationConfig, GradientConfig, HexColor } from './new-api'
import { ShapeType } from './core-enums'

// ============================================================================
// TYPES
// ============================================================================

interface RenderContext {
	pptx: PptxGenJS
	slide: InstanceType<typeof PptxGenJS.prototype.addSlide>
	defaults: {
		fontFace?: string
		fontSize?: number
		color?: HexColor
		accentColor?: HexColor
	}
	/** Current animation sequence index for onClick ordering */
	animationSequence: number
}

interface Bounds {
	x: number
	y: number
	w: number
	h: number
}

// ============================================================================
// MAIN RENDER FUNCTION
// ============================================================================

export async function renderPresentationToOldAPI(pres: $pptxNewAPI, fileName: string): Promise<void> {
	const pptx = new PptxGenJS()

	// Set layout
	const config = (pres as unknown as { config: { layout?: string | { width: number; height: number } } }).config
	if (config.layout) {
		if (typeof config.layout === 'string') {
			const layoutMap: Record<string, string> = {
				'16x9': 'LAYOUT_16x9',
				'16x10': 'LAYOUT_16x10',
				'4x3': 'LAYOUT_4x3',
				'WIDE': 'LAYOUT_WIDE',
			}
			pptx.layout = layoutMap[config.layout] || 'LAYOUT_16x9'
		} else {
			pptx.defineLayout({ name: 'CUSTOM', width: config.layout.width, height: config.layout.height })
			pptx.layout = 'CUSTOM'
		}
	}

	// Render each slide
	for (const slideData of pres.getSlides()) {
		renderSlide(pptx, slideData)
	}

	await pptx.writeFile({ fileName })
}

function renderSlide(pptx: PptxGenJS, slideData: $Slide): void {
	const slide = pptx.addSlide()

	// Apply slide config
	if (slideData.config.background) {
		if (typeof slideData.config.background === 'string') {
			slide.background = { color: slideData.config.background }
		} else {
			// Gradient
			const grad = slideData.config.background as GradientConfig
			slide.background = {
				type: 'gradient' as const,
				gradient: {
					type: grad.type || 'linear',
					angle: grad.angle || 180,
					stops: grad.stops.map(s => ({ position: s.position, color: s.color })),
				},
			} as never // Type workaround
		}
	}

	const ctx: RenderContext = {
		pptx,
		slide,
		defaults: slideData.config.defaults || {},
		animationSequence: 0,
	}

	// Apply slide-level defaults
	if (ctx.defaults.fontFace) {
		slide.fontFace = ctx.defaults.fontFace
	}
	if (ctx.defaults.color) {
		slide.color = ctx.defaults.color
	}
	if (ctx.defaults.accentColor) {
		slide.accentColor = ctx.defaults.accentColor
	}

	// Render elements with auto-positioning
	let currentY = 0.5 // Start with some margin

	for (const element of slideData.elements) {
		const height = renderElement(ctx, element, { x: 0.5, y: currentY, w: 9, h: 5 })
		currentY += height + 0.3 // Add gap between top-level elements
	}
}

// ============================================================================
// ELEMENT RENDERERS
// ============================================================================

/**
 * Render an element and return its height (for auto-positioning)
 */
function renderElement(ctx: RenderContext, node: ElementNode, bounds: Bounds): number {
	switch (node._tag) {
		case 'text':
			return renderText(ctx, node, bounds)
		case 'shape':
			return renderShape(ctx, node, bounds)
		case 'pill':
			return renderPill(ctx, node, bounds)
		case 'card':
			return renderCard(ctx, node, bounds)
		case 'image':
			return renderImage(ctx, node, bounds)
		case 'numberedList':
			return renderNumberedList(ctx, node, bounds)
		case 'bulletList':
			return renderBulletList(ctx, node, bounds)
		case 'stack':
			return renderStack(ctx, node, bounds)
		case 'columns':
			return renderColumns(ctx, node, bounds)
		case 'grid':
			return renderGrid(ctx, node, bounds)
		case 'absolute':
			return renderAbsolute(ctx, node)
		default:
			console.warn(`Unknown element type: ${node._tag}`)
			return 0
	}
}

interface StyledFragment {
	text: string
	style: {
		bold?: boolean
		italic?: boolean
		underline?: boolean
		strike?: boolean
		color?: string
		fontSize?: number
		fontFace?: string
	}
}

function renderText(ctx: RenderContext, node: ElementNode, bounds: Bounds): number {
	const props = node.props as {
		content?: string                              // Simple string mode
		fragments?: Array<string | StyledFragment>   // Tagged template mode
		fontSize?: number
		fontFace?: string
		color?: HexColor
		bold?: boolean
		italic?: boolean
		align?: 'left' | 'center' | 'right'
		x?: number
		y?: number
		w?: number
		h?: number
	}

	const fontSize = props.fontSize || ctx.defaults.fontSize || 18
	const h = props.h || (fontSize / 72) * 1.5 // Estimate height from font size

	// Determine what to pass to addText
	let textContent: string | Array<{ text: string; options?: Record<string, unknown> }>

	if (props.fragments) {
		// Tagged template mode - convert fragments to TextProps[]
		textContent = props.fragments.map(frag => {
			if (typeof frag === 'string') {
				// Plain string fragment - use base styles
				return {
					text: frag,
					options: {
						color: props.color || ctx.defaults.color,
						bold: props.bold,
						italic: props.italic,
						fontSize,
						fontFace: props.fontFace || ctx.defaults.fontFace,
					},
				}
			} else {
				// Styled fragment - merge with base styles
				return {
					text: frag.text,
					options: {
						color: frag.style.color || props.color || ctx.defaults.color,
						bold: frag.style.bold ?? props.bold,
						italic: frag.style.italic ?? props.italic,
						underline: frag.style.underline,
						strike: frag.style.strike,
						fontSize: frag.style.fontSize || fontSize,
						fontFace: frag.style.fontFace || props.fontFace || ctx.defaults.fontFace,
					},
				}
			}
		})
	} else {
		// Simple string mode
		textContent = props.content || ''
	}

	const ref = ctx.slide.addText(textContent, {
		x: props.x ?? bounds.x,
		y: props.y ?? bounds.y,
		w: props.w ?? bounds.w,
		h,
		fontSize,
		fontFace: props.fontFace || ctx.defaults.fontFace,
		color: props.color || ctx.defaults.color,
		bold: props.bold,
		italic: props.italic,
		align: props.align,
	})

	applyAnimation(ctx, ref, node.animation)

	return h
}

function renderShape(ctx: RenderContext, node: ElementNode, bounds: Bounds): number {
	const props = node.props as {
		type?: 'rect' | 'roundRect' | 'ellipse' | 'line'
		fill?: HexColor
		line?: { color: HexColor; width?: number }
		rectRadius?: number
		x?: number
		y?: number
		w?: number
		h?: number
	}

	const shapeTypeMap: Record<string, typeof ShapeType[keyof typeof ShapeType]> = {
		rect: ShapeType.rect,
		roundRect: ShapeType.roundRect,
		ellipse: ShapeType.ellipse,
		line: ShapeType.line,
	}

	const h = props.h ?? 1

	const ref = ctx.slide.addShape(shapeTypeMap[props.type || 'rect'], {
		x: props.x ?? bounds.x,
		y: props.y ?? bounds.y,
		w: props.w ?? bounds.w,
		h,
		fill: props.fill ? { color: props.fill } : undefined,
		line: props.line ? { color: props.line.color, width: props.line.width || 1 } : undefined,
		rectRadius: props.rectRadius,
	})

	applyAnimation(ctx, ref, node.animation)

	return h
}

function renderPill(ctx: RenderContext, node: ElementNode, bounds: Bounds): number {
	const props = node.props as {
		text: string
		fill: HexColor
		color?: HexColor
		fontSize?: number
		fontFace?: string
		h?: number
		rectRadius?: number
	}

	const h = props.h || 0.65

	const ref = ctx.slide.addPill({
		x: bounds.x,
		y: bounds.y,
		w: bounds.w,
		h,
		text: props.text,
		fill: props.fill,
		color: props.color || 'FFFFFF',
		fontSize: props.fontSize || 16,
		fontFace: props.fontFace || ctx.defaults.fontFace,
		rectRadius: props.rectRadius || 0.15,
		animation: node.animation ? mapAnimationToOldAPI(node.animation) : undefined,
	})

	// Animation handled internally by addPill when animation prop is passed

	return h
}

function renderCard(ctx: RenderContext, node: ElementNode, bounds: Bounds): number {
	const props = node.props as {
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

	const h = props.h || 2

	const ref = ctx.slide.addCard({
		x: props.x ?? bounds.x,
		y: props.y ?? bounds.y,
		w: props.w ?? bounds.w,
		h,
		heading: props.heading,
		body: props.body,
		fill: props.fill,
		headingColor: props.headingColor,
		bodyColor: props.bodyColor,
		shadow: props.shadow,
	})

	applyAnimation(ctx, ref, node.animation)

	return h
}

function renderImage(ctx: RenderContext, node: ElementNode, bounds: Bounds): number {
	const props = node.props as {
		path?: string
		data?: string
		x?: number
		y?: number
		w?: number
		h?: number
		sizing?: { type: 'cover' | 'contain'; w: number; h: number }
	}

	const h = props.h || 3

	const ref = ctx.slide.addImage({
		path: props.path,
		data: props.data,
		x: props.x ?? bounds.x,
		y: props.y ?? bounds.y,
		w: props.w ?? bounds.w,
		h,
		sizing: props.sizing,
	})

	applyAnimation(ctx, ref, node.animation)

	return h
}

function renderNumberedList(ctx: RenderContext, node: ElementNode, bounds: Bounds): number {
	const props = node.props as {
		items: string[]
		x?: number
		y?: number
		w?: number
		fontSize?: number
		fontFace?: string
		color?: HexColor
		accentColor?: HexColor
		itemGap?: number
	}

	const fontSize = props.fontSize || 18
	const itemGap = props.itemGap || 0.4
	const itemHeight = (fontSize / 72) * 1.5
	const totalHeight = props.items.length * itemHeight + (props.items.length - 1) * itemGap

	ctx.slide.addNumberedList({
		x: props.x ?? bounds.x,
		y: props.y ?? bounds.y,
		w: props.w ?? bounds.w,
		fontSize,
		fontFace: props.fontFace || ctx.defaults.fontFace,
		color: props.color || ctx.defaults.color,
		accentColor: props.accentColor || ctx.defaults.accentColor,
		itemGap,
		items: props.items,
	})

	// TODO: Animation for list items

	return totalHeight
}

function renderBulletList(ctx: RenderContext, node: ElementNode, bounds: Bounds): number {
	const props = node.props as {
		items: string[]
		x?: number
		y?: number
		w?: number
		fontSize?: number
		fontFace?: string
		color?: HexColor
		accentColor?: HexColor
		itemGap?: number
	}

	const fontSize = props.fontSize || 18
	const itemGap = props.itemGap || 0.4
	const itemHeight = (fontSize / 72) * 1.5
	const totalHeight = props.items.length * itemHeight + (props.items.length - 1) * itemGap

	// Use bullet list (we'd need to add this to old API or use addText with bullet)
	for (let i = 0; i < props.items.length; i++) {
		ctx.slide.addText(`\u2022  ${props.items[i]}`, {
			x: props.x ?? bounds.x,
			y: (props.y ?? bounds.y) + i * (itemHeight + itemGap),
			w: props.w ?? bounds.w,
			h: itemHeight,
			fontSize,
			fontFace: props.fontFace || ctx.defaults.fontFace,
			color: props.color || ctx.defaults.color,
		})
	}

	return totalHeight
}

// ============================================================================
// LAYOUT RENDERERS
// ============================================================================

function renderStack(ctx: RenderContext, node: ElementNode, bounds: Bounds): number {
	const props = node.props as {
		gap?: number
		x?: number
		y?: number
		w?: number
	}

	const gap = props.gap ?? 0.2
	const x = props.x ?? bounds.x
	const w = props.w ?? bounds.w
	let y = props.y ?? bounds.y

	const stagger = node.animation?.stagger

	for (let childIndex = 0; childIndex < node.children.length; childIndex++) {
		const child = node.children[childIndex]
		let childWithAnimation = child

		if (node.animation && stagger) {
			// Stagger mode: first child onClick, rest afterPrevious
			if (childIndex === 0) {
				childWithAnimation = {
					...child,
					animation: {
						...node.animation,
						trigger: 'onClick' as const,
						delayMs: 0,
						stagger: undefined, // Don't pass stagger down
					},
				}
			} else {
				childWithAnimation = {
					...child,
					animation: {
						...node.animation,
						trigger: 'afterPrevious' as const,
						delayMs: stagger,
						stagger: undefined,
					},
				}
			}
		} else if (node.animation) {
			// No stagger: all children get same animation (if they don't have their own)
			if (!child.animation) {
				childWithAnimation = {
					...child,
					animation: node.animation,
				}
			}
		}
		// If no parent animation, children keep their own (or none)

		const childHeight = renderElement(ctx, childWithAnimation, { x, y, w, h: bounds.h })
		y += childHeight + gap
	}

	return y - (props.y ?? bounds.y) - gap // Total height minus last gap
}

function renderColumns(ctx: RenderContext, node: ElementNode, bounds: Bounds): number {
	const props = node.props as {
		gap?: number
		x?: number
		y?: number
		w?: number
		h?: number
		ratio?: number[]
	}

	const gap = props.gap ?? 0.3
	const x = props.x ?? bounds.x
	const y = props.y ?? bounds.y
	const w = props.w ?? bounds.w
	const h = props.h ?? bounds.h

	const numCols = node.children.length
	const ratio = props.ratio || node.children.map(() => 1)
	const totalRatio = ratio.reduce((a, b) => a + b, 0)
	const totalGapWidth = gap * (numCols - 1)
	const availableWidth = w - totalGapWidth

	let currentX = x
	let maxHeight = 0

	for (let i = 0; i < node.children.length; i++) {
		const colWidth = (ratio[i] / totalRatio) * availableWidth
		const child = node.children[i]

		// Inherit parent animation if child doesn't have one
		const childWithAnimation = child.animation ? child : (node.animation ? { ...child, animation: node.animation } : child)

		const childHeight = renderElement(ctx, childWithAnimation, { x: currentX, y, w: colWidth, h })
		maxHeight = Math.max(maxHeight, childHeight)
		currentX += colWidth + gap
	}

	return maxHeight
}

function renderGrid(ctx: RenderContext, node: ElementNode, bounds: Bounds): number {
	const props = node.props as {
		cols: number
		gap?: number
		x?: number
		y?: number
		w?: number
		h?: number
	}

	const cols = props.cols
	const gap = props.gap ?? 0.2
	const x = props.x ?? bounds.x
	const y = props.y ?? bounds.y
	const w = props.w ?? bounds.w

	const cellWidth = (w - gap * (cols - 1)) / cols
	const rows = Math.ceil(node.children.length / cols)

	let maxRowHeight = 0
	let totalHeight = 0

	for (let i = 0; i < node.children.length; i++) {
		const row = Math.floor(i / cols)
		const col = i % cols

		const cellX = x + col * (cellWidth + gap)
		const cellY = y + totalHeight

		if (col === 0 && i > 0) {
			totalHeight += maxRowHeight + gap
			maxRowHeight = 0
		}

		const child = node.children[i]
		const childHeight = renderElement(ctx, child, { x: cellX, y: cellY, w: cellWidth, h: bounds.h })
		maxRowHeight = Math.max(maxRowHeight, childHeight)
	}

	return totalHeight + maxRowHeight
}

function renderAbsolute(ctx: RenderContext, node: ElementNode): number {
	const props = node.props as { x: number; y: number; w?: number; h?: number }

	if (node.children.length > 0) {
		renderElement(ctx, node.children[0], {
			x: props.x,
			y: props.y,
			w: props.w ?? 4,
			h: props.h ?? 2,
		})
	}

	return 0 // Absolute positioned elements don't affect flow
}

// ============================================================================
// ANIMATION HELPERS
// ============================================================================

function applyAnimation(ctx: RenderContext, ref: { _shapeIndex: number }, animation?: AnimationConfig): void {
	if (!animation) return

	ctx.slide.addAnimation(ref, mapAnimationToOldAPI(animation))
}

function mapAnimationToOldAPI(animation: AnimationConfig): {
	type: string
	trigger: 'onClick' | 'withPrevious' | 'afterPrevious'
	direction?: string
	durationMs?: number
	delayMs?: number
} {
	const directionMap: Record<string, string> = {
		'from-left': 'from-left',
		'from-right': 'from-right',
		'from-top': 'from-top',
		'from-bottom': 'from-bottom',
	}

	return {
		type: animation.type,
		trigger: animation.trigger,
		direction: animation.direction ? directionMap[animation.direction] : undefined,
		durationMs: animation.durationMs,
		delayMs: animation.delayMs,
	}
}
