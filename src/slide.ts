/**
 * PptxGenJS: Slide Class
 */

import { ANIMATION_DIRECTIONS, ANIMATION_PRESETS, CHART_NAME, SHAPE_NAME, ShapeType } from './core-enums'
import {
	AddSlideProps,
	AnimationProps,
	BackgroundProps,
	HexColor,
	IChartMulti,
	IChartOpts,
	IChartOptsLib,
	IOptsChartData,
	ISlideAnimation,
	ISlideObject,
	ISlideRel,
	ISlideRelChart,
	ISlideRelMedia,
	ImageProps,
	MediaProps,
	PresLayout,
	PresSlide,
	ShapeProps,
	ShapeRef,
	SlideLayout,
	SlideNumberProps,
	TableProps,
	TableRow,
	TextProps,
	TextPropsOptions,
	TransitionProps,
} from './core-interfaces'
import * as genObj from './gen-objects'
import { calculateGridLayout } from './layout'
import type { GridLayoutOptions, GapValue } from './layout'
import { resolveCardConfig } from './components'
import type { CardOptions } from './components'
import { interpolateColors } from './utils'
import { convertRichTextToTextProps, isStyledTextFragment } from './rich-text'
import type { StyledTextFragment, RichTextOptions } from './rich-text'

/**
 * Options for addTitle convenience method.
 */
export interface TitleOptions {
	/** X position (default: 0.5) */
	x?: number
	/** Y position (default: 0.4) */
	y?: number
	/** Width (default: 9) */
	w?: number
	/** Height (default: 0.7) */
	h?: number
	/** Font size (default: 32) */
	fontSize?: number
	/** Font face (default: 'Arial') */
	fontFace?: string
	/** Text color (hex string) */
	color?: HexColor
	/** Bold text (default: true) */
	bold?: boolean
	/**
	 * Gradient colors for the title text.
	 * Specify { from, to } to create a gradient effect across the text.
	 */
	gradient?: { from: HexColor; to: HexColor }
}

/**
 * Options for addGradientText method.
 */
export interface GradientTextOptions extends Omit<TextPropsOptions, 'color'> {
	/** Starting color of the gradient */
	gradientFrom: HexColor
	/** Ending color of the gradient */
	gradientTo: HexColor
	/**
	 * Granularity of the gradient.
	 * - 'word': Each word gets a different color (default)
	 * - 'character': Each character gets a different color (smoother but more objects)
	 */
	gradientMode?: 'word' | 'character'
}

/**
 * Options for addBadge method.
 */
export interface BadgeOptions {
	/** X position (inches) */
	x: number
	/** Y position (inches) */
	y: number
	/** Size of the badge (inches). Badge is always circular. */
	size: number
	/** Text to display inside the badge */
	text: string
	/** Background color of the badge */
	color: HexColor
	/** Text color (default: 'FFFFFF') */
	textColor?: HexColor
	/** Font size (default: auto-calculated based on size) */
	fontSize?: number
	/** Font face (default: 'Arial') */
	fontFace?: string
	/** Bold text (default: true) */
	bold?: boolean
}

/**
 * A single item in a bullet list.
 */
export interface BulletListItem {
	/** Text content of the list item */
	text: string
	/** Optional colored badge before the text */
	badge?: {
		/** Text inside the badge (e.g., '1', '2', 'A') */
		text: string
		/** Background color of the badge */
		color: HexColor
		/** Text color inside the badge (default: 'FFFFFF') */
		textColor?: HexColor
	}
	/** Text color for this item (overrides list default) */
	color?: HexColor
	/** Bold text for this item */
	bold?: boolean
}

/**
 * Options for addBulletList method.
 */
export interface BulletListOptions {
	/** X position (inches) */
	x: number
	/** Y position (inches) */
	y: number
	/** Width (inches) */
	w: number
	/** Height per item (inches, default: 0.5) */
	itemHeight?: number
	/** List items */
	items: BulletListItem[]
	/** Show bullet points before items (default: true) */
	showBullets?: boolean
	/** Bullet character (default: '•') */
	bulletChar?: string
	/** Text color (default: '555555') */
	color?: HexColor
	/** Font size (default: 16) */
	fontSize?: number
	/** Font face (default: 'Arial') */
	fontFace?: string
	/** Badge size (default: 0.3) */
	badgeSize?: number
	/** Gap between bullet and badge (default: 0.1) */
	bulletBadgeGap?: number
	/** Gap between badge and text (default: 0.15) */
	badgeTextGap?: number
}

/**
 * Options for addTwoColumn layout method.
 */
export interface TwoColumnOptions {
	/** X position (inches) */
	x: number
	/** Y position (inches) */
	y: number
	/** Total width (inches) */
	w: number
	/** Total height (inches) */
	h: number
	/** Gap between columns (default: 0.5) */
	gap?: number
	/** Left column configuration */
	left?: {
		/** Width of left column (alternative to ratio) */
		w?: number
		/** Width ratio (0-1, e.g., 0.45 = 45% of total width minus gap) */
		ratio?: number
	}
	/** Right column configuration */
	right?: {
		/** Width of right column (alternative to ratio) */
		w?: number
		/** Width ratio (0-1, e.g., 0.55 = 55% of total width minus gap) */
		ratio?: number
	}
	/** Render function for left column */
	renderLeft: (bounds: { x: number; y: number; w: number; h: number }) => void
	/** Render function for right column */
	renderRight: (bounds: { x: number; y: number; w: number; h: number }) => void
}

/**
 * Options for addStack vertical layout.
 */
export interface StackOptions {
	/** X position (inches) */
	x: number
	/** Y position (inches) */
	y: number
	/** Width (inches) */
	w: number
	/** Gap between items (inches, default: 0.2) */
	gap?: number
	/** Default text options applied to all items */
	defaults?: TextPropsOptions
}

/**
 * Item options that include height specification.
 */
export interface StackItemOptions extends TextPropsOptions {
	/** Height of this item (inches). Required for proper stacking. */
	h: number
}

/**
 * Builder object passed to addStack callback for adding stacked items.
 */
export interface StackBuilder {
	/** Current Y position for the next item */
	readonly currentY: number

	/**
	 * Add text to the stack.
	 * @param text - Text content
	 * @param options - Text options (h is required)
	 */
	text(text: string | TextProps[], options: StackItemOptions): ShapeRef

	/**
	 * Add rich text with tagged template to the stack.
	 * @param options - Text options (h is required)
	 * @returns Tagged template function
	 */
	richText(options: StackItemOptions): (strings: TemplateStringsArray, ...values: (string | StyledTextFragment)[]) => ShapeRef

	/**
	 * Add a spacer to the stack.
	 * @param height - Height of the spacer (inches)
	 */
	space(height: number): void

	/**
	 * Add a card to the stack.
	 * @param options - Card options (h is required, x/y/w are ignored)
	 */
	card(options: Omit<CardOptions, 'x' | 'y' | 'w'> & { h: number }): ShapeRef
}

export default class Slide {
	private readonly _setSlideNum: (value: SlideNumberProps) => void

	public addSlide: (options?: AddSlideProps) => PresSlide
	public getSlide: (slideNum: number) => PresSlide
	public _name: string
	public _presLayout: PresLayout
	public _rels: ISlideRel[]
	public _relsChart: ISlideRelChart[]
	public _relsMedia: ISlideRelMedia[]
	public _rId: number
	public _slideId: number
	public _slideLayout: SlideLayout | undefined
	public _slideNum: number
	public _slideNumberProps: SlideNumberProps | undefined
	public _slideObjects: ISlideObject[]
	public _newAutoPagedSlides: PresSlide[] = []
	public _transition: TransitionProps | undefined
	public _animations: ISlideAnimation[]

	constructor(params: {
		addSlide: (options?: AddSlideProps) => PresSlide
		getSlide: (slideNum: number) => PresSlide
		presLayout: PresLayout
		setSlideNum: (value: SlideNumberProps) => void
		slideId: number
		slideRId: number
		slideNumber: number
		slideLayout?: SlideLayout
	}) {
		this.addSlide = params.addSlide
		this.getSlide = params.getSlide
		this._name = `Slide ${params.slideNumber}`
		this._presLayout = params.presLayout
		this._rId = params.slideRId
		this._rels = []
		this._relsChart = []
		this._relsMedia = []
		this._setSlideNum = params.setSlideNum
		this._slideId = params.slideId
		this._slideLayout = params.slideLayout
		this._slideNum = params.slideNumber
		this._slideObjects = []
		this._animations = []
		/** NOTE: Slide Numbers: In order for Slide Numbers to function they need to be in all 3 files: master/layout/slide
		 * `defineSlideMaster` and `addNewSlide.slideNumber` will add {slideNumber} to `this.masterSlide` and `this.slideLayouts`
		 * so, lastly, add to the Slide now.
		 */
		this._slideNumberProps = this._slideLayout?._slideNumberProps
	}

	/**
	 * Background color or image
	 * @type {BackgroundProps}
	 * @example solid color `background: { color:'FF0000' }`
	 * @example color+trans `background: { color:'FF0000', transparency:0.5 }`
	 * @example base64 `background: { data:'image/png;base64,ABC[...]123' }`
	 * @example url `background: { path:'https://some.url/image.jpg'}`
	 * @since v3.3.0
	 */
	private _background: BackgroundProps | undefined
	public set background(props: BackgroundProps) {
		this._background = props
		// Add background (image data/path must be captured before `exportPresentation()` is called)
		if (props) genObj.addBackgroundDefinition(props, this)
	}

	public get background(): BackgroundProps | undefined {
		return this._background
	}

	/**
	 * Default font color
	 * @type {HexColor}
	 */
	private _color: HexColor | undefined
	public set color(value: HexColor) {
		this._color = value
	}

	public get color(): HexColor | undefined {
		return this._color
	}

	/**
	 * @type {boolean}
	 */
	private _hidden: boolean = false
	public set hidden(value: boolean) {
		this._hidden = value
	}

	public get hidden(): boolean {
		return this._hidden
	}

	/**
	 * @type {SlideNumberProps}
	 */
	public set slideNumber(value: SlideNumberProps) {
		// NOTE: Slide Numbers: In order for Slide Numbers to function they need to be in all 3 files: master/layout/slide
		this._slideNumberProps = value
		this._setSlideNum(value)
	}

	public get slideNumber(): SlideNumberProps | undefined {
		return this._slideNumberProps
	}

	/**
	 * Slide transition
	 * @since v4.1.0
	 * @example slide.transition = { type: 'fade' }
	 * @example slide.transition = { type: 'morph', durationMs: 2000 }
	 * @example slide.transition = { type: 'push', direction: 'l', speed: 'slow' }
	 */
	public set transition(value: TransitionProps) {
		this._transition = value
	}

	public get transition(): TransitionProps | undefined {
		return this._transition
	}

	public get newAutoPagedSlides(): PresSlide[] {
		return this._newAutoPagedSlides
	}

	/**
	 * Add chart to Slide
	 * @param {CHART_NAME|IChartMulti[]} type - chart type
	 * @param {object[]} data - data object
	 * @param {IChartOpts} options - chart options
	 * @return {ShapeRef} reference to the added chart for animation targeting
	 * @since v4.2.0 - returns ShapeRef instead of Slide
	 */
	addChart(type: CHART_NAME | IChartMulti[], data: IOptsChartData[], options?: IChartOpts): ShapeRef {
		// FUTURE: TODO-VERSION-4: Remove first arg - only take data and opts, with "type" required on opts
		// Set `_type` on IChartOptsLib as its what is used as object is passed around
		const optionsWithType: IChartOptsLib = options || {}
		optionsWithType._type = type
		genObj.addChartDefinition(this, type, data, optionsWithType)
		return this._createShapeRef()
	}

	/**
	 * Add image to Slide
	 * @param {ImageProps} options - image options
	 * @return {ShapeRef} reference to the added image for animation targeting
	 * @since v4.2.0 - returns ShapeRef instead of Slide
	 */
	addImage(options: ImageProps): ShapeRef {
		genObj.addImageDefinition(this, options)
		return this._createShapeRef()
	}

	/**
	 * Add media (audio/video) to Slide
	 * @param {MediaProps} options - media options
	 * @return {Slide} this Slide
	 */
	addMedia(options: MediaProps): Slide {
		genObj.addMediaDefinition(this, options)
		return this
	}

	/**
	 * Add speaker notes to Slide
	 * @docs https://gitbrent.github.io/PptxGenJS/docs/speaker-notes.html
	 * @param {string} notes - notes to add to slide
	 * @return {Slide} this Slide
	 */
	addNotes(notes: string): Slide {
		genObj.addNotesDefinition(this, notes)
		return this
	}

	/**
	 * Add shape to Slide
	 * @param {SHAPE_NAME} shapeName - shape name
	 * @param {ShapeProps} options - shape options
	 * @return {ShapeRef} reference to the added shape for animation targeting
	 * @since v4.2.0 - returns ShapeRef instead of Slide
	 */
	addShape(shapeName: SHAPE_NAME, options?: ShapeProps): ShapeRef {
		// NOTE: As of v3.1.0, <script> users are passing the old shape object from the shapes file (orig to the project)
		// But React/TypeScript users are passing the shapeName from an enum, which is a simple string, so lets cast
		// <script./> => `pptx.shapes.RECTANGLE` [string] "rect" ... shapeName['name'] = 'rect'
		// TypeScript => `pptxgen.shapes.RECTANGLE` [string] "rect" ... shapeName = 'rect'
		// let shapeNameDecode = typeof shapeName === 'object' && shapeName['name'] ? shapeName['name'] : shapeName
		genObj.addShapeDefinition(this, shapeName, options || {})
		return this._createShapeRef()
	}

	/**
	 * Add table to Slide
	 * @param {TableRow[]} tableRows - table rows
	 * @param {TableProps} options - table options
	 * @return {Slide} this Slide
	 */
	addTable(tableRows: TableRow[], options?: TableProps): Slide {
		// FUTURE: we pass `this` - we dont need to pass layouts - they can be read from this!
		this._newAutoPagedSlides = genObj.addTableDefinition(this, tableRows, options || {}, this._slideLayout, this._presLayout, this.addSlide, this.getSlide)
		return this
	}

	/**
	 * Add text to Slide
	 * @param {string|TextProps[]} text - text string or complex object
	 * @param {TextPropsOptions} options - text options
	 * @return {ShapeRef} reference to the added text for animation targeting
	 * @since v4.2.0 - returns ShapeRef instead of Slide
	 */
	addText(text: string | TextProps[], options?: TextPropsOptions): ShapeRef {
		const textParam = typeof text === 'string' || typeof text === 'number' ? [{ text, options }] : text
		genObj.addTextDefinition(this, textParam, options || {}, false)
		return this._createShapeRef()
	}

	/**
	 * Add rich text with inline styling using tagged template literals.
	 *
	 * @example
	 * ```typescript
	 * import { textStyle } from 'pptxgenjs'
	 *
	 * const keyword = textStyle({ bold: true, color: '5DADE2' })
	 * const emphasis = textStyle({ italic: true })
	 *
	 * slide.addRichText`By ${keyword('committing')} to a choice, the agent is ${keyword('justified')} in being in it.`
	 *
	 * // With positioning options:
	 * slide.addRichText({ x: 1, y: 2, w: 8, fontSize: 18, color: 'FFFFFF' })`Hello ${keyword('world')}!`
	 * ```
	 *
	 * @param options - Text positioning and default styling options
	 * @returns Tagged template function
	 * @since v5.0.0
	 */
	addRichText(options: RichTextOptions): (strings: TemplateStringsArray, ...values: (string | StyledTextFragment)[]) => ShapeRef
	addRichText(strings: TemplateStringsArray, ...values: (string | StyledTextFragment)[]): ShapeRef
	addRichText(
		optionsOrStrings: RichTextOptions | TemplateStringsArray,
		...values: (string | StyledTextFragment)[]
	): ShapeRef | ((strings: TemplateStringsArray, ...values: (string | StyledTextFragment)[]) => ShapeRef) {
		// Check if called as tagged template directly (no options)
		if (Array.isArray(optionsOrStrings) && 'raw' in optionsOrStrings) {
			const strings = optionsOrStrings as TemplateStringsArray
			const textProps = convertRichTextToTextProps(strings, values, {})
			genObj.addTextDefinition(this, textProps, {}, false)
			return this._createShapeRef()
		}

		// Called with options first - return a tagged template function
		const options = optionsOrStrings as RichTextOptions
		return (strings: TemplateStringsArray, ...templateValues: (string | StyledTextFragment)[]): ShapeRef => {
			const textProps = convertRichTextToTextProps(strings, templateValues, options)
			genObj.addTextDefinition(this, textProps, options, false)
			return this._createShapeRef()
		}
	}

	/**
	 * Add animation to a shape on this slide
	 * @since v4.1.0
	 * @since v4.2.0 - accepts ShapeRef in addition to numeric index
	 * @param {ShapeRef|number} shapeOrIndex - ShapeRef returned by addShape/addText/addImage, or numeric index (0-based)
	 * @param {AnimationProps} options - animation options
	 * @return {Slide} this Slide
	 * @example slide.addAnimation(shape, { type: 'fade' }) // using ShapeRef (recommended)
	 * @example slide.addAnimation(0, { type: 'fade' }) // using numeric index
	 */
	addAnimation(shapeOrIndex: ShapeRef | number, options: AnimationProps): Slide {
		// Resolve shape index from ShapeRef or number
		let shapeIndex: number
		if (typeof shapeOrIndex === 'number') {
			shapeIndex = shapeOrIndex
		} else if (shapeOrIndex && typeof shapeOrIndex === 'object' && '_shapeIndex' in shapeOrIndex) {
			// Validate ShapeRef belongs to this slide
			if (shapeOrIndex._slideRef !== this) {
				console.warn('PptxGenJS: addAnimation - ShapeRef belongs to a different slide')
				return this
			}
			shapeIndex = shapeOrIndex._shapeIndex
		} else {
			console.warn('PptxGenJS: addAnimation - invalid shapeOrIndex parameter')
			return this
		}

		// Validate shape index
		if (shapeIndex < 0 || shapeIndex >= this._slideObjects.length) {
			console.warn(`PptxGenJS: addAnimation - invalid shapeIndex ${shapeIndex}. Slide has ${this._slideObjects.length} shapes.`)
			return this
		}

		// Look up animation preset
		const preset = ANIMATION_PRESETS[options.type]
		if (!preset) {
			console.warn(`PptxGenJS: addAnimation - unknown animation type '${options.type}'`)
			return this
		}

		// Resolve direction subtype
		let presetSubtype: number | undefined
		if (options.direction && ANIMATION_DIRECTIONS[options.direction]) {
			presetSubtype = ANIMATION_DIRECTIONS[options.direction]
		}

		// Create animation object
		const animation: ISlideAnimation = {
			shapeIndex,
			options,
			presetId: preset.presetId,
			presetClass: preset.presetClass,
			presetSubtype,
		}

		this._animations.push(animation)
		return this
	}

	/**
	 * Create a ShapeRef for the most recently added shape
	 * @internal
	 */
	private _createShapeRef(): ShapeRef {
		return {
			_shapeIndex: this._slideObjects.length - 1,
			_slideRef: this as unknown as PresSlide,
		}
	}

	// ============================================================================
	// COMPOSITIONAL API - High-level components and layouts
	// ============================================================================

	/**
	 * Add a title to the slide with sensible defaults.
	 * Convenience method that wraps addText with common title styling.
	 *
	 * @since v5.0.0
	 * @param text - Title text
	 * @param options - Optional title configuration
	 * @returns ShapeRef to the title text
	 *
	 * @example
	 * slide.addTitle('My Presentation')
	 *
	 * @example // With custom color
	 * slide.addTitle('My Presentation', { color: '2A9D8F' })
	 *
	 * @example // With gradient
	 * slide.addTitle('Gradient Title', { gradient: { from: '1E88E5', to: '26A69A' } })
	 */
	addTitle(text: string, options?: TitleOptions): ShapeRef {
		const titleDefaults = {
			x: 0.5,
			y: 0.4,
			w: 9,
			h: 0.7,
			fontSize: 32,
			fontFace: 'Arial',
			bold: true,
		}

		const config = { ...titleDefaults, ...options }

		// If gradient is specified, use addGradientText
		if (config.gradient) {
			return this.addGradientText(text, {
				x: config.x,
				y: config.y,
				w: config.w,
				h: config.h,
				fontSize: config.fontSize,
				fontFace: config.fontFace,
				bold: config.bold,
				gradientFrom: config.gradient.from,
				gradientTo: config.gradient.to,
				gradientMode: 'word',
			})
		}

		// Otherwise, use regular addText
		return this.addText(text, {
			x: config.x,
			y: config.y,
			w: config.w,
			h: config.h,
			fontSize: config.fontSize,
			fontFace: config.fontFace,
			bold: config.bold,
			color: config.color,
		})
	}

	/**
	 * Add text with a gradient color effect.
	 * Creates a visual gradient by splitting the text into segments with interpolated colors.
	 *
	 * @since v5.0.0
	 * @param text - Text to display
	 * @param options - Gradient and text options
	 * @returns ShapeRef to the text
	 *
	 * @example
	 * slide.addGradientText('Hello World', {
	 *   x: 1, y: 1, w: 6, h: 0.5,
	 *   gradientFrom: '1E88E5',
	 *   gradientTo: '26A69A',
	 *   fontSize: 24,
	 *   bold: true,
	 * })
	 */
	addGradientText(text: string, options: GradientTextOptions): ShapeRef {
		const { gradientFrom, gradientTo, gradientMode = 'word', ...textOptions } = options

		// Split text into segments
		const segments = gradientMode === 'character'
			? text.split('')
			: text.split(/(\s+)/) // Split by whitespace, keeping separators

		// Filter out empty segments
		const nonEmptySegments = segments.filter(s => s.length > 0)

		// Generate colors for each segment
		const colors = interpolateColors(gradientFrom, gradientTo, nonEmptySegments.length)

		// Create TextProps array with colors
		const textProps: TextProps[] = nonEmptySegments.map((segment, index) => ({
			text: segment,
			options: {
				color: colors[index],
				fontSize: textOptions.fontSize,
				fontFace: textOptions.fontFace,
				bold: textOptions.bold,
				italic: textOptions.italic,
				underline: textOptions.underline,
			},
		}))

		return this.addText(textProps, {
			x: textOptions.x,
			y: textOptions.y,
			w: textOptions.w,
			h: textOptions.h,
			valign: textOptions.valign,
			align: textOptions.align,
		})
	}

	/**
	 * Add a card component to the slide.
	 * A card is a rounded rectangle with optional shadow, heading, and body text.
	 *
	 * @since v5.0.0
	 * @param options - Card configuration
	 * @returns ShapeRef to the card's background shape
	 *
	 * @example
	 * slide.addCard({
	 *   x: 0.5, y: 1.0, w: 4, h: 2,
	 *   heading: '1. LEARNING',
	 *   headingColor: 'C5A636',
	 *   body: 'How machines acquire knowledge from data.',
	 *   shadow: 'sm',
	 * })
	 *
	 * @example // Card without border
	 * slide.addCard({ ..., border: false })
	 *
	 * @example // Highlighted card
	 * slide.addCard({ ..., highlight: true })
	 * slide.addCard({ ..., highlight: 'E3F2FD' }) // Custom highlight color
	 */
	addCard(options: CardOptions): ShapeRef {
		const config = resolveCardConfig(options)

		// Add background shape (rounded rectangle)
		this.addShape(ShapeType.roundRect, {
			x: config.x,
			y: config.y,
			w: config.w,
			h: config.h,
			fill: config.backgroundFill,
			line: config.hasBorder
				? { color: config.borderColor, width: config.borderWidth }
				: { color: 'FFFFFF', width: 0 }, // No visible border
			rectRadius: config.borderRadius,
			shadow: config.shadow,
		})

		// Store reference to the background shape
		const backgroundShapeRef = this._createShapeRef()

		// Add title text if provided (small text above heading)
		if (config.title) {
			this.addText(config.title, {
				x: config.titleX,
				y: config.titleY,
				w: config.titleW,
				h: config.titleH,
				fontSize: config.titleFontSize,
				fontFace: config.titleFontFace,
				color: config.titleColor,
				align: config.align,
			})
		}

		// Add heading text if provided (main prominent text)
		if (config.heading) {
			this.addText(config.heading, {
				x: config.headingX,
				y: config.headingY,
				w: config.headingW,
				h: config.headingH,
				fontSize: config.headingFontSize,
				fontFace: config.headingFontFace,
				bold: config.headingBold,
				color: config.headingColor,
				align: config.align,
			})
		}

		// Add body text if provided
		if (config.body) {
			this.addText(config.body, {
				x: config.bodyX,
				y: config.bodyY,
				w: config.bodyW,
				h: config.bodyH,
				fontSize: config.bodyFontSize,
				fontFace: config.bodyFontFace,
				color: config.bodyColor,
				italic: config.bodyItalic,
				align: config.align,
				valign: 'top',
			})
		}

		return backgroundShapeRef
	}

	/**
	 * Add a circular badge with text inside (e.g., numbered circle).
	 *
	 * @since v5.0.0
	 * @param options - Badge configuration
	 * @returns ShapeRef to the badge
	 *
	 * @example
	 * slide.addBadge({
	 *   x: 1, y: 1, size: 0.3,
	 *   text: '1',
	 *   color: '29B6F6',
	 * })
	 */
	addBadge(options: BadgeOptions): ShapeRef {
		const {
			x, y, size, text, color,
			textColor = 'FFFFFF',
			fontSize = Math.round(size * 72 * 0.5), // Auto-size based on badge size
			fontFace = 'Arial',
			bold = true,
		} = options

		// Add circle
		this.addShape(ShapeType.ellipse, {
			x,
			y,
			w: size,
			h: size,
			fill: { color },
			line: { color, width: 0 },
		})

		const shapeRef = this._createShapeRef()

		// Add centered text
		this.addText(text, {
			x,
			y,
			w: size,
			h: size,
			fontSize,
			fontFace,
			color: textColor,
			bold,
			align: 'center',
			valign: 'middle',
		})

		return shapeRef
	}

	/**
	 * Add a bullet list with optional colored badges.
	 *
	 * @since v5.0.0
	 * @param options - Bullet list configuration
	 * @returns This slide for chaining
	 *
	 * @example
	 * slide.addBulletList({
	 *   x: 0.7, y: 2.1, w: 4,
	 *   items: [
	 *     { badge: { text: '1', color: '29B6F6' }, text: 'First item' },
	 *     { badge: { text: '2', color: 'EF5350' }, text: 'Second item' },
	 *     { badge: { text: '3', color: '66BB6A' }, text: 'Third item' },
	 *   ],
	 * })
	 */
	addBulletList(options: BulletListOptions): Slide {
		const {
			x, y, w,
			items,
			itemHeight = 0.5,
			showBullets = true,
			bulletChar = '•',
			color = '555555',
			fontSize = 16,
			fontFace = 'Arial',
			badgeSize = 0.3,
			bulletBadgeGap = 0.1,
			badgeTextGap = 0.15,
		} = options

		// Calculate positions
		const bulletX = x
		const bulletW = showBullets ? 0.25 : 0
		const badgeX = bulletX + bulletW + bulletBadgeGap
		const textX = badgeX + badgeSize + badgeTextGap
		const textW = w - (textX - x)

		for (let i = 0; i < items.length; i++) {
			const item = items[i]
			const itemY = y + i * itemHeight

			// Add bullet if enabled
			if (showBullets) {
				this.addText(bulletChar, {
					x: bulletX,
					y: itemY,
					w: bulletW,
					h: itemHeight,
					fontSize,
					color,
					fontFace,
					valign: 'middle',
				})
			}

			// Add badge if provided
			if (item.badge) {
				const badgeY = itemY + (itemHeight - badgeSize) / 2
				this.addBadge({
					x: badgeX,
					y: badgeY,
					size: badgeSize,
					text: item.badge.text,
					color: item.badge.color,
					textColor: item.badge.textColor,
					fontSize: Math.round(badgeSize * 72 * 0.45),
				})
			}

			// Add text
			this.addText(item.text, {
				x: textX,
				y: itemY,
				w: textW,
				h: itemHeight,
				fontSize,
				fontFace,
				color: item.color ?? color,
				bold: item.bold,
				valign: 'middle',
			})
		}

		return this
	}

	/**
	 * Add a two-column layout with render functions for each column.
	 *
	 * @since v5.0.0
	 * @param options - Two-column layout configuration
	 * @returns This slide for chaining
	 *
	 * @example
	 * slide.addTwoColumn({
	 *   x: 0.5, y: 1, w: 9, h: 4,
	 *   gap: 0.5,
	 *   left: { ratio: 0.45 },
	 *   renderLeft: (bounds) => {
	 *     slide.addText('Left content', { ...bounds })
	 *   },
	 *   renderRight: (bounds) => {
	 *     slide.addCard({ ...bounds, heading: 'Right card' })
	 *   },
	 * })
	 */
	addTwoColumn(options: TwoColumnOptions): Slide {
		const {
			x, y, w, h,
			gap = 0.5,
			left = {},
			right = {},
			renderLeft,
			renderRight,
		} = options

		// Calculate available width after gap
		const availableWidth = w - gap

		// Determine left column width
		let leftWidth: number
		if (left.w !== undefined) {
			leftWidth = left.w
		} else if (left.ratio !== undefined) {
			leftWidth = availableWidth * left.ratio
		} else if (right.w !== undefined) {
			leftWidth = availableWidth - right.w
		} else if (right.ratio !== undefined) {
			leftWidth = availableWidth * (1 - right.ratio)
		} else {
			// Default: 50/50 split
			leftWidth = availableWidth / 2
		}

		// Right column width is the remainder
		const rightWidth = availableWidth - leftWidth

		// Calculate bounds
		const leftBounds = { x, y, w: leftWidth, h }
		const rightBounds = { x: x + leftWidth + gap, y, w: rightWidth, h }

		// Render columns
		renderLeft(leftBounds)
		renderRight(rightBounds)

		return this
	}

	/**
	 * Add vertically stacked elements with automatic Y positioning.
	 *
	 * @since v5.0.0
	 * @param options - Stack configuration (position, width, gap)
	 * @param builder - Callback function that receives a builder for adding items
	 * @returns This slide for chaining
	 *
	 * @example
	 * ```typescript
	 * const keyword = textStyle({ bold: true, color: '5DADE2' })
	 *
	 * slide.addStack({ x: 0.65, y: 1.0, w: 9, gap: 0.25 }, (add) => {
	 *   add.text('The Title', { h: 0.7, fontSize: 40, bold: true, color: 'FFFFFF' })
	 *   add.space(0.1) // extra spacing
	 *   add.richText({ h: 0.8, fontSize: 22, color: 'FFFFFF' })`First paragraph with ${keyword('emphasis')}.`
	 *   add.richText({ h: 0.8, fontSize: 22, color: '9EAAB8' })`Second paragraph.`
	 * })
	 * ```
	 */
	addStack(options: StackOptions, builder: (add: StackBuilder) => void): Slide {
		const { x, y, w, gap = 0.2, defaults = {} } = options
		let currentY = y

		const stackBuilder: StackBuilder = {
			get currentY() {
				return currentY
			},

			text: (text: string | TextProps[], itemOptions: StackItemOptions): ShapeRef => {
				const { h, ...textOptions } = itemOptions
				const mergedOptions = { ...defaults, ...textOptions, x, y: currentY, w, h }
				const ref = this.addText(text, mergedOptions)
				currentY += h + gap
				return ref
			},

			richText: (itemOptions: StackItemOptions) => {
				const { h, ...textOptions } = itemOptions
				return (strings: TemplateStringsArray, ...values: (string | StyledTextFragment)[]): ShapeRef => {
					const mergedOptions = { ...defaults, ...textOptions, x, y: currentY, w, h }
					const textProps = convertRichTextToTextProps(strings, values, mergedOptions)
					genObj.addTextDefinition(this, textProps, mergedOptions, false)
					const ref = this._createShapeRef()
					currentY += h + gap
					return ref
				}
			},

			space: (height: number): void => {
				currentY += height
			},

			card: (cardOptions: Omit<CardOptions, 'x' | 'y' | 'w'> & { h: number }): ShapeRef => {
				const { h, ...rest } = cardOptions
				const ref = this.addCard({ ...rest, x, y: currentY, w, h })
				currentY += h + gap
				return ref
			},
		}

		builder(stackBuilder)
		return this
	}

	/**
	 * Options for grid layout children.
	 * Each child can be a CardOptions (for cards) or a render function.
	 */
	// eslint-disable-next-line @typescript-eslint/no-explicit-any
	addGrid<T extends Record<string, any>>(
		options: {
			/** X position of the grid */
			x: number
			/** Y position of the grid */
			y: number
			/** Number of columns */
			cols: number
			/** Number of rows (optional - calculated from children count) */
			rows?: number
			/** Gap between cells */
			gap?: GapValue
			/** Width of each cell */
			cellWidth?: number
			/** Height of each cell */
			cellHeight?: number
			/** Total width (alternative to cellWidth) */
			w?: number
			/** Total height (alternative to cellHeight) */
			h?: number
			/** Children to place in the grid */
			children: T[]
			/** Render function to create each child. Receives child data and computed bounds. */
			render: (child: T, bounds: { x: number; y: number; w: number; h: number }, index: number) => void
		}
	): Slide {
		const { children, render, ...layoutOptions } = options

		// Calculate positions for all children
		const positions = calculateGridLayout(
			layoutOptions as GridLayoutOptions,
			children.length
		)

		// Render each child at its computed position
		for (let i = 0; i < children.length; i++) {
			render(children[i], positions[i], i)
		}

		return this
	}

	/**
	 * Convenience method to add a grid of cards.
	 * Simpler than addGrid when all children are cards.
	 *
	 * @since v5.0.0
	 * @example
	 * slide.addCardGrid({
	 *   x: 0.5, y: 1.0,
	 *   cols: 2, gap: 0.3,
	 *   cellWidth: 4, cellHeight: 1.5,
	 *   cards: [
	 *     { heading: '1. LEARNING', body: '...' },
	 *     { heading: '2. REASONING', body: '...' },
	 *   ]
	 * })
	 */
	addCardGrid(options: {
		x: number
		y: number
		cols: number
		rows?: number
		gap?: GapValue
		cellWidth?: number
		cellHeight?: number
		w?: number
		h?: number
		/** Card options without position (x, y, w, h will be set by grid) */
		cards: Array<Omit<CardOptions, 'x' | 'y' | 'w' | 'h'>>
	}): Slide {
		const { cards, ...gridOptions } = options

		return this.addGrid({
			...gridOptions,
			children: cards,
			render: (cardOptions, bounds) => {
				this.addCard({
					...cardOptions,
					x: bounds.x,
					y: bounds.y,
					w: bounds.w,
					h: bounds.h,
				})
			},
		})
	}
}
