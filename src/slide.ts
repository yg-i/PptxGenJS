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
			line: { color: config.borderColor, width: config.borderWidth },
			rectRadius: config.borderRadius,
			shadow: config.shadow,
		})

		// Store reference to the background shape
		const backgroundShapeRef = this._createShapeRef()

		// Add heading text if provided
		if (config.heading) {
			this.addText(config.heading, {
				x: config.headingX,
				y: config.headingY,
				w: config.headingW,
				h: config.headingFontSize / 72 * 1.5, // Approximate height
				fontSize: config.headingFontSize,
				fontFace: config.headingFontFace,
				bold: config.headingBold,
				color: config.headingColor,
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
				valign: 'top',
			})
		}

		return backgroundShapeRef
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
