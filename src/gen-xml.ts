/**
 * PptxGenJS: XML Generation
 * Refactored to use xmlbuilder2 for cleaner, type-safe XML generation
 */

import { create, fragment } from 'xmlbuilder2'

import {
	CRLF,
	DEF_PRES_LAYOUT_NAME,
	EMU,
	LAYOUT_IDX_SERIES_BASE,
	SLDNUMFLDID,
	SLIDE_OBJECT_TYPES,
} from './core-enums'
import type {
	IPresentationProps,
	ISlideAnimation,
	ISlideRel,
	ISlideRelChart,
	ISlideRelMedia,
	PresSlide,
	ShadowProps,
	SlideLayout,
	TransitionProps,
} from './core-interfaces'
import {
	encodeXmlEntities,
	genXmlColorSelection,
	getUuid,
} from './gen-utils'
import {
	NS_A,
	NS_P,
	NS_R,
	NS_C,
	NS_P14,
	NS_P159,
	NS_MC,
	NS_CP,
	NS_DC,
	NS_DCTERMS,
	NS_XSI,
	NS_RELATIONSHIPS,
	NS_CONTENT_TYPES,
	NS_EXTENDED_PROPERTIES,
	NS_VT,
	NS_ASVG,
	NS_MA14,
	NS_P15,
	NS_THM15,
	REL_TYPE_EXTENDED_PROPERTIES,
	REL_TYPE_CORE_PROPERTIES,
	REL_TYPE_OFFICE_DOCUMENT,
	REL_TYPE_SLIDE_MASTER,
	REL_TYPE_SLIDE,
	REL_TYPE_SLIDE_LAYOUT,
	REL_TYPE_NOTES_MASTER,
	REL_TYPE_NOTES_SLIDE,
	REL_TYPE_PRES_PROPS,
	REL_TYPE_VIEW_PROPS,
	REL_TYPE_THEME,
	REL_TYPE_TABLE_STYLES,
	REL_TYPE_HYPERLINK,
	REL_TYPE_IMAGE,
	REL_TYPE_AUDIO,
	REL_TYPE_VIDEO,
	REL_TYPE_CHART,
	REL_TYPE_MEDIA,
} from './xml'

// Re-export text functions from gen-xml-text module
export {
	genXmlTextBody,
	genXmlPlaceholder,
	genXmlParagraphProperties,
	genXmlTextRunProperties,
	genXmlTextRun,
	genXmlBodyProperties,
} from './gen-xml-text'

// Import slideObjectToXml from separate file (has @ts-nocheck due to legacy code)
import { slideObjectToXml } from './gen-xml-slide-objects'

// Import for internal use
import { genXmlParagraphProperties } from './gen-xml-text'

// Note: slideObjectToXml has been moved to gen-xml-slide-objects.ts
/**
 * Transforms slide relations to XML string.
 * Extra relations that are not dynamic can be passed using the 2nd arg (e.g. theme relation in master file).
 * These relations use rId series that starts with 1-increased maximum of rIds used for dynamic relations.
 * @param {PresSlide | SlideLayout} slide - slide object whose relations are being transformed
 * @param {{ target: string; type: string }[]} defaultRels - array of default relations
 * @return {string} XML
 */
function slideObjectRelationsToXml (slide: PresSlide | SlideLayout, defaultRels: Array<{ target: string, type: string }>): string {
	const doc = create({ version: '1.0', encoding: 'UTF-8', standalone: 'yes' })
		.ele('Relationships', { xmlns: NS_RELATIONSHIPS })

	let lastRid = 0 // stores maximum rId used for dynamic relations
	const seenTargets = new Set<string>() // track targets to detect duplicates for media

	// STEP 1: Add all rels for this Slide
	slide._rels.forEach((rel: ISlideRel) => {
		lastRid = Math.max(lastRid, rel.rId)
		if (rel.type.toLowerCase().includes('hyperlink')) {
			if (rel.data === 'slide') {
				doc.ele('Relationship', {
					Id: `rId${rel.rId}`,
					Type: REL_TYPE_SLIDE,
					Target: `slide${rel.Target}.xml`,
				}).up()
			} else {
				doc.ele('Relationship', {
					Id: `rId${rel.rId}`,
					Type: REL_TYPE_HYPERLINK,
					Target: rel.Target,
					TargetMode: 'External',
				}).up()
			}
		} else if (rel.type.toLowerCase().includes('notesSlide')) {
			doc.ele('Relationship', {
				Id: `rId${rel.rId}`,
				Target: rel.Target,
				Type: REL_TYPE_NOTES_SLIDE,
			}).up()
		}
	})

	;(slide._relsChart || []).forEach((rel: ISlideRelChart) => {
		lastRid = Math.max(lastRid, rel.rId)
		doc.ele('Relationship', {
			Id: `rId${rel.rId}`,
			Type: REL_TYPE_CHART,
			Target: rel.Target,
		}).up()
	})

	;(slide._relsMedia || []).forEach((rel: ISlideRelMedia) => {
		lastRid = Math.max(lastRid, rel.rId)
		const relTypeLower = rel.type.toLowerCase()
		const targetAlreadySeen = seenTargets.has(rel.Target)
		seenTargets.add(rel.Target)

		if (relTypeLower.includes('image')) {
			doc.ele('Relationship', {
				Id: `rId${rel.rId}`,
				Type: REL_TYPE_IMAGE,
				Target: rel.Target,
			}).up()
		} else if (relTypeLower.includes('audio')) {
			// As media has *TWO* rel entries per item, check for first one, if found add second rel with alt style
			if (targetAlreadySeen) {
				doc.ele('Relationship', {
					Id: `rId${rel.rId}`,
					Type: REL_TYPE_MEDIA,
					Target: rel.Target,
				}).up()
			} else {
				doc.ele('Relationship', {
					Id: `rId${rel.rId}`,
					Type: REL_TYPE_AUDIO,
					Target: rel.Target,
				}).up()
			}
		} else if (relTypeLower.includes('video')) {
			// As media has *TWO* rel entries per item, check for first one, if found add second rel with alt style
			if (targetAlreadySeen) {
				doc.ele('Relationship', {
					Id: `rId${rel.rId}`,
					Type: REL_TYPE_MEDIA,
					Target: rel.Target,
				}).up()
			} else {
				doc.ele('Relationship', {
					Id: `rId${rel.rId}`,
					Type: REL_TYPE_VIDEO,
					Target: rel.Target,
				}).up()
			}
		} else if (relTypeLower.includes('online')) {
			// As media has *TWO* rel entries per item, check for first one, if found add second rel with alt style
			if (targetAlreadySeen) {
				doc.ele('Relationship', {
					Id: `rId${rel.rId}`,
					Type: 'http://schemas.microsoft.com/office/2007/relationships/image',
					Target: rel.Target,
				}).up()
			} else {
				doc.ele('Relationship', {
					Id: `rId${rel.rId}`,
					Target: rel.Target,
					TargetMode: 'External',
					Type: REL_TYPE_VIDEO,
				}).up()
			}
		}
	})

	// STEP 2: Add default rels
	defaultRels.forEach((rel, idx) => {
		doc.ele('Relationship', {
			Id: `rId${lastRid + idx + 1}`,
			Type: rel.type,
			Target: rel.target,
		}).up()
	})

	return doc.end({ prettyPrint: false })
}

// Text XML generation functions moved to gen-xml-text.ts

// XML-GEN: First 6 functions create the base /ppt files

// NOTE: Text XML generation functions (genXmlParagraphProperties, genXmlTextRunProperties,
// genXmlTextRun, genXmlBodyProperties, genXmlTextBody, genXmlPlaceholder) moved to gen-xml-text.ts

// XML-GEN: First 6 functions create the base /ppt files

/**
 * Generate XML ContentType
 * @param {PresSlide[]} slides - slides
 * @param {SlideLayout[]} slideLayouts - slide layouts
 * @param {PresSlide} masterSlide - master slide
 * @returns XML
 */
export function makeXmlContTypes (slides: PresSlide[], slideLayouts: SlideLayout[], masterSlide?: PresSlide): string {
	const doc = create({ version: '1.0', encoding: 'UTF-8', standalone: 'yes' })
		.ele('Types', { xmlns: NS_CONTENT_TYPES })

	// Track added content types to avoid duplicates
	const addedContentTypes = new Set<string>()

	// Standard defaults
	doc.ele('Default', { Extension: 'xml', ContentType: 'application/xml' }).up()
	doc.ele('Default', { Extension: 'rels', ContentType: 'application/vnd.openxmlformats-package.relationships+xml' }).up()
	doc.ele('Default', { Extension: 'jpeg', ContentType: 'image/jpeg' }).up()
	doc.ele('Default', { Extension: 'jpg', ContentType: 'image/jpg' }).up()
	doc.ele('Default', { Extension: 'svg', ContentType: 'image/svg+xml' }).up()

	// STEP 1: Add standard/any media types used in Presentation
	doc.ele('Default', { Extension: 'png', ContentType: 'image/png' }).up()
	doc.ele('Default', { Extension: 'gif', ContentType: 'image/gif' }).up()
	doc.ele('Default', { Extension: 'm4v', ContentType: 'video/mp4' }).up() // NOTE: Hard-Code this extension as it wont be created in loop below (as extn !== type)
	doc.ele('Default', { Extension: 'mp4', ContentType: 'video/mp4' }).up() // NOTE: Hard-Code this extension as it wont be created in loop below (as extn !== type)

	slides.forEach(slide => {
		(slide._relsMedia || []).forEach(rel => {
			if (rel.type !== 'image' && rel.type !== 'online' && rel.type !== 'chart' && rel.extn !== 'm4v' && !addedContentTypes.has(rel.type)) {
				doc.ele('Default', { Extension: rel.extn, ContentType: rel.type }).up()
				addedContentTypes.add(rel.type)
			}
		})
	})

	doc.ele('Default', { Extension: 'vml', ContentType: 'application/vnd.openxmlformats-officedocument.vmlDrawing' }).up()
	doc.ele('Default', { Extension: 'xlsx', ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }).up()

	// STEP 2: Add presentation and slide master(s)/slide(s)
	doc.ele('Override', {
		PartName: '/ppt/presentation.xml',
		ContentType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml',
	}).up()
	doc.ele('Override', {
		PartName: '/ppt/notesMasters/notesMaster1.xml',
		ContentType: 'application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml',
	}).up()

	slides.forEach((slide, idx) => {
		doc.ele('Override', {
			PartName: `/ppt/slideMasters/slideMaster${idx + 1}.xml`,
			ContentType: 'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml',
		}).up()
		doc.ele('Override', {
			PartName: `/ppt/slides/slide${idx + 1}.xml`,
			ContentType: 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml',
		}).up()
		// Add charts if any
		slide._relsChart.forEach(rel => {
			doc.ele('Override', {
				PartName: rel.Target,
				ContentType: 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
			}).up()
		})
	})

	// STEP 3: Core PPT
	doc.ele('Override', {
		PartName: '/ppt/presProps.xml',
		ContentType: 'application/vnd.openxmlformats-officedocument.presentationml.presProps+xml',
	}).up()
	doc.ele('Override', {
		PartName: '/ppt/viewProps.xml',
		ContentType: 'application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml',
	}).up()
	doc.ele('Override', {
		PartName: '/ppt/theme/theme1.xml',
		ContentType: 'application/vnd.openxmlformats-officedocument.theme+xml',
	}).up()
	doc.ele('Override', {
		PartName: '/ppt/tableStyles.xml',
		ContentType: 'application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml',
	}).up()

	// STEP 4: Add Slide Layouts
	slideLayouts.forEach((layout, idx) => {
		doc.ele('Override', {
			PartName: `/ppt/slideLayouts/slideLayout${idx + 1}.xml`,
			ContentType: 'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml',
		}).up()
		;(layout._relsChart || []).forEach(rel => {
			doc.ele('Override', {
				PartName: rel.Target,
				ContentType: 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
			}).up()
		})
	})

	// STEP 5: Add notes slide(s)
	slides.forEach((_slide, idx) => {
		doc.ele('Override', {
			PartName: `/ppt/notesSlides/notesSlide${idx + 1}.xml`,
			ContentType: 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml',
		}).up()
	})

	// STEP 6: Add rels
	if (masterSlide) {
		masterSlide._relsChart.forEach(rel => {
			doc.ele('Override', {
				PartName: rel.Target,
				ContentType: 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
			}).up()
		})
		masterSlide._relsMedia.forEach(rel => {
			if (rel.type !== 'image' && rel.type !== 'online' && rel.type !== 'chart' && rel.extn !== 'm4v' && !addedContentTypes.has(rel.type)) {
				doc.ele('Default', { Extension: rel.extn, ContentType: rel.type }).up()
				addedContentTypes.add(rel.type)
			}
		})
	}

	// LAST: Finish XML (Resume core)
	doc.ele('Override', {
		PartName: '/docProps/core.xml',
		ContentType: 'application/vnd.openxmlformats-package.core-properties+xml',
	}).up()
	doc.ele('Override', {
		PartName: '/docProps/app.xml',
		ContentType: 'application/vnd.openxmlformats-officedocument.extended-properties+xml',
	}).up()

	return doc.end({ prettyPrint: false })
}

/**
 * Creates `_rels/.rels`
 * @returns XML
 */
export function makeXmlRootRels (): string {
	const doc = create({ version: '1.0', encoding: 'UTF-8', standalone: 'yes' })
		.ele('Relationships', { xmlns: NS_RELATIONSHIPS })
			.ele('Relationship', {
				Id: 'rId1',
				Type: REL_TYPE_EXTENDED_PROPERTIES,
				Target: 'docProps/app.xml',
			}).up()
			.ele('Relationship', {
				Id: 'rId2',
				Type: REL_TYPE_CORE_PROPERTIES,
				Target: 'docProps/core.xml',
			}).up()
			.ele('Relationship', {
				Id: 'rId3',
				Type: REL_TYPE_OFFICE_DOCUMENT,
				Target: 'ppt/presentation.xml',
			}).up()
		.up()

	return doc.end({ prettyPrint: false })
}

/**
 * Creates `docProps/app.xml`
 * @param {PresSlide[]} slides - Presenation Slides
 * @param {string} company - "Company" metadata
 * @returns XML
 */
export function makeXmlApp (slides: PresSlide[], company: string): string {
	const doc = create({ version: '1.0', encoding: 'UTF-8', standalone: 'yes' })
		.ele('Properties', { xmlns: NS_EXTENDED_PROPERTIES, 'xmlns:vt': NS_VT })

	doc.ele('TotalTime').txt('0').up()
	doc.ele('Words').txt('0').up()
	doc.ele('Application').txt('Microsoft Office PowerPoint').up()
	doc.ele('PresentationFormat').txt('On-screen Show (16:9)').up()
	doc.ele('Paragraphs').txt('0').up()
	doc.ele('Slides').txt(String(slides.length)).up()
	doc.ele('Notes').txt(String(slides.length)).up()
	doc.ele('HiddenSlides').txt('0').up()
	doc.ele('MMClips').txt('0').up()
	doc.ele('ScaleCrop').txt('false').up()

	// HeadingPairs
	const headingPairs = doc.ele('HeadingPairs')
	const headingVector = headingPairs.ele('vt:vector', { size: '6', baseType: 'variant' })
	headingVector.ele('vt:variant').ele('vt:lpstr').txt('Fonts Used').up().up()
	headingVector.ele('vt:variant').ele('vt:i4').txt('2').up().up()
	headingVector.ele('vt:variant').ele('vt:lpstr').txt('Theme').up().up()
	headingVector.ele('vt:variant').ele('vt:i4').txt('1').up().up()
	headingVector.ele('vt:variant').ele('vt:lpstr').txt('Slide Titles').up().up()
	headingVector.ele('vt:variant').ele('vt:i4').txt(String(slides.length)).up().up()
	headingPairs.up()

	// TitlesOfParts
	const titlesOfParts = doc.ele('TitlesOfParts')
	const titlesVector = titlesOfParts.ele('vt:vector', { size: String(slides.length + 3), baseType: 'lpstr' })
	titlesVector.ele('vt:lpstr').txt('Arial').up()
	titlesVector.ele('vt:lpstr').txt('Calibri').up()
	titlesVector.ele('vt:lpstr').txt('Office Theme').up()
	slides.forEach((_slideObj, idx) => {
		titlesVector.ele('vt:lpstr').txt(`Slide ${idx + 1}`).up()
	})
	titlesOfParts.up()

	doc.ele('Company').txt(company).up()
	doc.ele('LinksUpToDate').txt('false').up()
	doc.ele('SharedDoc').txt('false').up()
	doc.ele('HyperlinksChanged').txt('false').up()
	doc.ele('AppVersion').txt('16.0000').up()

	return doc.end({ prettyPrint: false })
}

/**
 * Creates `docProps/core.xml`
 * @param {string} title - metadata data
 * @param {string} subject - metadata data
 * @param {string} author - metadata value
 * @param {string} revision - metadata value
 * @returns XML
 */
export function makeXmlCore (title: string, subject: string, author: string, revision: string): string {
	const isoTimestamp = new Date().toISOString().replace(/\.\d\d\dZ/, 'Z')

	const doc = create({ version: '1.0', encoding: 'UTF-8', standalone: 'yes' })
		.ele('cp:coreProperties', {
			'xmlns:cp': NS_CP,
			'xmlns:dc': NS_DC,
			'xmlns:dcterms': NS_DCTERMS,
			'xmlns:dcmitype': 'http://purl.org/dc/dcmitype/',
			'xmlns:xsi': NS_XSI,
		})

	doc.ele('dc:title').txt(title).up()
	doc.ele('dc:subject').txt(subject).up()
	doc.ele('dc:creator').txt(author).up()
	doc.ele('cp:lastModifiedBy').txt(author).up()
	doc.ele('cp:revision').txt(revision).up()
	doc.ele('dcterms:created', { 'xsi:type': 'dcterms:W3CDTF' }).txt(isoTimestamp).up()
	doc.ele('dcterms:modified', { 'xsi:type': 'dcterms:W3CDTF' }).txt(isoTimestamp).up()

	return doc.end({ prettyPrint: false })
}

/**
 * Creates `ppt/_rels/presentation.xml.rels`
 * @param {PresSlide[]} slides - Presenation Slides
 * @returns XML
 */
export function makeXmlPresentationRels (slides: PresSlide[]): string {
	const doc = create({ version: '1.0', encoding: 'UTF-8', standalone: 'yes' })
		.ele('Relationships', { xmlns: NS_RELATIONSHIPS })

	let relNum = 1

	// Slide Master
	doc.ele('Relationship', {
		Id: `rId${relNum}`,
		Type: REL_TYPE_SLIDE_MASTER,
		Target: 'slideMasters/slideMaster1.xml',
	}).up()

	// Slides
	for (let idx = 1; idx <= slides.length; idx++) {
		relNum++
		doc.ele('Relationship', {
			Id: `rId${relNum}`,
			Type: REL_TYPE_SLIDE,
			Target: `slides/slide${idx}.xml`,
		}).up()
	}

	relNum++
	// Notes Master
	doc.ele('Relationship', {
		Id: `rId${relNum}`,
		Type: REL_TYPE_NOTES_MASTER,
		Target: 'notesMasters/notesMaster1.xml',
	}).up()

	// Presentation Properties
	doc.ele('Relationship', {
		Id: `rId${relNum + 1}`,
		Type: REL_TYPE_PRES_PROPS,
		Target: 'presProps.xml',
	}).up()

	// View Properties
	doc.ele('Relationship', {
		Id: `rId${relNum + 2}`,
		Type: REL_TYPE_VIEW_PROPS,
		Target: 'viewProps.xml',
	}).up()

	// Theme
	doc.ele('Relationship', {
		Id: `rId${relNum + 3}`,
		Type: REL_TYPE_THEME,
		Target: 'theme/theme1.xml',
	}).up()

	// Table Styles
	doc.ele('Relationship', {
		Id: `rId${relNum + 4}`,
		Type: REL_TYPE_TABLE_STYLES,
		Target: 'tableStyles.xml',
	}).up()

	return doc.end({ prettyPrint: false })
}

// XML-GEN: Functions that run 1-N times (once for each Slide)

// =================================================================================================
// TRANSITION XML GENERATION
// =================================================================================================

/** Set of modern transitions that require p14 namespace (PowerPoint 2010+) */
const MODERN_TRANSITIONS = new Set([
	'morph', 'cube', 'box', 'doors', 'pan', 'ferris', 'gallery', 'conveyor',
	'flip', 'flythrough', 'glitter', 'honeycomb', 'origami', 'reveal',
	'ripple', 'shred', 'switch', 'vortex', 'warp', 'window'
])

/**
 * Generates XML for slide transition
 * @param {TransitionProps} transition - transition options
 * @return {string} XML for <p:transition> element
 */
function makeXmlTransition (transition: TransitionProps): string {
	if (!transition || transition.type === 'none') return ''

	const isMorph = transition.type === 'morph'
	const isModern = MODERN_TRANSITIONS.has(transition.type)

	// Build transition attributes
	const attrs: string[] = []

	// Speed/duration
	if (transition.speed) {
		attrs.push(`spd="${transition.speed}"`)
	} else if (transition.durationMs) {
		if (isModern) {
			// Modern transitions use p14:dur in 1/1000ths of a second
			attrs.push(`p14:dur="${transition.durationMs}"`)
		} else {
			// Classic transitions use spd attribute
			if (transition.durationMs <= 500) attrs.push('spd="fast"')
			else if (transition.durationMs <= 1500) attrs.push('spd="med"')
			else attrs.push('spd="slow"')
		}
	}

	// Advance on click
	if (transition.advanceOnClick === false) {
		attrs.push('advClick="0"')
	}

	// Auto-advance time
	if (transition.advanceAfterMs !== undefined) {
		attrs.push(`advTm="${transition.advanceAfterMs}"`)
	}

	const attrStr = attrs.length > 0 ? ' ' + attrs.join(' ') : ''

	// MORPH transition - requires special namespace (2015/09) and mc:AlternateContent wrapper
	if (isMorph) {
		const morphOption = transition.morphOption || 'byObject'
		// Morph uses the 2015/09 namespace, not p14
		return `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">` +
			`<mc:Choice xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main" Requires="p159">` +
			`<p:transition${attrStr.replace('p14:dur', 'p159:dur')} xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main">` +
			`<p159:morph option="${morphOption}"/>` +
			`</p:transition>` +
			`</mc:Choice>` +
			`<mc:Fallback>` +
			`<p:transition${attrStr.replace(/p14:dur="[^"]*"/, '').trim()}>` +
			`<p:fade/>` +
			`</p:transition>` +
			`</mc:Fallback>` +
			`</mc:AlternateContent>`
	}

	// Other modern transitions (p14 namespace)
	if (isModern) {
		// Build type element attributes
		const typeAttrs: string[] = []
		if (transition.direction) {
			typeAttrs.push(`dir="${transition.direction}"`)
		}

		if (transition.type === 'wheel') {
			typeAttrs.push('spokes="4"')
		}

		const typeAttrStr = typeAttrs.length > 0 ? ' ' + typeAttrs.join(' ') : ''

		// Use mc:AlternateContent for modern transitions
		return `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">` +
			`<mc:Choice xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" Requires="p14">` +
			`<p:transition${attrStr} xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">` +
			`<p14:${transition.type}${typeAttrStr}/>` +
			`</p:transition>` +
			`</mc:Choice>` +
			`<mc:Fallback>` +
			`<p:transition${attrStr.replace(/p14:dur="[^"]*"/, '').trim()}>` +
			`<p:fade/>` +
			`</p:transition>` +
			`</mc:Fallback>` +
			`</mc:AlternateContent>`
	}

	// Classic transitions (no special namespace needed)
	const typeAttrs: string[] = []
	if (transition.direction) {
		typeAttrs.push(`dir="${transition.direction}"`)
	}
	if (transition.type === 'wheel') {
		typeAttrs.push('spokes="4"')
	} else if (['wipe', 'push', 'cover', 'pull'].includes(transition.type) && !transition.direction) {
		typeAttrs.push('dir="l"')
	} else if (['split', 'blinds', 'comb', 'randomBar'].includes(transition.type) && !transition.direction) {
		typeAttrs.push('dir="horz"')
	}

	const typeAttrStr = typeAttrs.length > 0 ? ' ' + typeAttrs.join(' ') : ''

	return `<p:transition${attrStr}><p:${transition.type}${typeAttrStr}/></p:transition>`
}

// =================================================================================================
// ANIMATION XML GENERATION
// =================================================================================================

/**
 * Generates XML for slide animations (timing tree)
 * @param {PresSlide} slide - slide with animations
 * @return {string} XML for <p:timing> element
 */
function makeXmlTiming (slide: PresSlide): string {
	const animations = slide._animations
	if (!animations || animations.length === 0) return ''

	// Generate unique IDs starting from 2 (1 is reserved for root)
	let nextId = 2

	// Build animations as siblings - simple approach that matches working output
	// onClick → delay="indefinite", withPrevious → delay="0"
	let sequenceXml = ''

	for (const anim of animations) {
		const shapeId = anim.shapeIndex + 2
		const trigger = anim.options.trigger || 'onClick'

		// Determine delay based on trigger
		const delay = trigger === 'onClick' ? 'indefinite' : '0'

		const animNodeXml = makeAnimationNodeSimple(anim, shapeId, nextId, delay)
		sequenceXml += animNodeXml.xml
		nextId = animNodeXml.nextId
	}

	// Build main sequence container
	const mainSeqId = nextId++

	const animationXml = `<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst><p:seq concurrent="1" nextAc="seek"><p:cTn id="${mainSeqId}" dur="indefinite" nodeType="mainSeq"><p:childTnLst>${sequenceXml}</p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst><p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>`

	return animationXml
}

/**
 * Simple animation node generator
 * onClick uses delay="indefinite", withPrevious uses delay="0"
 */
function makeAnimationNodeSimple (
	anim: ISlideAnimation,
	shapeId: number,
	startId: number,
	delay: string
): { xml: string; nextId: number } {
	let id = startId
	const durationMs = anim.options.durationMs || 500

	// Build target element
	let targetXml = `<p:spTgt spid="${shapeId}"`
	if (anim.options.paragraphIndex !== undefined) {
		targetXml += `><p:txEl><p:pRg st="${anim.options.paragraphIndex}" end="${anim.options.paragraphIndex}"/></p:txEl></p:spTgt>`
	} else {
		targetXml += '/>'
	}

	// Build presetSubtype attribute
	const subtypeAttr = anim.presetSubtype ? ` presetSubtype="${anim.presetSubtype}"` : ''

	// IDs for the animation structure
	const outerId = id++
	const innerId = id++
	const effectParId = id++

	// Build the effect children
	let effectChildrenXml = ''

	// Set element (makes shape visible for entrance animations)
	if (anim.presetClass === 'entr') {
		const setCTnId = id++
		effectChildrenXml += `<p:set><p:cBhvr><p:cTn id="${setCTnId}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl>${targetXml}</p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>`
	}

	// Anim element (property animation for movement)
	const animId = id++
	effectChildrenXml += `<p:anim calcmode="lin" valueType="num"><p:cBhvr additive="base"><p:cTn id="${animId}" dur="${durationMs}"/><p:tgtEl>${targetXml}</p:tgtEl><p:attrNameLst><p:attrName>ppt_y</p:attrName></p:attrNameLst></p:cBhvr><p:tavLst><p:tav tm="0"><p:val><p:strVal val="#ppt_y+#ppt_h*0.1"/></p:val></p:tav><p:tav tm="100000"><p:val><p:strVal val="#ppt_y"/></p:val></p:tav></p:tavLst></p:anim>`

	// AnimEffect element (transition filter like fade)
	const animEffectId = id++
	const filter = getAnimationFilter(anim)
	if (filter) {
		effectChildrenXml += `<p:animEffect transition="in" filter="${filter}"><p:cBhvr><p:cTn id="${animEffectId}" dur="${durationMs}"/><p:tgtEl>${targetXml}</p:tgtEl></p:cBhvr></p:animEffect>`
	}

	const xml = `<p:par><p:cTn id="${outerId}" fill="hold"><p:stCondLst><p:cond delay="${delay}"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="${innerId}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="${effectParId}" presetID="${anim.presetId}" presetClass="${anim.presetClass}"${subtypeAttr} fill="hold" nodeType="clickEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>${effectChildrenXml}</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>`

	return { xml, nextId: id }
}
/**
 * Get animation filter string based on preset
 */
function getAnimationFilter (anim: ISlideAnimation): string {
	switch (anim.presetId) {
		case 10: // fade
			return 'fade'
		case 1: // appear
			return ''
		case 2: // fly-in
			return 'wipe(down)'
		case 3: // blinds
			return 'blinds(horizontal)'
		case 22: // split
			return 'split(horizontal)'
		case 28: // wipe
			return 'wipe(left)'
		case 29: // zoom
			return 'zoom'
		default:
			return 'fade'
	}
}

/**
 * Generates XML for the slide file (`ppt/slides/slide1.xml`)
 * @param {PresSlide} slide - the slide object to transform into XML
 * @return {string} XML
 */
export function makeXmlSlide (slide: PresSlide): string {
	// Build transition XML if present
	const transitionXml = slide._transition ? makeXmlTransition(slide._transition) : ''

	// Build timing/animation XML if present
	const timingXml = makeXmlTiming(slide)

	// Add extra namespaces for modern transitions (p14, mc, p159 for morph)
	const hasModernTransition = slide._transition && MODERN_TRANSITIONS.has(slide._transition.type)
	let extraNamespaces = ''
	if (hasModernTransition) {
		extraNamespaces = ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"' +
			' xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"'
		if (slide._transition?.type === 'morph') {
			extraNamespaces += ' xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main"'
		}
	}

	return (
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${CRLF}` +
		'<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
		`xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"${extraNamespaces}` +
		`${slide?.hidden ? ' show="0"' : ''}>` +
		`${slideObjectToXml(slide)}` +
		'<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>' +
		`${transitionXml}${timingXml}</p:sld>`
	)
}

/**
 * Get text content of Notes from Slide
 * @param {PresSlide} slide - the slide object to transform into XML
 * @return {string} notes text
 */
export function getNotesFromSlide (slide: PresSlide): string {
	let notesText = ''

	;(slide._slideObjects || []).forEach(data => {
		if (data._type === SLIDE_OBJECT_TYPES.notes) notesText += data?.text && data.text[0] ? data.text[0].text : ''
	})

	return notesText.replace(/\r*\n/g, CRLF)
}

/**
 * Generate XML for Notes Master (notesMaster1.xml)
 * @returns {string} XML
 */
export function makeXmlNotesMaster (): string {
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${CRLF}<p:notesMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="2" name="Header Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="hdr" sz="quarter"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2971800" cy="458788"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0"/><a:lstStyle><a:lvl1pPr algn="l"><a:defRPr sz="1200"/></a:lvl1pPr></a:lstStyle><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="3" name="Date Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="dt" idx="1"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="3884613" y="0"/><a:ext cx="2971800" cy="458788"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0"/><a:lstStyle><a:lvl1pPr algn="r"><a:defRPr sz="1200"/></a:lvl1pPr></a:lstStyle><a:p><a:fld id="{5282F153-3F37-0F45-9E97-73ACFA13230C}" type="datetimeFigureOut"><a:rPr lang="en-US"/><a:t>7/23/19</a:t></a:fld><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="4" name="Slide Image Placeholder 3"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldImg" idx="2"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="685800" y="1143000"/><a:ext cx="5486400" cy="3086100"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln w="12700"><a:solidFill><a:prstClr val="black"/></a:solidFill></a:ln></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="5" name="Notes Placeholder 4"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" sz="quarter" idx="3"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="685800" y="4400550"/><a:ext cx="5486400" cy="3600450"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0"/><a:lstStyle/><a:p><a:pPr lvl="0"/><a:r><a:rPr lang="en-US"/><a:t>Click to edit Master text styles</a:t></a:r></a:p><a:p><a:pPr lvl="1"/><a:r><a:rPr lang="en-US"/><a:t>Second level</a:t></a:r></a:p><a:p><a:pPr lvl="2"/><a:r><a:rPr lang="en-US"/><a:t>Third level</a:t></a:r></a:p><a:p><a:pPr lvl="3"/><a:r><a:rPr lang="en-US"/><a:t>Fourth level</a:t></a:r></a:p><a:p><a:pPr lvl="4"/><a:r><a:rPr lang="en-US"/><a:t>Fifth level</a:t></a:r></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="6" name="Footer Placeholder 5"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="ftr" sz="quarter" idx="4"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="8685213"/><a:ext cx="2971800" cy="458787"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="b"/><a:lstStyle><a:lvl1pPr algn="l"><a:defRPr sz="1200"/></a:lvl1pPr></a:lstStyle><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="7" name="Slide Number Placeholder 6"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldNum" sz="quarter" idx="5"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="3884613" y="8685213"/><a:ext cx="2971800" cy="458787"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="b"/><a:lstStyle><a:lvl1pPr algn="r"><a:defRPr sz="1200"/></a:lvl1pPr></a:lstStyle><a:p><a:fld id="{CE5E9CC1-C706-0F49-92D6-E571CC5EEA8F}" type="slidenum"><a:rPr lang="en-US"/><a:t>‹#›</a:t></a:fld><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp></p:spTree><p:extLst><p:ext uri="{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}"><p14:creationId xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="1024086991"/></p:ext></p:extLst></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:notesStyle><a:lvl1pPr marL="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl1pPr><a:lvl2pPr marL="457200" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl2pPr><a:lvl3pPr marL="914400" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl3pPr><a:lvl4pPr marL="1371600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl4pPr><a:lvl5pPr marL="1828800" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl5pPr><a:lvl6pPr marL="2286000" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl6pPr><a:lvl7pPr marL="2743200" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl7pPr><a:lvl8pPr marL="3200400" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl8pPr><a:lvl9pPr marL="3657600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl9pPr></p:notesStyle></p:notesMaster>`
}

/**
 * Creates Notes Slide (`ppt/notesSlides/notesSlide1.xml`)
 * @param {PresSlide} slide - the slide object to transform into XML
 * @return {string} XML
 */
export function makeXmlNotesSlide (slide: PresSlide): string {
	return (
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${CRLF}<p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp><p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" dirty="0"/><a:t>${encodeXmlEntities(getNotesFromSlide(slide))}</a:t></a:r><a:endParaRPr lang="en-US" dirty="0"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="4" name="Slide Number Placeholder 3"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldNum" sz="quarter" idx="10"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:fld id="${SLDNUMFLDID}" type="slidenum"><a:rPr lang="en-US"/><a:t>${slide._slideNum}</a:t></a:fld><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp></p:spTree><p:extLst><p:ext uri="{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}"><p14:creationId xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="1024086991"/></p:ext></p:extLst></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>`
	)
}

/**
 * Generates the XML layout resource from a layout object
 * @param {SlideLayout} layout - slide layout (master)
 * @return {string} XML
 */
export function makeXmlLayout (layout: SlideLayout): string {
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
		<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" preserve="1">
		${slideObjectToXml(layout)}
		<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`
}

/**
 * Creates Slide Master 1 (`ppt/slideMasters/slideMaster1.xml`)
 * @param {PresSlide} slide - slide object that represents master slide layout
 * @param {SlideLayout[]} layouts - slide layouts
 * @return {string} XML
 */
export function makeXmlMaster (slide: PresSlide, layouts: SlideLayout[]): string {
	// NOTE: Pass layouts as static rels because they are not referenced any time
	const layoutDefs = layouts.map((_layoutDef, idx) => `<p:sldLayoutId id="${LAYOUT_IDX_SERIES_BASE + idx}" r:id="rId${slide._rels.length + idx + 1}"/>`)

	let strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + CRLF
	strXml +=
		'<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
	strXml += slideObjectToXml(slide)
	strXml +=
		'<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>'
	strXml += '<p:sldLayoutIdLst>' + layoutDefs.join('') + '</p:sldLayoutIdLst>'
	strXml += '<p:hf sldNum="0" hdr="0" ftr="0" dt="0"/>'
	strXml +=
		'<p:txStyles>' +
		' <p:titleStyle>' +
		'  <a:lvl1pPr algn="ctr" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="0"/></a:spcBef><a:buNone/><a:defRPr sz="4400" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mj-lt"/><a:ea typeface="+mj-ea"/><a:cs typeface="+mj-cs"/></a:defRPr></a:lvl1pPr>' +
		' </p:titleStyle>' +
		' <p:bodyStyle>' +
		'  <a:lvl1pPr marL="342900" indent="-342900" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="3200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl1pPr>' +
		'  <a:lvl2pPr marL="742950" indent="-285750" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="–"/><a:defRPr sz="2800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl2pPr>' +
		'  <a:lvl3pPr marL="1143000" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="2400" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl3pPr>' +
		'  <a:lvl4pPr marL="1600200" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="–"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl4pPr>' +
		'  <a:lvl5pPr marL="2057400" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="»"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl5pPr>' +
		'  <a:lvl6pPr marL="2514600" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl6pPr>' +
		'  <a:lvl7pPr marL="2971800" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl7pPr>' +
		'  <a:lvl8pPr marL="3429000" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl8pPr>' +
		'  <a:lvl9pPr marL="3886200" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl9pPr>' +
		' </p:bodyStyle>' +
		' <p:otherStyle>' +
		'  <a:defPPr><a:defRPr lang="en-US"/></a:defPPr>' +
		'  <a:lvl1pPr marL="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl1pPr>' +
		'  <a:lvl2pPr marL="457200" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl2pPr>' +
		'  <a:lvl3pPr marL="914400" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl3pPr>' +
		'  <a:lvl4pPr marL="1371600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl4pPr>' +
		'  <a:lvl5pPr marL="1828800" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl5pPr>' +
		'  <a:lvl6pPr marL="2286000" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl6pPr>' +
		'  <a:lvl7pPr marL="2743200" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl7pPr>' +
		'  <a:lvl8pPr marL="3200400" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl8pPr>' +
		'  <a:lvl9pPr marL="3657600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl9pPr>' +
		' </p:otherStyle>' +
		'</p:txStyles>'
	strXml += '</p:sldMaster>'

	return strXml
}

/**
 * Generates XML string for a slide layout relation file
 * @param {number} layoutNumber - 1-indexed number of a layout that relations are generated for
 * @param {SlideLayout[]} slideLayouts - Slide Layouts
 * @return {string} XML
 */
export function makeXmlSlideLayoutRel (layoutNumber: number, slideLayouts: SlideLayout[]): string {
	return slideObjectRelationsToXml(slideLayouts[layoutNumber - 1], [
		{
			target: '../slideMasters/slideMaster1.xml',
			type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster',
		},
	])
}

/**
 * Creates `ppt/_rels/slide*.xml.rels`
 * @param {PresSlide[]} slides
 * @param {SlideLayout[]} slideLayouts - Slide Layout(s)
 * @param {number} `slideNumber` 1-indexed number of a layout that relations are generated for
 * @return {string} XML
 */
export function makeXmlSlideRel (slides: PresSlide[], slideLayouts: SlideLayout[], slideNumber: number): string {
	return slideObjectRelationsToXml(slides[slideNumber - 1], [
		{
			target: `../slideLayouts/slideLayout${getLayoutIdxForSlide(slides, slideLayouts, slideNumber)}.xml`,
			type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout',
		},
		{
			target: `../notesSlides/notesSlide${slideNumber}.xml`,
			type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide',
		},
	])
}

/**
 * Generates XML string for a slide relation file.
 * @param {number} slideNumber - 1-indexed number of a layout that relations are generated for
 * @return {string} XML
 */
export function makeXmlNotesSlideRel (slideNumber: number): string {
	const doc = create({ version: '1.0', encoding: 'UTF-8', standalone: 'yes' })
		.ele('Relationships', { xmlns: NS_RELATIONSHIPS })
			.ele('Relationship', {
				Id: 'rId1',
				Type: REL_TYPE_NOTES_MASTER,
				Target: '../notesMasters/notesMaster1.xml',
			}).up()
			.ele('Relationship', {
				Id: 'rId2',
				Type: REL_TYPE_SLIDE,
				Target: `../slides/slide${slideNumber}.xml`,
			}).up()
		.up()

	return doc.end({ prettyPrint: false })
}

/**
 * Creates `ppt/slideMasters/_rels/slideMaster1.xml.rels`
 * @param {PresSlide} masterSlide - Slide object
 * @param {SlideLayout[]} slideLayouts - Slide Layouts
 * @return {string} XML
 */
export function makeXmlMasterRel (masterSlide: PresSlide, slideLayouts: SlideLayout[]): string {
	const defaultRels = slideLayouts.map((_layoutDef, idx) => ({
		target: `../slideLayouts/slideLayout${idx + 1}.xml`,
		type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout',
	}))
	defaultRels.push({ target: '../theme/theme1.xml', type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme' })

	return slideObjectRelationsToXml(masterSlide, defaultRels)
}

/**
 * Creates `ppt/notesMasters/_rels/notesMaster1.xml.rels`
 * @return {string} XML
 */
export function makeXmlNotesMasterRel (): string {
	const doc = create({ version: '1.0', encoding: 'UTF-8', standalone: 'yes' })
		.ele('Relationships', { xmlns: NS_RELATIONSHIPS })
			.ele('Relationship', {
				Id: 'rId1',
				Type: REL_TYPE_THEME,
				Target: '../theme/theme1.xml',
			}).up()
		.up()

	return doc.end({ prettyPrint: false })
}

/**
 * For the passed slide number, resolves name of a layout that is used for.
 * @param {PresSlide[]} slides - srray of slides
 * @param {SlideLayout[]} slideLayouts - array of slideLayouts
 * @param {number} slideNumber
 * @return {number} slide number
 */
function getLayoutIdxForSlide (slides: PresSlide[], slideLayouts: SlideLayout[], slideNumber: number): number {
	for (let i = 0; i < slideLayouts.length; i++) {
		if (slideLayouts[i]._name === slides[slideNumber - 1]?._slideLayout?._name) {
			return i + 1
		}
	}

	// IMPORTANT: Return 1 (for `slideLayout1.xml`) when no def is found
	// So all objects are in Layout1 and every slide that references it uses this layout.
	return 1
}

// XML-GEN: Last 5 functions create root /ppt files

/**
 * Creates `ppt/theme/theme1.xml`
 * @return {string} XML
 */
export function makeXmlTheme (pres: IPresentationProps): string {
	const majorFont = pres.theme?.headFontFace ? `<a:latin typeface="${pres.theme?.headFontFace}"/>` : '<a:latin typeface="Calibri Light" panose="020F0302020204030204"/>'
	const minorFont = pres.theme?.bodyFontFace ? `<a:latin typeface="${pres.theme?.bodyFontFace}"/>` : '<a:latin typeface="Calibri" panose="020F0502020204030204"/>'
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont>${majorFont}<a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="游ゴシック Light"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="等线 Light"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Times New Roman"/><a:font script="Hebr" typeface="Times New Roman"/><a:font script="Thai" typeface="Angsana New"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="MoolBoran"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Times New Roman"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/><a:font script="Armn" typeface="Arial"/><a:font script="Bugi" typeface="Leelawadee UI"/><a:font script="Bopo" typeface="Microsoft JhengHei"/><a:font script="Java" typeface="Javanese Text"/><a:font script="Lisu" typeface="Segoe UI"/><a:font script="Mymr" typeface="Myanmar Text"/><a:font script="Nkoo" typeface="Ebrima"/><a:font script="Olck" typeface="Nirmala UI"/><a:font script="Osma" typeface="Ebrima"/><a:font script="Phag" typeface="Phagspa"/><a:font script="Syrn" typeface="Estrangelo Edessa"/><a:font script="Syrj" typeface="Estrangelo Edessa"/><a:font script="Syre" typeface="Estrangelo Edessa"/><a:font script="Sora" typeface="Nirmala UI"/><a:font script="Tale" typeface="Microsoft Tai Le"/><a:font script="Talu" typeface="Microsoft New Tai Lue"/><a:font script="Tfng" typeface="Ebrima"/></a:majorFont><a:minorFont>${minorFont}<a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="游ゴシック"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="等线"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Arial"/><a:font script="Hebr" typeface="Arial"/><a:font script="Thai" typeface="Cordia New"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="DaunPenh"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Arial"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/><a:font script="Armn" typeface="Arial"/><a:font script="Bugi" typeface="Leelawadee UI"/><a:font script="Bopo" typeface="Microsoft JhengHei"/><a:font script="Java" typeface="Javanese Text"/><a:font script="Lisu" typeface="Segoe UI"/><a:font script="Mymr" typeface="Myanmar Text"/><a:font script="Nkoo" typeface="Ebrima"/><a:font script="Olck" typeface="Nirmala UI"/><a:font script="Osma" typeface="Ebrima"/><a:font script="Phag" typeface="Phagspa"/><a:font script="Syrn" typeface="Estrangelo Edessa"/><a:font script="Syrj" typeface="Estrangelo Edessa"/><a:font script="Syre" typeface="Estrangelo Edessa"/><a:font script="Sora" typeface="Nirmala UI"/><a:font script="Tale" typeface="Microsoft Tai Le"/><a:font script="Talu" typeface="Microsoft New Tai Lue"/><a:font script="Tfng" typeface="Ebrima"/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="103000"/><a:tint val="73000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:satMod val="103000"/><a:lumMod val="102000"/><a:tint val="94000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:satMod val="110000"/><a:lumMod val="100000"/><a:shade val="100000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="99000"/><a:satMod val="120000"/><a:shade val="78000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="98000"/><a:satMod val="130000"/><a:shade val="90000"/><a:lumMod val="103000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/><a:satMod val="120000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/><a:extLst><a:ext uri="{05A4C25C-085E-4340-85A3-A5531E510DB2}"><thm15:themeFamily xmlns:thm15="http://schemas.microsoft.com/office/thememl/2012/main" name="Office Theme" id="{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}" vid="{4A3C46E8-61CC-4603-A589-7422A47A8E4A}"/></a:ext></a:extLst></a:theme>`
}

/**
 * Create presentation file (`ppt/presentation.xml`)
 * @see https://docs.microsoft.com/en-us/office/open-xml/structure-of-a-presentationml-document
 * @see http://www.datypic.com/sc/ooxml/t-p_CT_Presentation.html
 * @param {IPresentationProps} pres - presentation
 * @return {string} XML
 */
export function makeXmlPresentation (pres: IPresentationProps): string {
	const rootAttrs: Record<string, string> = {
		'xmlns:a': NS_A,
		'xmlns:r': NS_R,
		'xmlns:p': NS_P,
		saveSubsetFonts: '1',
		autoCompressPictures: '0',
	}
	if (pres.rtlMode) {
		rootAttrs.rtl = '1'
	}

	const doc = create({ version: '1.0', encoding: 'UTF-8', standalone: 'yes' })
		.ele('p:presentation', rootAttrs)

	// STEP 1: Add slide master (SPEC: tag 1 under <presentation>)
	doc.ele('p:sldMasterIdLst')
		.ele('p:sldMasterId', { id: '2147483648', 'r:id': 'rId1' }).up()
		.up()

	// STEP 2: Add all Slides (SPEC: tag 3 under <presentation>)
	const sldIdLst = doc.ele('p:sldIdLst')
	pres.slides.forEach(slide => {
		sldIdLst.ele('p:sldId', { id: String(slide._slideId), 'r:id': `rId${slide._rId}` }).up()
	})
	sldIdLst.up()

	// STEP 3: Add Notes Master (SPEC: tag 2 under <presentation>)
	// (NOTE: length+2 is from `presentation.xml.rels` func (since we have to match this rId, we just use same logic))
	// IMPORTANT: In this order (matches PPT2019) PPT will give corruption message on open!
	// IMPORTANT: Placing this before `<p:sldIdLst>` causes warning in modern powerpoint!
	// IMPORTANT: Presentations open without warning Without this line, however, the pres isnt preview in Finder anymore or viewable in iOS!
	doc.ele('p:notesMasterIdLst')
		.ele('p:notesMasterId', { 'r:id': `rId${pres.slides.length + 2}` }).up()
		.up()

	// STEP 4: Add sizes
	doc.ele('p:sldSz', { cx: String(pres.presLayout.width), cy: String(pres.presLayout.height) }).up()
	doc.ele('p:notesSz', { cx: String(pres.presLayout.height), cy: String(pres.presLayout.width) }).up()

	// STEP 5: Add text styles
	const defaultTextStyle = doc.ele('p:defaultTextStyle')
	for (let idy = 1; idy < 10; idy++) {
		const lvlPPr = defaultTextStyle.ele(`a:lvl${idy}pPr`, {
			marL: String((idy - 1) * 457200),
			algn: 'l',
			defTabSz: '914400',
			rtl: '0',
			eaLnBrk: '1',
			latinLnBrk: '0',
			hangingPunct: '1',
		})
		const defRPr = lvlPPr.ele('a:defRPr', { sz: '1800', kern: '1200' })
		defRPr.ele('a:solidFill').ele('a:schemeClr', { val: 'tx1' }).up().up()
		defRPr.ele('a:latin', { typeface: '+mn-lt' }).up()
		defRPr.ele('a:ea', { typeface: '+mn-ea' }).up()
		defRPr.ele('a:cs', { typeface: '+mn-cs' }).up()
		defRPr.up()
		lvlPPr.up()
	}
	defaultTextStyle.up()

	// STEP 6: Add Sections (if any)
	if (pres.sections && pres.sections.length > 0) {
		const extLst = doc.ele('p:extLst')
		const ext1 = extLst.ele('p:ext', { uri: '{521415D9-36F7-43E2-AB2F-B90AF26B5E84}' })
		const sectionLst = ext1.ele('p14:sectionLst', { 'xmlns:p14': NS_P14 })

		pres.sections.forEach(sect => {
			const section = sectionLst.ele('p14:section', {
				name: sect.title,
				id: `{${getUuid('xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx')}}`,
			})
			const sldIdLstSect = section.ele('p14:sldIdLst')
			sect._slides.forEach(slide => {
				sldIdLstSect.ele('p14:sldId', { id: String(slide._slideId) }).up()
			})
			sldIdLstSect.up()
			section.up()
		})

		sectionLst.up()
		ext1.up()

		extLst.ele('p:ext', { uri: '{EFAFB233-063F-42B5-8137-9DF3F51BA10A}' })
			.ele('p15:sldGuideLst', { 'xmlns:p15': NS_P15 }).up()
			.up()
		extLst.up()
	}

	return doc.end({ prettyPrint: false })
}

/**
 * Create `ppt/presProps.xml`
 * @return {string} XML
 */
export function makeXmlPresProps (): string {
	const doc = create({ version: '1.0', encoding: 'UTF-8', standalone: 'yes' })
		.ele('p:presentationPr', {
			'xmlns:a': NS_A,
			'xmlns:r': NS_R,
			'xmlns:p': NS_P,
		})

	return doc.end({ prettyPrint: false })
}

/**
 * Create `ppt/tableStyles.xml`
 * @see: http://openxmldeveloper.org/discussions/formats/f/13/p/2398/8107.aspx
 * @return {string} XML
 */
export function makeXmlTableStyles (): string {
	const doc = create({ version: '1.0', encoding: 'UTF-8', standalone: 'yes' })
		.ele('a:tblStyleLst', {
			'xmlns:a': NS_A,
			def: '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}',
		})

	return doc.end({ prettyPrint: false })
}

/**
 * Creates `ppt/viewProps.xml`
 * @return {string} XML
 */
export function makeXmlViewProps (): string {
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${CRLF}<p:viewPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:normalViewPr horzBarState="maximized"><p:restoredLeft sz="15611"/><p:restoredTop sz="94610"/></p:normalViewPr><p:slideViewPr><p:cSldViewPr snapToGrid="0" snapToObjects="1"><p:cViewPr varScale="1"><p:scale><a:sx n="136" d="100"/><a:sy n="136" d="100"/></p:scale><p:origin x="216" y="312"/></p:cViewPr><p:guideLst/></p:cSldViewPr></p:slideViewPr><p:notesTextViewPr><p:cViewPr><p:scale><a:sx n="1" d="1"/><a:sy n="1" d="1"/></p:scale><p:origin x="0" y="0"/></p:cViewPr></p:notesTextViewPr><p:gridSpacing cx="76200" cy="76200"/></p:viewPr>`
}

/**
 * Checks shadow options passed by user and performs corrections if needed.
 * @param {ShadowProps} shadowProps - shadow options
 */
export function correctShadowOptions (shadowProps: ShadowProps): void {
	if (!shadowProps || typeof shadowProps !== 'object') {
		// console.warn("`shadow` options must be an object. Ex: `{shadow: {type:'none'}}`")
		return
	}

	// OPT: `type`
	if (shadowProps.type !== 'outer' && shadowProps.type !== 'inner' && shadowProps.type !== 'none') {
		console.warn('Warning: shadow.type options are `outer`, `inner` or `none`.')
		shadowProps.type = 'outer'
	}

	// OPT: `angle`
	if (shadowProps.angle) {
		// A: REALITY-CHECK
		if (isNaN(Number(shadowProps.angle)) || shadowProps.angle < 0 || shadowProps.angle > 359) {
			console.warn('Warning: shadow.angle can only be 0-359')
			shadowProps.angle = 270
		}

		// B: ROBUST: Cast any type of valid arg to int: '12', 12.3, etc. -> 12
		shadowProps.angle = Math.round(Number(shadowProps.angle))
	}

	// OPT: `opacity`
	if (shadowProps.opacity) {
		// A: REALITY-CHECK
		if (isNaN(Number(shadowProps.opacity)) || shadowProps.opacity < 0 || shadowProps.opacity > 1) {
			console.warn('Warning: shadow.opacity can only be 0-1')
			shadowProps.opacity = 0.75
		}

		// B: ROBUST: Cast any type of valid arg to int: '12', 12.3, etc. -> 12
		shadowProps.opacity = Number(shadowProps.opacity)
	}
}
