// @ts-nocheck
/**
 * PptxGenJS: Slide Object XML Generation
 *
 * This file contains the legacy slideObjectToXml function and its helpers.
 * It uses @ts-nocheck because the original code has many TypeScript issues
 * (mostly 'possibly undefined' errors) that would take significant time to fix.
 *
 * The rest of gen-xml.ts has been refactored to be fully typed.
 */

import { fragment } from 'xmlbuilder2'

import {
	DEF_CELL_MARGIN_IN,
	DEF_PRES_LAYOUT_NAME,
	DEF_TEXT_SHADOW,
	EMU,
	SLDNUMFLDID,
	SLIDE_OBJECT_TYPES,
} from './core-enums'
import type {
	ISlideObject,
	ObjectOptions,
	PresSlide,
	ShadowProps,
	SlideLayout,
	TableCell,
	TableCellProps,
} from './core-interfaces'
import {
	convertRotationDegrees,
	createColorElement,
	encodeXmlEntities,
	genXmlColorSelection,
	getSmartParseNumber,
	inch2Emu,
	valToPts,
} from './gen-utils'
import { genXmlTextBody, genXmlPlaceholder } from './gen-xml-text'

/**
 * Compute shadow XML values without mutating the input object.
 * Returns undefined if shadow is not defined or type is 'none'.
 */
function computeShadowXmlValues(shadow: ShadowProps | undefined): {
	type: string
	blur: number
	offset: number
	angle: number
	opacity: number
	color: string
} | undefined {
	if (!shadow || shadow.type === 'none') {
		return undefined
	}
	return {
		type: shadow.type || 'outer',
		blur: valToPts(shadow.blur || 8),
		offset: valToPts(shadow.offset || 4),
		angle: Math.round((shadow.angle || 270) * 60000),
		opacity: Math.round((shadow.opacity || 0.75) * 100000),
		color: shadow.color || DEF_TEXT_SHADOW.color,
	}
}

/**
 * Generate shadow effect XML string.
 */
function genShadowEffectXml(shadow: ShadowProps | undefined): string {
	const computed = computeShadowXmlValues(shadow)
	if (!computed) return ''

	const shadowElementName = `a:${computed.type}Shdw`
	const shadowAttrs: Record<string, string> = {
		blurRad: String(computed.blur),
		dist: String(computed.offset),
		dir: String(computed.angle),
	}

	if (computed.type === 'outer') {
		Object.assign(shadowAttrs, {
			sx: '100000',
			sy: '100000',
			kx: '0',
			ky: '0',
			algn: 'bl',
			rotWithShape: '0',
		})
	}

	const frag = fragment()
		.ele('a:effectLst')
			.ele(shadowElementName, shadowAttrs)
				.ele('a:srgbClr', { val: computed.color })
					.ele('a:alpha', { val: String(computed.opacity) }).up()
				.up()
			.up()
		.up()

	return frag.toString({ prettyPrint: false })
}

type ImageDimensions = { w: number, h: number }
type BoxDimensions = { w: number, h: number, x: number, y: number }

const ImageSizingXml = {
	cover: function (imgSize: ImageDimensions, boxDim: BoxDimensions): string {
		const imgRatio = imgSize.h / imgSize.w
		const boxRatio = boxDim.h / boxDim.w
		const isBoxBased = boxRatio > imgRatio
		const width = isBoxBased ? boxDim.h / imgRatio : boxDim.w
		const height = isBoxBased ? boxDim.h : boxDim.w * imgRatio
		const hzPerc = Math.round(1e5 * 0.5 * (1 - boxDim.w / width))
		const vzPerc = Math.round(1e5 * 0.5 * (1 - boxDim.h / height))

		return fragment()
			.ele('a:srcRect', { l: String(hzPerc), r: String(hzPerc), t: String(vzPerc), b: String(vzPerc) }).up()
			.ele('a:stretch').up()
			.toString({ prettyPrint: false })
	},
	contain: function (imgSize: ImageDimensions, boxDim: BoxDimensions): string {
		const imgRatio = imgSize.h / imgSize.w
		const boxRatio = boxDim.h / boxDim.w
		const widthBased = boxRatio > imgRatio
		const width = widthBased ? boxDim.w : boxDim.h / imgRatio
		const height = widthBased ? boxDim.w * imgRatio : boxDim.h
		const hzPerc = Math.round(1e5 * 0.5 * (1 - boxDim.w / width))
		const vzPerc = Math.round(1e5 * 0.5 * (1 - boxDim.h / height))

		return fragment()
			.ele('a:srcRect', { l: String(hzPerc), r: String(hzPerc), t: String(vzPerc), b: String(vzPerc) }).up()
			.ele('a:stretch').up()
			.toString({ prettyPrint: false })
	},
	crop: function (imgSize: ImageDimensions, boxDim: BoxDimensions): string {
		const l = boxDim.x
		const r = imgSize.w - (boxDim.x + boxDim.w)
		const t = boxDim.y
		const b = imgSize.h - (boxDim.y + boxDim.h)
		const lPerc = Math.round(1e5 * (l / imgSize.w))
		const rPerc = Math.round(1e5 * (r / imgSize.w))
		const tPerc = Math.round(1e5 * (t / imgSize.h))
		const bPerc = Math.round(1e5 * (b / imgSize.h))

		return fragment()
			.ele('a:srcRect', { l: String(lPerc), r: String(rPerc), t: String(tPerc), b: String(bPerc) }).up()
			.ele('a:stretch').up()
			.toString({ prettyPrint: false })
	},
}

/**
 * Transforms a slide or slideLayout to resulting XML string - Creates `ppt/slide*.xml`
 * @param {PresSlide|SlideLayout} slideObject - slide object created within createSlideObject
 * @return {string} XML string with <p:cSld> as the root
 */
export function slideObjectToXml (slide: PresSlide | SlideLayout): string {
	let strSlideXml: string = slide._name ? '<p:cSld name="' + slide._name + '">' : '<p:cSld>'
	let intTableNum = 1

	// STEP 1: Add background color/image (ensure only a single `<p:bg>` tag is created, ex: when master-baskground has both `color` and `path`)
	if (slide._bkgdImgRid) {
		strSlideXml += `<p:bg><p:bgPr><a:blipFill dpi="0" rotWithShape="1"><a:blip r:embed="rId${slide._bkgdImgRid}"><a:lum/></a:blip><a:srcRect/><a:stretch><a:fillRect/></a:stretch></a:blipFill><a:effectLst/></p:bgPr></p:bg>`
	} else if (slide.background?.color) {
		strSlideXml += `<p:bg><p:bgPr>${genXmlColorSelection(slide.background)}</p:bgPr></p:bg>`
	} else if (!slide.bkgd && slide._name && slide._name === DEF_PRES_LAYOUT_NAME) {
		// NOTE: Default [white] background is needed on slideMaster1.xml to avoid gray background in Keynote (and Finder previews)
		strSlideXml += '<p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>'
	}

	// STEP 2: Continue slide by starting spTree node
	strSlideXml += '<p:spTree>'
	strSlideXml += '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
	strSlideXml += '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
	strSlideXml += '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'

	// STEP 3: Loop over all Slide.data objects and add them to this slide
	slide._slideObjects.forEach((slideItemObj: ISlideObject, idx: number) => {
		let x = 0
		let y = 0
		let cx = getSmartParseNumber('75%', 'X', slide._presLayout)
		let cy = 0
		let placeholderObj: ISlideObject
		let locationAttr = ''
		let arrTabRows: TableCell[][] = null
		let objTabOpts: ObjectOptions = null
		let intColCnt = 0
		let intColW = 0
		let cellOpts: TableCellProps = null
		let strXml: string = null
		const sizing: ObjectOptions['sizing'] = slideItemObj.options?.sizing
		const rounding = slideItemObj.options?.rounding

		if (
			(slide as PresSlide)._slideLayout !== undefined &&
			(slide as PresSlide)._slideLayout._slideObjects !== undefined &&
			slideItemObj.options &&
			slideItemObj.options.placeholder
		) {
			placeholderObj = (slide as PresSlide)._slideLayout._slideObjects.filter(
				(object: ISlideObject) => object.options.placeholder === slideItemObj.options.placeholder
			)[0]
		}

		// A: Set option vars
		slideItemObj.options = slideItemObj.options || {}

		if (typeof slideItemObj.options.x !== 'undefined') x = getSmartParseNumber(slideItemObj.options.x, 'X', slide._presLayout)
		if (typeof slideItemObj.options.y !== 'undefined') y = getSmartParseNumber(slideItemObj.options.y, 'Y', slide._presLayout)
		if (typeof slideItemObj.options.w !== 'undefined') cx = getSmartParseNumber(slideItemObj.options.w, 'X', slide._presLayout)
		if (typeof slideItemObj.options.h !== 'undefined') cy = getSmartParseNumber(slideItemObj.options.h, 'Y', slide._presLayout)

		// Set w/h now that smart parse is done
		let imgWidth = cx
		let imgHeight = cy

		// If using a placeholder then inherit it's position
		if (placeholderObj) {
			if (placeholderObj.options.x || placeholderObj.options.x === 0) x = getSmartParseNumber(placeholderObj.options.x, 'X', slide._presLayout)
			if (placeholderObj.options.y || placeholderObj.options.y === 0) y = getSmartParseNumber(placeholderObj.options.y, 'Y', slide._presLayout)
			if (placeholderObj.options.w || placeholderObj.options.w === 0) cx = getSmartParseNumber(placeholderObj.options.w, 'X', slide._presLayout)
			if (placeholderObj.options.h || placeholderObj.options.h === 0) cy = getSmartParseNumber(placeholderObj.options.h, 'Y', slide._presLayout)
		}
		//
		if (slideItemObj.options.flipH) locationAttr += ' flipH="1"'
		if (slideItemObj.options.flipV) locationAttr += ' flipV="1"'
		if (slideItemObj.options.rotate) locationAttr += ` rot="${convertRotationDegrees(slideItemObj.options.rotate)}"`

		// B: Add OBJECT to the current Slide
		switch (slideItemObj._type) {
			case SLIDE_OBJECT_TYPES.table:
				arrTabRows = slideItemObj.arrTabRows
				objTabOpts = slideItemObj.options
				intColCnt = 0
				intColW = 0

				// Calc number of columns
				// NOTE: Cells may have a colspan, so merely taking the length of the [0] (or any other) row is not
				// ....: sufficient to determine column count. Therefore, check each cell for a colspan and total cols as reqd
				arrTabRows[0].forEach(cell => {
					cellOpts = cell.options || null
					intColCnt += cellOpts?.colspan ? Number(cellOpts.colspan) : 1
				})

				// STEP 1: Start Table XML
				// NOTE: Non-numeric cNvPr id values will trigger "presentation needs repair" type warning in MS-PPT-2013
				strXml = `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${intTableNum * slide._slideNum + 1}" name="${slideItemObj.options.objectName}"/>`
				strXml +=
					'<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>' +
					'  <p:nvPr><p:extLst><p:ext uri="{D42A27DB-BD31-4B8C-83A1-F6EECF244321}"><p14:modId xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="1579011935"/></p:ext></p:extLst></p:nvPr>' +
					'</p:nvGraphicFramePr>'
				strXml += `<p:xfrm><a:off x="${x || (x === 0 ? 0 : EMU)}" y="${y || (y === 0 ? 0 : EMU)}"/><a:ext cx="${cx || (cx === 0 ? 0 : EMU)}" cy="${cy || EMU
				}"/></p:xfrm>`
				strXml += '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr/>'
				// + '        <a:tblPr bandRow="1"/>';
				// TODO: Support banded rows, first/last row, etc.
				// NOTE: Banding, etc. only shows when using a table style! (or set alt row color if banding)
				// <a:tblPr firstCol="0" firstRow="0" lastCol="0" lastRow="0" bandCol="0" bandRow="1">

				// STEP 2: Set column widths
				// Evenly distribute cols/rows across size provided when applicable (calc them if only overall dimensions were provided)
				// A: Col widths provided?
				// B: Table Width provided without colW? Then distribute cols
				if (Array.isArray(objTabOpts.colW)) {
					strXml += '<a:tblGrid>'
					for (let col = 0; col < intColCnt; col++) {
						let w = inch2Emu(objTabOpts.colW[col])
						if (w == null || isNaN(w)) {
							w = (typeof slideItemObj.options.w === 'number' ? slideItemObj.options.w : 1) / intColCnt
						}
						strXml += `<a:gridCol w="${Math.round(w)}"/>`
					}
					strXml += '</a:tblGrid>'
				} else {
					intColW = objTabOpts.colW ? objTabOpts.colW : EMU
					if (slideItemObj.options.w && !objTabOpts.colW) intColW = Math.round((typeof slideItemObj.options.w === 'number' ? slideItemObj.options.w : 1) / intColCnt)
					strXml += '<a:tblGrid>'
					for (let colw = 0; colw < intColCnt; colw++) {
						strXml += `<a:gridCol w="${intColW}"/>`
					}
					strXml += '</a:tblGrid>'
				}

				// STEP 3: Build our row arrays into an actual grid to match the XML we will be building next (ISSUE #36)
				// Note row arrays can arrive "lopsided" as in row1:[1,2,3] row2:[3] when first two cols rowspan!,
				// so a simple loop below in XML building wont suffice to build table correctly.
				// We have to build an actual grid now
				/*
					EX: (A0:rowspan=3, B1:rowspan=2, C1:colspan=2)

					/------|------|------|------\
					|  A0  |  B0  |  C0  |  D0  |
					|      |  B1  |  C1  |      |
					|      |      |  C2  |  D2  |
					\------|------|------|------/
				*/
				// A: add _hmerge cell for colspan. should reserve rowspan
				arrTabRows.forEach(cells => {
					for (let cIdx = 0; cIdx < cells.length;) {
						const cell = cells[cIdx]
						const colspan = cell.options?.colspan
						const rowspan = cell.options?.rowspan
						if (colspan && colspan > 1) {
							const vMergeCells = new Array(colspan - 1).fill(undefined).map(() => {
								return { _type: SLIDE_OBJECT_TYPES.tablecell, options: { rowspan }, _hmerge: true } as const
							})
							cells.splice(cIdx + 1, 0, ...vMergeCells)
							cIdx += colspan
						} else {
							cIdx += 1
						}
					}
				})
				// B: add _vmerge cell for rowspan. should reserve colspan/_hmerge
				arrTabRows.forEach((cells, rIdx) => {
					const nextRow = arrTabRows[rIdx + 1]
					if (!nextRow) return
					cells.forEach((cell, cIdx) => {
						const rowspan = cell._rowContinue || cell.options?.rowspan
						const colspan = cell.options?.colspan
						const _hmerge = cell._hmerge
						if (rowspan && rowspan > 1) {
							const hMergeCell = { _type: SLIDE_OBJECT_TYPES.tablecell, options: { colspan }, _rowContinue: rowspan - 1, _vmerge: true, _hmerge } as const
							nextRow.splice(cIdx, 0, hMergeCell)
						}
					})
				})

				// STEP 4: Build table rows/cells
				arrTabRows.forEach((cells, rIdx) => {
					// A: Table Height provided without rowH? Then distribute rows
					let intRowH = 0 // IMPORTANT: Default must be zero for auto-sizing to work
					if (Array.isArray(objTabOpts.rowH) && objTabOpts.rowH[rIdx]) intRowH = inch2Emu(Number(objTabOpts.rowH[rIdx]))
					else if (objTabOpts.rowH && !isNaN(Number(objTabOpts.rowH))) intRowH = inch2Emu(Number(objTabOpts.rowH))
					else if (slideItemObj.options.cy || slideItemObj.options.h) {
						intRowH = Math.round(
							(slideItemObj.options.h ? inch2Emu(slideItemObj.options.h) : typeof slideItemObj.options.cy === 'number' ? slideItemObj.options.cy : 1) /
							arrTabRows.length
						)
					}

					// B: Start row
					strXml += `<a:tr h="${intRowH}">`

					// C: Loop over each CELL
					cells.forEach(cellObj => {
						const cell: TableCell = cellObj

						const cellSpanAttrs = {
							rowSpan: cell.options?.rowspan > 1 ? cell.options.rowspan : undefined,
							gridSpan: cell.options?.colspan > 1 ? cell.options.colspan : undefined,
							vMerge: cell._vmerge ? 1 : undefined,
							hMerge: cell._hmerge ? 1 : undefined,
						}
						let cellSpanAttrStr = Object.keys(cellSpanAttrs)
							.map(k => [k, cellSpanAttrs[k]])
							.filter(([, v]) => !!v)
							.map(([k, v]) => `${String(k)}="${String(v)}"`)
							.join(' ')
						if (cellSpanAttrStr) cellSpanAttrStr = ' ' + cellSpanAttrStr

						// 1: COLSPAN/ROWSPAN: Add dummy cells for any active colspan/rowspan
						if (cell._hmerge || cell._vmerge) {
							strXml += `<a:tc${cellSpanAttrStr}><a:tcPr/></a:tc>`
							return
						}

						// 2: OPTIONS: Build/set cell options
						const cellOpts = cell.options || {}
						cell.options = cellOpts

						// B: Inherit some options from table when cell options dont exist
						// @see: http://officeopenxml.com/drwTableCellProperties-alignment.php
						;['align', 'bold', 'border', 'color', 'fill', 'fontFace', 'fontSize', 'margin', 'textDirection', 'underline', 'valign'].forEach(name => {
							if (objTabOpts[name] && !cellOpts[name] && cellOpts[name] !== 0) cellOpts[name] = objTabOpts[name]
						})

						const cellValign = cellOpts.valign
							? ` anchor="${cellOpts.valign.replace(/^c$/i, 'ctr').replace(/^m$/i, 'ctr').replace('center', 'ctr').replace('middle', 'ctr').replace('top', 't').replace('btm', 'b').replace('bottom', 'b')}"`
							: ''
						const cellTextDir = (cellOpts.textDirection && cellOpts.textDirection !== 'horz') ? ` vert="${cellOpts.textDirection}"` : ''

						let fillColor =
							cell._optImp?.fill?.color
								? cell._optImp.fill.color
								: cell._optImp?.fill && typeof cell._optImp.fill === 'string'
									? cell._optImp.fill
									: ''
						fillColor = fillColor || cellOpts.fill ? cellOpts.fill : ''
						const cellFill = fillColor ? genXmlColorSelection(fillColor) : ''

						let cellMargin = cellOpts.margin === 0 || cellOpts.margin ? cellOpts.margin : DEF_CELL_MARGIN_IN
						if (!Array.isArray(cellMargin) && typeof cellMargin === 'number') cellMargin = [cellMargin, cellMargin, cellMargin, cellMargin]
						/** FUTURE: DEPRECATED:
						 * - Backwards-Compat: Oops! Discovered we were still using points for cell margin before v3.8.0 (UGH!)
						 * - We cant introduce a breaking change before v4.0, so...
						 */
						let cellMarginXml = ''
						if (cellMargin[0] >= 1) {
							cellMarginXml = ` marL="${valToPts(cellMargin[3])}" marR="${valToPts(cellMargin[1])}" marT="${valToPts(cellMargin[0])}" marB="${valToPts(
								cellMargin[2]
							)}"`
						} else {
							cellMarginXml = ` marL="${inch2Emu(cellMargin[3])}" marR="${inch2Emu(cellMargin[1])}" marT="${inch2Emu(cellMargin[0])}" marB="${inch2Emu(
								cellMargin[2]
							)}"`
						}

						// FUTURE: Cell NOWRAP property (textwrap: add to a:tcPr (horzOverflow="overflow" or whatever options exist)

						// 4: Set CELL content and properties ==================================
						strXml += `<a:tc${cellSpanAttrStr}>${genXmlTextBody(cell)}<a:tcPr${cellMarginXml}${cellValign}${cellTextDir}>`
						// strXml += `<a:tc${cellColspan}${cellRowspan}>${genXmlTextBody(cell)}<a:tcPr${cellMarginXml}${cellValign}${cellTextDir}>`
						// FIXME: 20200525: ^^^
						// <a:tcPr marL="38100" marR="38100" marT="38100" marB="38100" vert="vert270">

						// 5: Borders: Add any borders
						if (cellOpts.border && Array.isArray(cellOpts.border)) {
							// NOTE: *** IMPORTANT! *** LRTB order matters! (Reorder a line below to watch the borders go wonky in MS-PPT-2013!!)
							[
								{ idx: 3, name: 'lnL' },
								{ idx: 1, name: 'lnR' },
								{ idx: 0, name: 'lnT' },
								{ idx: 2, name: 'lnB' },
							].forEach(obj => {
								if (cellOpts.border[obj.idx].type !== 'none') {
									strXml += `<a:${obj.name} w="${valToPts(cellOpts.border[obj.idx].pt)}" cap="flat" cmpd="sng" algn="ctr">`
									strXml += `<a:solidFill>${createColorElement(cellOpts.border[obj.idx].color)}</a:solidFill>`
									strXml += `<a:prstDash val="${cellOpts.border[obj.idx].type === 'dash' ? 'sysDash' : 'solid'
									}"/><a:round/><a:headEnd type="none" w="med" len="med"/><a:tailEnd type="none" w="med" len="med"/>`
									strXml += `</a:${obj.name}>`
								} else {
									strXml += `<a:${obj.name} w="0" cap="flat" cmpd="sng" algn="ctr"><a:noFill/></a:${obj.name}>`
								}
							})
						}

						// 6: Close cell Properties & Cell
						strXml += cellFill
						strXml += '  </a:tcPr>'
						strXml += ' </a:tc>'
					})

					// D: Complete row
					strXml += '</a:tr>'
				})

				// STEP 5: Complete table
				strXml += '      </a:tbl>'
				strXml += '    </a:graphicData>'
				strXml += '  </a:graphic>'
				strXml += '</p:graphicFrame>'

				// STEP 6: Set table XML
				strSlideXml += strXml

				// LAST: Increment counter
				intTableNum++
				break

			case SLIDE_OBJECT_TYPES.text:
			case SLIDE_OBJECT_TYPES.placeholder:
				// Lines can have zero cy, but text should not
				if (!slideItemObj.options.line && cy === 0) cy = EMU * 0.3

				// Margin/Padding/Inset for textboxes
				if (!slideItemObj.options._bodyProp) slideItemObj.options._bodyProp = {}
				if (slideItemObj.options.margin && Array.isArray(slideItemObj.options.margin)) {
					slideItemObj.options._bodyProp.lIns = valToPts(slideItemObj.options.margin[0] || 0)
					slideItemObj.options._bodyProp.rIns = valToPts(slideItemObj.options.margin[1] || 0)
					slideItemObj.options._bodyProp.bIns = valToPts(slideItemObj.options.margin[2] || 0)
					slideItemObj.options._bodyProp.tIns = valToPts(slideItemObj.options.margin[3] || 0)
				} else if (typeof slideItemObj.options.margin === 'number') {
					slideItemObj.options._bodyProp.lIns = valToPts(slideItemObj.options.margin)
					slideItemObj.options._bodyProp.rIns = valToPts(slideItemObj.options.margin)
					slideItemObj.options._bodyProp.bIns = valToPts(slideItemObj.options.margin)
					slideItemObj.options._bodyProp.tIns = valToPts(slideItemObj.options.margin)
				}

				// A: Start SHAPE =======================================================
				strSlideXml += '<p:sp>'

				// B: The addition of the "txBox" attribute is the sole determiner of if an object is a shape or textbox
				strSlideXml += `<p:nvSpPr><p:cNvPr id="${idx + 2}" name="${slideItemObj.options.objectName}">`
				// <Hyperlink>
				if (slideItemObj.options.hyperlink?.url) {
					strSlideXml += `<a:hlinkClick r:id="rId${slideItemObj.options.hyperlink._rId}" tooltip="${slideItemObj.options.hyperlink.tooltip ? encodeXmlEntities(slideItemObj.options.hyperlink.tooltip) : ''}"/>`
				}
				if (slideItemObj.options.hyperlink?.slide) {
					strSlideXml += `<a:hlinkClick r:id="rId${slideItemObj.options.hyperlink._rId}" tooltip="${slideItemObj.options.hyperlink.tooltip ? encodeXmlEntities(slideItemObj.options.hyperlink.tooltip) : ''}" action="ppaction://hlinksldjump"/>`
				}
				// </Hyperlink>
				strSlideXml += '</p:cNvPr>'
				strSlideXml += '<p:cNvSpPr' + (slideItemObj.options?.isTextBox ? ' txBox="1"/>' : '/>')
				strSlideXml += `<p:nvPr>${slideItemObj._type === 'placeholder' ? genXmlPlaceholder(slideItemObj) : genXmlPlaceholder(placeholderObj)}</p:nvPr>`
				strSlideXml += '</p:nvSpPr><p:spPr>'
				strSlideXml += `<a:xfrm${locationAttr}>`
				strSlideXml += `<a:off x="${x}" y="${y}"/>`
				strSlideXml += `<a:ext cx="${cx}" cy="${cy}"/></a:xfrm>`

				if (slideItemObj.shape === 'custGeom') {
					strSlideXml += '<a:custGeom><a:avLst />'
					strSlideXml += '<a:gdLst>'
					strSlideXml += '</a:gdLst>'
					strSlideXml += '<a:ahLst />'
					strSlideXml += '<a:cxnLst>'
					strSlideXml += '</a:cxnLst>'
					strSlideXml += '<a:rect l="l" t="t" r="r" b="b" />'

					strSlideXml += '<a:pathLst>'
					strSlideXml += `<a:path w="${cx}" h="${cy}">`

					slideItemObj.options.points?.forEach((point, i) => {
						if ('curve' in point) {
							switch (point.curve.type) {
								case 'arc':
									strSlideXml += `<a:arcTo hR="${getSmartParseNumber(point.curve.hR, 'Y', slide._presLayout)}" wR="${getSmartParseNumber(
										point.curve.wR,
										'X',
										slide._presLayout
									)}" stAng="${convertRotationDegrees(point.curve.stAng)}" swAng="${convertRotationDegrees(point.curve.swAng)}" />`
									break
								case 'cubic':
									strSlideXml += `<a:cubicBezTo>
									<a:pt x="${getSmartParseNumber(point.curve.x1, 'X', slide._presLayout)}" y="${getSmartParseNumber(point.curve.y1, 'Y', slide._presLayout)}" />
									<a:pt x="${getSmartParseNumber(point.curve.x2, 'X', slide._presLayout)}" y="${getSmartParseNumber(point.curve.y2, 'Y', slide._presLayout)}" />
									<a:pt x="${getSmartParseNumber(point.x, 'X', slide._presLayout)}" y="${getSmartParseNumber(point.y, 'Y', slide._presLayout)}" />
									</a:cubicBezTo>`
									break
								case 'quadratic':
									strSlideXml += `<a:quadBezTo>
									<a:pt x="${getSmartParseNumber(point.curve.x1, 'X', slide._presLayout)}" y="${getSmartParseNumber(point.curve.y1, 'Y', slide._presLayout)}" />
									<a:pt x="${getSmartParseNumber(point.x, 'X', slide._presLayout)}" y="${getSmartParseNumber(point.y, 'Y', slide._presLayout)}" />
									</a:quadBezTo>`
									break
								default:
									break
							}
						} else if ('close' in point) {
							strSlideXml += '<a:close />'
						} else if (point.moveTo || i === 0) {
							strSlideXml += `<a:moveTo><a:pt x="${getSmartParseNumber(point.x, 'X', slide._presLayout)}" y="${getSmartParseNumber(
								point.y,
								'Y',
								slide._presLayout
							)}" /></a:moveTo>`
						} else {
							strSlideXml += `<a:lnTo><a:pt x="${getSmartParseNumber(point.x, 'X', slide._presLayout)}" y="${getSmartParseNumber(
								point.y,
								'Y',
								slide._presLayout
							)}" /></a:lnTo>`
						}
					})

					strSlideXml += '</a:path>'
					strSlideXml += '</a:pathLst>'
					strSlideXml += '</a:custGeom>'
				} else {
					strSlideXml += '<a:prstGeom prst="' + slideItemObj.shape + '"><a:avLst>'
					if (slideItemObj.options.rectRadius) {
						strSlideXml += `<a:gd name="adj" fmla="val ${Math.round((slideItemObj.options.rectRadius * EMU * 100000) / Math.min(cx, cy))}"/>`
					} else if (slideItemObj.options.angleRange) {
						for (let i = 0; i < 2; i++) {
							const angle = slideItemObj.options.angleRange[i]
							strSlideXml += `<a:gd name="adj${i + 1}" fmla="val ${convertRotationDegrees(angle)}" />`
						}

						if (slideItemObj.options.arcThicknessRatio) {
							strSlideXml += `<a:gd name="adj3" fmla="val ${Math.round(slideItemObj.options.arcThicknessRatio * 50000)}" />`
						}
					}
					strSlideXml += '</a:avLst></a:prstGeom>'
				}

				// Option: FILL
				strSlideXml += slideItemObj.options.fill ? genXmlColorSelection(slideItemObj.options.fill) : '<a:noFill/>'

				// shape Type: LINE: line color
				if (slideItemObj.options.line) {
					strSlideXml += slideItemObj.options.line.width ? `<a:ln w="${valToPts(slideItemObj.options.line.width)}">` : '<a:ln>'
					if (slideItemObj.options.line.color) strSlideXml += genXmlColorSelection(slideItemObj.options.line)
					if (slideItemObj.options.line.dashType) strSlideXml += `<a:prstDash val="${slideItemObj.options.line.dashType}"/>`
					if (slideItemObj.options.line.beginArrowType) strSlideXml += `<a:headEnd type="${slideItemObj.options.line.beginArrowType}"/>`
					if (slideItemObj.options.line.endArrowType) strSlideXml += `<a:tailEnd type="${slideItemObj.options.line.endArrowType}"/>`
					// FUTURE: `endArrowSize` < a: headEnd type = "arrow" w = "lg" len = "lg" /> 'sm' | 'med' | 'lg'(values are 1 - 9, making a 3x3 grid of w / len possibilities)
					strSlideXml += '</a:ln>'
				}

				// EFFECTS > SHADOW: REF: @see http://officeopenxml.com/drwSp-effects.php
				strSlideXml += genShadowEffectXml(slideItemObj.options.shadow)

				/* TODO: FUTURE: Text wrapping (copied from MS-PPTX export)
					// Commented out b/c i'm not even sure this works - current code produces text that wraps in shapes and textboxes, so...
					if ( slideItemObj.options.textWrap ) {
						strSlideXml += '<a:extLst>'
									+ '<a:ext uri="{C572A759-6A51-4108-AA02-DFA0A04FC94B}">'
									+ '<ma14:wrappingTextBoxFlag xmlns:ma14="http://schemas.microsoft.com/office/mac/drawingml/2011/main" val="1"/>'
									+ '</a:ext>'
									+ '</a:extLst>';
					}
				*/

				// B: Close shape Properties
				strSlideXml += '</p:spPr>'

				// C: Add formatted text (text body "bodyPr")
				strSlideXml += genXmlTextBody(slideItemObj)

				// LAST: Close SHAPE =======================================================
				strSlideXml += '</p:sp>'
				break

			case SLIDE_OBJECT_TYPES.image:
				strSlideXml += '<p:pic>'
				strSlideXml += '  <p:nvPicPr>'
				strSlideXml += `<p:cNvPr id="${idx + 2}" name="${slideItemObj.options.objectName}" descr="${encodeXmlEntities(
					slideItemObj.options.altText || slideItemObj.image
				)}">`
				if (slideItemObj.hyperlink?.url) {
					strSlideXml += `<a:hlinkClick r:id="rId${slideItemObj.hyperlink._rId}" tooltip="${slideItemObj.hyperlink.tooltip ? encodeXmlEntities(slideItemObj.hyperlink.tooltip) : ''
					}"/>`
				}
				if (slideItemObj.hyperlink?.slide) {
					strSlideXml += `<a:hlinkClick r:id="rId${slideItemObj.hyperlink._rId}" tooltip="${slideItemObj.hyperlink.tooltip ? encodeXmlEntities(slideItemObj.hyperlink.tooltip) : ''
					}" action="ppaction://hlinksldjump"/>`
				}
				strSlideXml += '    </p:cNvPr>'
				strSlideXml += '    <p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>'
				strSlideXml += '    <p:nvPr>' + genXmlPlaceholder(placeholderObj) + '</p:nvPr>'
				strSlideXml += '  </p:nvPicPr>'
				strSlideXml += '<p:blipFill>'
				// NOTE: This works for both cases: either `path` or `data` contains the SVG
				if (
					(slide._relsMedia || []).filter(rel => rel.rId === slideItemObj.imageRid)[0] &&
					(slide._relsMedia || []).filter(rel => rel.rId === slideItemObj.imageRid)[0].extn === 'svg'
				) {
					strSlideXml += `<a:blip r:embed="rId${slideItemObj.imageRid - 1}">`
					strSlideXml += slideItemObj.options.transparency ? ` <a:alphaModFix amt="${Math.round((100 - slideItemObj.options.transparency) * 1000)}"/>` : ''
					strSlideXml += ' <a:extLst>'
					strSlideXml += '  <a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">'
					strSlideXml += `   <asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="rId${slideItemObj.imageRid}"/>`
					strSlideXml += '  </a:ext>'
					strSlideXml += ' </a:extLst>'
					strSlideXml += '</a:blip>'
				} else {
					strSlideXml += `<a:blip r:embed="rId${slideItemObj.imageRid}">`
					strSlideXml += slideItemObj.options.transparency ? `<a:alphaModFix amt="${Math.round((100 - slideItemObj.options.transparency) * 1000)}"/>` : ''
					strSlideXml += '</a:blip>'
				}
				if (sizing?.type) {
					const boxW = sizing.w ? getSmartParseNumber(sizing.w, 'X', slide._presLayout) : cx
					const boxH = sizing.h ? getSmartParseNumber(sizing.h, 'Y', slide._presLayout) : cy
					const boxX = getSmartParseNumber(sizing.x || 0, 'X', slide._presLayout)
					const boxY = getSmartParseNumber(sizing.y || 0, 'Y', slide._presLayout)

					strSlideXml += ImageSizingXml[sizing.type]({ w: imgWidth, h: imgHeight }, { w: boxW, h: boxH, x: boxX, y: boxY })
					imgWidth = boxW
					imgHeight = boxH
				} else {
					strSlideXml += '  <a:stretch><a:fillRect/></a:stretch>'
				}
				strSlideXml += '</p:blipFill>'
				strSlideXml += '<p:spPr>'
				strSlideXml += ' <a:xfrm' + locationAttr + '>'
				strSlideXml += `  <a:off x="${x}" y="${y}"/>`
				strSlideXml += `  <a:ext cx="${imgWidth}" cy="${imgHeight}"/>`
				strSlideXml += ' </a:xfrm>'
				strSlideXml += ` <a:prstGeom prst="${rounding ? 'ellipse' : 'rect'}"><a:avLst/></a:prstGeom>`

				// EFFECTS > SHADOW: REF: @see http://officeopenxml.com/drwSp-effects.php
				strSlideXml += genShadowEffectXml(slideItemObj.options.shadow)
				strSlideXml += '</p:spPr>'
				strSlideXml += '</p:pic>'
				break

			case SLIDE_OBJECT_TYPES.media:
				if (slideItemObj.mtype === 'online') {
					strSlideXml += '<p:pic>'
					strSlideXml += ' <p:nvPicPr>'
					// IMPORTANT: <p:cNvPr id="" value is critical - if its not the same number as preview image `rId`, PowerPoint throws error!
					strSlideXml += `<p:cNvPr id="${slideItemObj.mediaRid + 2}" name="${slideItemObj.options.objectName}"/>`
					strSlideXml += ' <p:cNvPicPr/>'
					strSlideXml += ' <p:nvPr>'
					strSlideXml += `  <a:videoFile r:link="rId${slideItemObj.mediaRid}"/>`
					strSlideXml += ' </p:nvPr>'
					strSlideXml += ' </p:nvPicPr>'
					// NOTE: `blip` is diferent than videos; also there's no preview "p:extLst" above but exists in videos
					strSlideXml += ` <p:blipFill><a:blip r:embed="rId${slideItemObj.mediaRid + 1}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` // NOTE: Preview image is required!
					strSlideXml += ' <p:spPr>'
					strSlideXml += `  <a:xfrm${locationAttr}><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>`
					strSlideXml += '  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
					strSlideXml += ' </p:spPr>'
					strSlideXml += '</p:pic>'
				} else {
					strSlideXml += '<p:pic>'
					strSlideXml += ' <p:nvPicPr>'
					// IMPORTANT: <p:cNvPr id="" value is critical - if not the same number as preiew image rId, PowerPoint throws error!
					strSlideXml += `<p:cNvPr id="${slideItemObj.mediaRid + 2}" name="${slideItemObj.options.objectName
					}"><a:hlinkClick r:id="" action="ppaction://media"/></p:cNvPr>`
					strSlideXml += ' <p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>'
					strSlideXml += ' <p:nvPr>'
					strSlideXml += `  <a:videoFile r:link="rId${slideItemObj.mediaRid}"/>`
					strSlideXml += '  <p:extLst>'
					strSlideXml += '   <p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}">'
					strSlideXml += `    <p14:media xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" r:embed="rId${slideItemObj.mediaRid + 1}"/>`
					strSlideXml += '   </p:ext>'
					strSlideXml += '  </p:extLst>'
					strSlideXml += ' </p:nvPr>'
					strSlideXml += ' </p:nvPicPr>'
					strSlideXml += ` <p:blipFill><a:blip r:embed="rId${slideItemObj.mediaRid + 2}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` // NOTE: Preview image is required!
					strSlideXml += ' <p:spPr>'
					strSlideXml += `  <a:xfrm${locationAttr}><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>`
					strSlideXml += '  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
					strSlideXml += ' </p:spPr>'
					strSlideXml += '</p:pic>'
				}
				break

			case SLIDE_OBJECT_TYPES.chart:
				strSlideXml += '<p:graphicFrame>'
				strSlideXml += ' <p:nvGraphicFramePr>'
				strSlideXml += `   <p:cNvPr id="${idx + 2}" name="${slideItemObj.options.objectName}" descr="${encodeXmlEntities(slideItemObj.options.altText || '')}"/>`
				strSlideXml += '   <p:cNvGraphicFramePr/>'
				strSlideXml += `   <p:nvPr>${genXmlPlaceholder(placeholderObj)}</p:nvPr>`
				strSlideXml += ' </p:nvGraphicFramePr>'
				strSlideXml += ` <p:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></p:xfrm>`
				strSlideXml += ' <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
				strSlideXml += '  <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
				strSlideXml += `   <c:chart r:id="rId${slideItemObj.chartRid}" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>`
				strSlideXml += '  </a:graphicData>'
				strSlideXml += ' </a:graphic>'
				strSlideXml += '</p:graphicFrame>'
				break

			default:
				strSlideXml += ''
				break
		}
	})

	// STEP 4: Add slide numbers (if any) last
	if (slide._slideNumberProps) {
		// Set some defaults (done here b/c SlideNumber canbe added to masters or slides and has numerous entry points)
		if (!slide._slideNumberProps.align) slide._slideNumberProps.align = 'left'

		strSlideXml += '<p:sp>'
		strSlideXml += ' <p:nvSpPr>'
		strSlideXml += '  <p:cNvPr id="25" name="Slide Number Placeholder 0"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
		strSlideXml += '  <p:nvPr><p:ph type="sldNum" sz="quarter" idx="4294967295"/></p:nvPr>'
		strSlideXml += ' </p:nvSpPr>'
		strSlideXml += ' <p:spPr>'
		strSlideXml += '<a:xfrm>' +
			`<a:off x="${getSmartParseNumber(slide._slideNumberProps.x, 'X', slide._presLayout)}" y="${getSmartParseNumber(slide._slideNumberProps.y, 'Y', slide._presLayout)}"/>` +
			`<a:ext cx="${slide._slideNumberProps.w ? getSmartParseNumber(slide._slideNumberProps.w, 'X', slide._presLayout) : '800000'}" cy="${slide._slideNumberProps.h ? getSmartParseNumber(slide._slideNumberProps.h, 'Y', slide._presLayout) : '300000'}"/>` +
			'</a:xfrm>' +
			' <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
			' <a:extLst><a:ext uri="{C572A759-6A51-4108-AA02-DFA0A04FC94B}"><ma14:wrappingTextBoxFlag val="0" xmlns:ma14="http://schemas.microsoft.com/office/mac/drawingml/2011/main"/></a:ext></a:extLst>' +
			'</p:spPr>'
		strSlideXml += '<p:txBody>'
		strSlideXml += '<a:bodyPr'
		if (slide._slideNumberProps.margin && Array.isArray(slide._slideNumberProps.margin)) {
			strSlideXml += ` lIns="${valToPts(slide._slideNumberProps.margin[3] || 0)}"`
			strSlideXml += ` tIns="${valToPts(slide._slideNumberProps.margin[0] || 0)}"`
			strSlideXml += ` rIns="${valToPts(slide._slideNumberProps.margin[1] || 0)}"`
			strSlideXml += ` bIns="${valToPts(slide._slideNumberProps.margin[2] || 0)}"`
		} else if (typeof slide._slideNumberProps.margin === 'number') {
			strSlideXml += ` lIns="${valToPts(slide._slideNumberProps.margin || 0)}"`
			strSlideXml += ` tIns="${valToPts(slide._slideNumberProps.margin || 0)}"`
			strSlideXml += ` rIns="${valToPts(slide._slideNumberProps.margin || 0)}"`
			strSlideXml += ` bIns="${valToPts(slide._slideNumberProps.margin || 0)}"`
		}
		if (slide._slideNumberProps.valign) {
			strSlideXml += ` anchor="${slide._slideNumberProps.valign.replace('top', 't').replace('middle', 'ctr').replace('bottom', 'b')}"`
		}
		strSlideXml += '/>'
		strSlideXml += '  <a:lstStyle><a:lvl1pPr>'
		if (slide._slideNumberProps.fontFace || slide._slideNumberProps.fontSize || slide._slideNumberProps.color) {
			strSlideXml += `<a:defRPr sz="${Math.round((slide._slideNumberProps.fontSize || 12) * 100)}">`
			if (slide._slideNumberProps.color) strSlideXml += genXmlColorSelection(slide._slideNumberProps.color)
			if (slide._slideNumberProps.fontFace) { strSlideXml += `<a:latin typeface="${slide._slideNumberProps.fontFace}"/><a:ea typeface="${slide._slideNumberProps.fontFace}"/><a:cs typeface="${slide._slideNumberProps.fontFace}"/>` }
			strSlideXml += '</a:defRPr>'
		}
		strSlideXml += '</a:lvl1pPr></a:lstStyle>'
		strSlideXml += '<a:p>'
		if (slide._slideNumberProps.align.startsWith('l')) strSlideXml += '<a:pPr algn="l"/>'
		else if (slide._slideNumberProps.align.startsWith('c')) strSlideXml += '<a:pPr algn="ctr"/>'
		else if (slide._slideNumberProps.align.startsWith('r')) strSlideXml += '<a:pPr algn="r"/>'
		else strSlideXml += '<a:pPr algn="l"/>'
		strSlideXml += `<a:fld id="${SLDNUMFLDID}" type="slidenum"><a:rPr b="${slide._slideNumberProps.bold ? 1 : 0}" lang="en-US"/>`
		strSlideXml += `<a:t>${slide._slideNum}</a:t></a:fld><a:endParaRPr lang="en-US"/></a:p>`
		strSlideXml += '</p:txBody></p:sp>'
	}

	// STEP 5: Close spTree and finalize slide XML
	strSlideXml += '</p:spTree>'
	strSlideXml += '</p:cSld>'

	// LAST: Return
	return strSlideXml
}
