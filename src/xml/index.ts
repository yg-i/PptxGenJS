/**
 * XML utilities for PptxGenJS
 *
 * This module provides:
 * - OOXML namespace constants (namespaces.ts)
 * - Fluent XML builder API (builder.ts)
 */

// Re-export all namespace constants
export {
	// Namespaces
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
	NS_DCMITYPE,
	NS_XSI,
	NS_RELATIONSHIPS,
	NS_CONTENT_TYPES,
	NS_EXTENDED_PROPERTIES,
	NS_VT,
	NS_ASVG,
	NS_MA14,
	NS_P15,
	NS_THM15,
	NS_MS_OFFICE_2007_RELATIONSHIPS,
	// Relationship types
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
} from './namespaces'

// Re-export XML builder
export { XmlBuilder, xml, OoxmlElements } from './builder'
