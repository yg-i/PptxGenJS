/**
 * "The Activist Solution (cont.)" - attempting with NEW API
 *
 * FRICTION POINTS DISCOVERED:
 *
 * 1. INLINE STYLED TEXT - The slide has "By **committing** to a choice..." where
 *    "committing" is both bold AND a different color (cyan). Current Text() element
 *    doesn't support inline mixed styles. Need something like:
 *
 *    Text("By **committing** to a choice...", { boldColor: '4FC3F7' })
 *
 *    Or a richer syntax:
 *
 *    RichText([
 *      "By ",
 *      { text: "committing", color: '4FC3F7', bold: true },
 *      " to a choice situation..."
 *    ])
 *
 * 2. GRADIENT BACKGROUND - Need to use gradient() helper, which works but the
 *    slide config syntax is a bit verbose.
 *
 * 3. TEXT HEIGHT AUTO-CALCULATION - For multi-line paragraphs, I have to manually
 *    estimate h based on fontSize and line count. Should be automatic.
 *
 * 4. PARAGRAPH SPACING - No built-in way to say "these are 3 paragraphs with
 *    consistent spacing". Stack works but requires explicit h for each.
 */

import {
	$pptxNewAPI,
	Text,
	Stack,
	gradient,
} from '../dist/pptxgen.es.js'

const CYAN = '4FC3F7'
const WHITE = 'FFFFFF'
const GRAY = 'B0B8C4'

const $pptx = new $pptxNewAPI({ layout: '16x9' })

// ATTEMPT 1: Using plain Text() - loses the inline color highlighting
// This produces the structure but NOT the cyan-colored keywords

$pptx.slide(
	{
		background: gradient({ from: '0A0F1A', to: '152238', angle: 180 }),
		defaults: { fontFace: 'Outfit', color: WHITE },
	},

	// Title
	Text('The Activist Solution (cont.)', {
		fontSize: 40,
		fontFace: 'Outfit ExtraBold',
		bold: true,
		x: 0.7, y: 0.5, w: 9, h: 0.9,
	}),

	// Body paragraphs - using Stack for vertical flow
	// PROBLEM: Can't do inline colored text like "By **committing**..."
	Stack({ x: 0.7, y: 1.5, w: 9, gap: 0.4 }, [
		// Paragraph 1 - ideally would be:
		// Text("By **committing** to a choice situation, the agent is **justified** in being in it.", { boldColor: CYAN })
		Text('By committing to a choice situation, the agent is justified in being in it.', {
			fontSize: 26,
			fontFace: 'Outfit SemiBold',
			h: 1.0,
		}),

		// Paragraph 2
		Text('She makes being in that choice situation better for her with respect to having a meaningful life.', {
			fontSize: 26,
			fontFace: 'Outfit SemiBold',
			color: GRAY,
			h: 1.0,
		}),

		// Paragraph 3
		Text('She thus has meaning in her life in virtue of her capacity to make it true through creating reasons that she has most reason to be in one choice situation rather than another.', {
			fontSize: 26,
			fontFace: 'Outfit SemiBold',
			color: GRAY,
			h: 1.5,
		}),
	])
)

await $pptx.write('lab/activist-solution-new-api.pptx')
console.log('Created: lab/activist-solution-new-api.pptx')

/*
 * =========================================================================
 * IDEAL API FOR THIS SLIDE
 * =========================================================================
 *
 * $pptx.slide(
 *   {
 *     background: gradient({ from: '0A0F1A', to: '152238' }),
 *     defaults: { fontFace: 'Outfit SemiBold', color: WHITE, accentColor: CYAN },
 *   },
 *
 *   Text('The Activist Solution (cont.)', { fontSize: 40, bold: true }),
 *
 *   // Option A: Markdown with accentColor for **bold**
 *   Paragraphs({ fontSize: 26, gap: 0.4 }, [
 *     "By **committing** to a choice situation, the agent is **justified** in being in it.",
 *     { text: "She **makes** being in that choice...", color: GRAY },
 *     { text: "She thus has meaning... **make it true** through...", color: GRAY },
 *   ])
 *
 *   // Option B: Tagged template literal
 *   Paragraphs({ fontSize: 26, gap: 0.4 }, [
 *     md`By ${accent('committing')} to a choice situation, the agent is ${accent('justified')} in being in it.`,
 *     md`She ${accent('makes')} being in that choice...`.gray(),
 *   ])
 * )
 *
 * KEY INSIGHTS:
 * - Need a way to mix colors WITHIN a single text run
 * - Paragraphs() element for multiple text blocks with consistent styling
 * - Auto-height based on content + width
 * - accentColor should apply to **bold** text automatically
 */
