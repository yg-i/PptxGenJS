/**
 * "The Activist Solution (cont.)" - using chalk-like styled text!
 */

import {
	$pptxNewAPI,
	Text,
	Stack,
	gradient,
	// Chalk-like style builders
	bold,
	cyan,
	gray,
} from '../dist/pptxgen.es.js'

// Combine styles: bold + cyan
const accent = bold.cyan

const $pptx = new $pptxNewAPI({ layout: '16x9' })

$pptx.slide(
	{
		background: gradient({ from: '0A0F1A', to: '152238', angle: 180 }),
		defaults: { fontFace: 'Outfit SemiBold', color: 'FFFFFF' },
	},

	// Title
	Text('The Activist Solution (cont.)', {
		fontSize: 40,
		fontFace: 'Outfit ExtraBold',
		bold: true,
		x: 0.7, y: 0.5, w: 9, h: 0.9,
	}),

	// Body paragraphs with inline styled text!
	Stack({ x: 0.7, y: 1.5, w: 9, gap: 0.3 }, [

		// Paragraph 1 - white with cyan accents
		Text`By ${accent("committing")} to a choice situation, the agent is ${accent("justified")} in being in it.`
			.config({ fontSize: 26, h: 1.0 }),

		// Paragraph 2 - gray with cyan accent
		Text`She ${accent("makes")} being in that choice situation better for her with respect to having a meaningful life.`
			.config({ fontSize: 26, color: 'B0B8C4', h: 1.0 }),

		// Paragraph 3 - gray with cyan accent
		Text`She thus has meaning in her life in virtue of her capacity to ${accent("make it true")} through creating reasons that she has most reason to be in one choice situation rather than another.`
			.config({ fontSize: 26, color: 'B0B8C4', h: 1.6 }),
	])
)

await $pptx.write('lab/activist-solution-chalk.pptx')
console.log('Created: lab/activist-solution-chalk.pptx')
