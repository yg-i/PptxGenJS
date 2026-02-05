/**
 * "From Trichotomy to Quadchotomy" - using the NEW declarative API
 */

import {
	$pptxNewAPI,
	Text,
	Pill,
	Stack,
	Columns,
	FadeIn,
} from '../dist/pptxgen.es.js'

// Colors
const DARK_TEXT = '2D3436'
const GREEN = '2ECC71'
const BLUE = '3498DB'
const ORANGE = 'E67E22'
const TEAL = '1ABC9C'
const YELLOW_ORANGE = 'F5A623'
const PINK = 'E84A5F'

// Create presentation
const $pptx = new $pptxNewAPI({ layout: '16x9' })

$pptx.slide(
	{ defaults: { fontFace: 'Outfit SemiBold', color: DARK_TEXT } },

	// Title
	Text('From Trichotomy to Quadchotomy', {
		fontSize: 32,
		fontFace: 'Outfit ExtraBold',
		x: 0.5, y: 0.4, w: 9, h: 0.7,
	}),

	// Two columns
	Columns({ x: 0.5, y: 1.3, w: 9, gap: 0.4, ratio: [1, 1] }, [

		// Left column - header outside FadeIn, pills inside
		Stack({ gap: 0.15 }, [
			Text('Traditional View', { fontSize: 18, align: 'center', h: 0.5 }),
			FadeIn({ onClick: true, stagger: 150 },
				Stack({ gap: 0.15 }, [
					Pill({ text: 'Equally good', fill: GREEN }),
					Pill({ text: 'Better than', fill: BLUE }),
					Pill({ text: 'Worse than', fill: ORANGE }),
				])
			),
		]),

		// Right column - header outside FadeIn, pills inside
		Stack({ gap: 0.15 }, [
			Text('Parity View', { fontSize: 18, align: 'center', color: TEAL, h: 0.5 }),
			FadeIn({ onClick: true, stagger: 150 },
				Stack({ gap: 0.15 }, [
					Pill({ text: 'Equally good', fill: TEAL }),
					Pill({ text: 'Better than', fill: TEAL }),
					Pill({ text: 'On a par', fill: YELLOW_ORANGE }),
					Pill({ text: 'Worse than', fill: PINK }),
				])
			),
		]),
	])
)

await $pptx.write('lab/trichotomy-new-api.pptx')
console.log('Created: lab/trichotomy-new-api.pptx')
