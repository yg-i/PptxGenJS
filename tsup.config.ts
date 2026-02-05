import { defineConfig } from 'tsup'

export default defineConfig({
	entry: ['src/pptxgen.ts'],
	format: ['cjs', 'esm', 'iife'],
	outDir: 'dist',
	dts: true,
	clean: true,
	sourcemap: true,
	minify: false, // Keep readable for debugging; use minify: true for production
	globalName: 'PptxGenJS',
	external: ['jszip'],
	// For IIFE build, provide global for jszip
	esbuildOptions(options, context) {
		if (context.format === 'iife') {
			options.globalName = 'PptxGenJS'
			options.footer = {
				js: 'if (typeof module !== "undefined") module.exports = PptxGenJS;'
			}
		}
	},
	// Rename outputs for backward compatibility
	outExtension({ format }) {
		switch (format) {
			case 'cjs':
				return { js: '.cjs.js' }
			case 'esm':
				return { js: '.es.js' }
			case 'iife':
				return { js: '.bundle.js' }
			default:
				return { js: '.js' }
		}
	},
})
