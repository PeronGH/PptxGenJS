import assert from 'node:assert/strict'
import test from 'node:test'
import { createRequire } from 'node:module'

const require = createRequire(import.meta.url)
const JSZip = require('jszip')
const pptxgen = require('../dist/pptxgen.cjs.js')

async function getSlideXml(transition) {
	const pres = new pptxgen()
	const slide = pres.addSlide()
	slide.addText('Transition Test', { x: 1, y: 1, w: 2, h: 0.5 })
	slide.transition = transition

	const pptx = await pres.write('nodebuffer')
	const zip = await JSZip.loadAsync(pptx)

	return zip.file('ppt/slides/slide1.xml').async('string')
}

test('serializes ISO slide transitions with JS-friendly option names', async () => {
	const cases = [
		{
			name: 'fade through black with common attributes',
			transition: { type: 'fade', throughBlack: true, speed: 'slow', advanceOnClick: false, advanceTime: 3000 },
			expected: '<p:transition spd="slow" advClick="0" advTm="3000"><p:fade thruBlk="1"/></p:transition>',
		},
		{
			name: 'orientation transition',
			transition: { type: 'blinds', orientation: 'vertical' },
			expected: '<p:transition><p:blinds dir="vert"/></p:transition>',
		},
		{
			name: 'side direction transition',
			transition: { type: 'push', direction: 'left' },
			expected: '<p:transition><p:push dir="l"/></p:transition>',
		},
		{
			name: 'eight direction transition',
			transition: { type: 'cover', direction: 'rightDown' },
			expected: '<p:transition><p:cover dir="rd"/></p:transition>',
		},
		{
			name: 'split transition',
			transition: { type: 'split', orientation: 'vertical', direction: 'in' },
			expected: '<p:transition><p:split orient="vert" dir="in"/></p:transition>',
		},
		{
			name: 'strips transition',
			transition: { type: 'strips', direction: 'leftUp' },
			expected: '<p:transition><p:strips dir="lu"/></p:transition>',
		},
		{
			name: 'wheel transition',
			transition: { type: 'wheel', spokes: 8 },
			expected: '<p:transition><p:wheel spokes="8"/></p:transition>',
		},
		{
			name: 'zoom transition',
			transition: { type: 'zoom', direction: 'in' },
			expected: '<p:transition><p:zoom dir="in"/></p:transition>',
		},
		{
			name: 'empty transition',
			transition: { type: 'random' },
			expected: '<p:transition><p:random/></p:transition>',
		},
	]

	for (const testCase of cases) {
		const slideXml = await getSlideXml(testCase.transition)
		assert.ok(slideXml.includes(testCase.expected), testCase.name)
		assert.ok(
			slideXml.includes(`<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>${testCase.expected}`),
			`${testCase.name} is emitted after clrMapOvr`
		)
	}
})

test('serializes duration using AlternateContent fallback', async () => {
	const slideXml = await getSlideXml({
		type: 'split',
		orientation: 'horizontal',
		direction: 'out',
		speed: 'slow',
		duration: 1250,
		advanceTime: 3000,
	})

	assert.ok(slideXml.includes('<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">'))
	assert.ok(slideXml.includes('<mc:Choice Requires="p14"><p:transition spd="slow" p14:dur="1250" advTm="3000"><p:split orient="horz" dir="out"/></p:transition></mc:Choice>'))
	assert.ok(slideXml.includes('<mc:Fallback><p:transition spd="slow" advTm="3000"><p:split orient="horz" dir="out"/></p:transition></mc:Fallback>'))
})

test('throws for invalid transition configs', () => {
	const pres = new pptxgen()
	const slide = pres.addSlide()

	assert.throws(() => {
		slide.transition = { type: 'push', direction: 'diagonal' }
	}, /slide\.transition\.direction must be one of: left, right, up, down/)

	assert.throws(() => {
		slide.transition = { type: 'fade', dir: 'l' }
	}, /does not support "dir"/)

	assert.throws(() => {
		slide.transition = { type: 'random', advanceTime: -1 }
	}, /slide\.transition\.advanceTime must be a non-negative integer/)
})
