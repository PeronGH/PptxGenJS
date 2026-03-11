/**
 * NAME: demo_transition.mjs
 * AUTH: Brent Ely (https://github.com/gitbrent/)
 * DESC: Common test/demo slides for slide transitions
 * DEPS: Used by various demos (./demos/browser, ./demos/node, etc.)
 * VER.: 4.0.1
 * BLD.: 20260311
 */

import { BASE_TABLE_OPTS, BASE_TEXT_OPTS_L, BASE_TEXT_OPTS_R } from "./enums.mjs";

export function genSlides_Transition(pptx) {
	pptx.addSection({ title: "Transitions" });

	genSlide01(pptx);
	genSlide02(pptx);
	genSlide03(pptx);
}

function genSlide01(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Transitions" });
	slide.transition = { type: "fade", throughBlack: true, speed: "slow" };
	slide.addTable([[{ text: "Transition Examples: Fade", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);
	slide.addText("This slide uses a slow fade transition through black.", {
		x: 0.7, y: 1.1, w: 5.4, h: 0.6, fontSize: 20, color: "2F5597", bold: true,
	});
	slide.addText("Open the exported deck in PowerPoint or Keynote and run the slide show to preview the transition.", {
		x: 0.7, y: 1.9, w: 10.8, h: 0.8, fontSize: 18, color: "404040",
	});
	slide.addNotes("API Docs: slide.transition = { type: 'fade', throughBlack: true, speed: 'slow' }");
}

function genSlide02(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Transitions" });
	slide.transition = { type: "push", direction: "left", speed: "fast", advanceOnClick: false, advanceTime: 2500 };
	slide.addTable([[{ text: "Transition Examples: Push", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);
	slide.addText("This slide auto-advances with a fast push from the left.", {
		x: 0.7, y: 1.1, w: 6.2, h: 0.6, fontSize: 20, color: "C00000", bold: true,
	});
	slide.addText("slide.transition = { type: 'push', direction: 'left', speed: 'fast', advanceOnClick: false, advanceTime: 2500 }", {
		x: 0.7, y: 1.9, w: 11.2, h: 1.1, fontFace: "Courier New", fontSize: 15, color: "404040", margin: 0.08,
	});
}

function genSlide03(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Transitions" });
	slide.transition = { type: "split", orientation: "vertical", direction: "in", speed: "med", duration: 1200 };
	slide.addTable([[{ text: "Transition Examples: Split + Duration", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);
	slide.addText("This slide uses an Office 2010 duration extension with an AlternateContent fallback.", {
		x: 0.7, y: 1.1, w: 11.2, h: 0.7, fontSize: 19, color: "548235", bold: true,
	});
	slide.addText("Clients that support p14:dur use the 1200ms duration. Older clients fall back to the same split transition without the duration attribute.", {
		x: 0.7, y: 1.95, w: 11.2, h: 1.1, fontSize: 17, color: "404040",
	});
}
