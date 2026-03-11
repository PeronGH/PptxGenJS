/**
 * PptxGenJS: Media Methods
 */

import { IMG_BROKEN } from "./core-enums.ts";
import { PresSlide, SlideLayout } from "./core-interfaces.ts";
import { readPathAsBase64 } from "./runtime.ts";

/**
 * Encode Image/Audio/Video into base64
 * @param {PresSlide | SlideLayout} layout - slide layout
 * @return {Promise} promise
 */
export function encodeSlideMediaRels(
  layout: PresSlide | SlideLayout,
): Array<Promise<string>> {
  // STEP 1: Prepare promises list
  const imageProms: Array<Promise<string>> = [];

  // A: Capture all audio/image/video candidates for encoding (filtering online/pre-encoded)
  const candidateRels = layout._relsMedia.filter(
    (rel) =>
      rel.type !== "online" && !rel.data &&
      (!rel.path || (rel.path && !rel.path.includes("preencoded"))),
  );

  // B: PERF: Mark dupes (same `path`) to avoid loading the same media over-and-over!
  const unqPaths: string[] = [];
  candidateRels.forEach((rel) => {
    if (!unqPaths.includes(rel.path)) {
      rel.isDuplicate = false;
      unqPaths.push(rel.path);
    } else {
      rel.isDuplicate = true;
    }
  });

  // STEP 2: Read/Encode each unique media item
  candidateRels
    .filter((rel) => !rel.isDuplicate)
    .forEach((rel) => {
      imageProms.push(
        (async () => {
          if (rel.isSvgPng) {
            rel.data = IMG_BROKEN;
          } else {
            try {
              rel.data = await readPathAsBase64(rel.path);
            } catch (ex) {
              rel.data = IMG_BROKEN;
              candidateRels
                .filter((dupe) => dupe.isDuplicate && dupe.path === rel.path)
                .forEach((dupe) => (dupe.data = rel.data));
              throw new Error(
                `ERROR: Unable to read media: "${rel.path}"\n${String(ex)}`,
              );
            }
          }

          candidateRels
            .filter((dupe) => dupe.isDuplicate && dupe.path === rel.path)
            .forEach((dupe) => (dupe.data = rel.data));
          return "done";
        })(),
      );
    });

  return imageProms;
}

/**
 * FIXME: TODO: currently unused
 * TODO: Should return a Promise
 */
/*
function getSizeFromImage (inImgUrl: string): { width: number, height: number } {
	const sizeOf = typeof require !== 'undefined' ? require('sizeof') : null // NodeJS

	if (sizeOf) {
		try {
			const dimensions = sizeOf(inImgUrl)
			return { width: dimensions.width, height: dimensions.height }
		} catch (ex) {
			console.error('ERROR: sizeOf: Unable to load image: ' + inImgUrl)
			return { width: 0, height: 0 }
		}
	} else if (Image && typeof Image === 'function') {
		// A: Create
		const image = new Image()

		// B: Set onload event
		image.onload = () => {
			// FIRST: Check for any errors: This is the best method (try/catch wont work, etc.)
			if (image.width + image.height === 0) {
				return { width: 0, height: 0 }
			}
			const obj = { width: image.width, height: image.height }
			return obj
		}
		image.onerror = () => {
			console.error(`ERROR: image.onload: Unable to load image: ${inImgUrl}`)
		}

		// C: Load image
		image.src = inImgUrl
	}
}
*/
