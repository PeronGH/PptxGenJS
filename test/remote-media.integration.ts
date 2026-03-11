import { assert, assertEquals } from "@std/assert";
import JSZip from "jszip";
import PptxGenJS from "../mod.ts";

const REMOTE_PNG_URL =
  "https://raw.githubusercontent.com/gitbrent/PptxGenJS/master/demos/common/images/logo_square_25.png";

Deno.test("remote media URLs are embedded into the pptx", async () => {
  const expectedResponse = await fetch(REMOTE_PNG_URL);
  assert(expectedResponse.ok);
  const expectedBytes = new Uint8Array(await expectedResponse.arrayBuffer());

  const pptx = new PptxGenJS();
  const slide = pptx.addSlide();
  slide.addImage({
    path: REMOTE_PNG_URL,
    x: 1,
    y: 1,
    w: 1,
    h: 1,
  });

  const bytes = await pptx.write();
  const zip = await JSZip.loadAsync(bytes);
  const mediaBytes = await zip.file("ppt/media/image-1-1.png")?.async(
    "uint8array",
  );

  assert(mediaBytes);
  assertEquals(mediaBytes.length, expectedBytes.length);
});
