import { assert, assertEquals, assertRejects } from "@std/assert";
import JSZip from "jszip";
import PptxGenJS from "../mod.ts";

const PNG_BASE64 =
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Wn0K5sAAAAASUVORK5CYII=";

function decodeBase64(base64: string): Uint8Array {
  const binary = atob(base64);
  return Uint8Array.from(binary, (byte) => byte.charCodeAt(0));
}

async function createTempPng(): Promise<string> {
  const filePath = await Deno.makeTempFile({ suffix: ".png" });
  await Deno.writeFile(filePath, decodeBase64(PNG_BASE64));
  return filePath;
}

Deno.test("write returns a non-empty pptx as Uint8Array", async () => {
  const pptx = new PptxGenJS();
  const slide = pptx.addSlide();

  slide.addText("Transition Test", { x: 1, y: 1, w: 2, h: 0.5 });
  slide.transition = {
    type: "split",
    orientation: "vertical",
    direction: "in",
    duration: 1250,
    advanceTime: 3000,
  };

  const bytes = await pptx.write();
  assert(bytes instanceof Uint8Array);
  assert(bytes.length > 0);

  const zip = await JSZip.loadAsync(bytes);
  const slideXml = await zip.file("ppt/slides/slide1.xml")?.async("string");

  assert(slideXml);
  assert(slideXml.includes("<mc:AlternateContent"));
  assert(slideXml.includes('<p:split orient="vert" dir="in"/>'));
});

Deno.test("writeFile writes a pptx to disk", async () => {
  const cwd = Deno.cwd();
  const tempDir = await Deno.makeTempDir();

  try {
    Deno.chdir(tempDir);

    const pptx = new PptxGenJS();
    pptx.addSlide().addText("Disk output", { x: 1, y: 1, w: 2, h: 0.5 });

    const fileName = await pptx.writeFile({ fileName: "disk-output" });
    assertEquals(fileName, "disk-output.pptx");

    const stats = await Deno.stat(`${tempDir}/${fileName}`);
    assert(stats.isFile);
    assert((stats.size ?? 0) > 0);
  } finally {
    Deno.chdir(cwd);
    await Deno.remove(tempDir, { recursive: true });
  }
});

Deno.test("local media paths are embedded into the pptx", async () => {
  const pngPath = await createTempPng();

  try {
    const pptx = new PptxGenJS();
    const slide = pptx.addSlide();
    slide.addImage({ path: pngPath, x: 1, y: 1, w: 1, h: 1 });

    const bytes = await pptx.write();
    const zip = await JSZip.loadAsync(bytes);
    const mediaBytes = await zip.file("ppt/media/image-1-1.png")?.async(
      "uint8array",
    );

    assert(mediaBytes);
    assertEquals(mediaBytes.length, decodeBase64(PNG_BASE64).length);
  } finally {
    await Deno.remove(pngPath);
  }
});

Deno.test("svg inputs are rejected explicitly", async () => {
  const pptx = new PptxGenJS();
  const slide = pptx.addSlide();

  await assertRejects(
    async () => {
      slide.addImage({
        data: "data:image/svg+xml;base64,PHN2Zy8+",
        x: 1,
        y: 1,
        w: 1,
        h: 1,
      });
      await pptx.write();
    },
    Error,
    "SVG images are not supported",
  );
});
