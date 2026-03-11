import PptxGenJS from "../mod.ts";

const pptx = new PptxGenJS();
const slide = pptx.addSlide();

slide.addText("Hello from Deno", { x: 1, y: 1, w: 4, h: 0.6, fontSize: 24 });

const fileName = await pptx.writeFile({ fileName: "example.pptx" });

console.log(`Wrote ${fileName}`);
