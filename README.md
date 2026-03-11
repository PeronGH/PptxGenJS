# PptxGenJS Deno Fork

This fork is a Deno-only PowerPoint generator. It keeps the core OOXML
generation logic from the original project and strips out the Node/browser
packaging layers that made Deno support awkward.

## What Changed

- `mod.ts` is the entrypoint.
- `write()` returns `Uint8Array`.
- `writeFile()` writes directly with `Deno.writeFile`.
- Source files import each other with explicit `.ts` extensions.
- HTML table scraping and browser download helpers are removed.
- SVG image preview support is removed.

## Usage

```ts
import PptxGenJS from "./mod.ts";

const pptx = new PptxGenJS();
const slide = pptx.addSlide();

slide.addText("Hello from Deno", {
  x: 1,
  y: 1,
  w: 4,
  h: 0.6,
  fontSize: 24,
});

await pptx.writeFile({ fileName: "hello-deno.pptx" });
```

If you want bytes instead of a file:

```ts
const bytes = await pptx.write({ compression: true });
```

## Tasks

```bash
deno task check
deno task lint
deno task test
deno task test:integration
deno task example
```

## Permissions

- `write()` needs no extra permissions beyond whatever your own code needs.
- `writeFile()` needs `--allow-write`.
- Local image/media paths need `--allow-read`.
- Remote image/media URLs need `--allow-net`.

`deno task test` runs the default local suite. `deno task test:integration` runs
the real remote-media integration test against an actual remote PNG URL.

## Unsupported In This Fork

- `tableToSlides()`
- Browser download flows
- SVG image input

## Example

`examples/basic.ts` writes `example.pptx` in the current directory.
