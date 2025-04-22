import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import type { CallToolResult } from "@modelcontextprotocol/sdk/types.js";
import { $, Glob } from "bun";
import * as fs from "fs/promises";
import * as path from "path";
import PptxGenJS from "pptxgenjs";
import { z } from "zod";

/**
 * Environment
 */

const envSchema = z.object({
  STORAGE_DIR: z.string().min(1, "STORAGE_DIR is required"),
});
const env = envSchema.parse(process.env);
const STORAGE_DIR = path.resolve(env.STORAGE_DIR);

/**
 * State Schema
 */

const fontSizeSchema = z.number().min(8).max(256);
const backgroundSchema = z.object({
  color: z.string().optional(),
  transparency: z.number().min(0).max(100).optional(),
});
const slideNumberSchema = z.object({
  x: z.number(),
  y: z.number(),
  color: z.string().optional(),
  fontFace: z.string().optional(),
  fontSize: fontSizeSchema.optional(),
});
const textUnderlineStyleSchema = z.enum([
  "dash",
  "dashHeavy",
  "dashLong",
  "dashLongHeavy",
  "dbl",
  "dotDash",
  "dotDashHeave",
  "dotDotDash",
  "dotDotDashHeavy",
  "dotted",
  "dottedHeavy",
  "heavy",
  "none",
  "sng",
  "wavy",
  "wavyDbl",
  "wavyHeavy",
]);
const textUnderlineSchema = z.object({
  style: textUnderlineStyleSchema.optional(),
  color: z.string().optional(),
});
const textAlignSchema = z.enum(["left", "center", "right"]);
const textValignSchema = z.enum(["top", "middle", "bottom"]);
const textBulletTypeSchema = z.enum(["number", "bullet"]);
const textBulletNumberTypeSchema = z.enum([
  "alphaLcParenBoth",
  "alphaLcParenR",
  "alphaLcPeriod",
  "alphaUcParenBoth",
  "alphaUcParenR",
  "alphaUcPeriod",
  "arabicParenBoth",
  "arabicParenR",
  "arabicPeriod",
  "arabicPlain",
  "romanLcParenBoth",
  "romanLcParenR",
  "romanLcPeriod",
  "romanUcParenBoth",
  "romanUcParenR",
  "romanUcPeriod",
]);
const textBulletDetailSchema = z.object({
  type: textBulletTypeSchema.optional(),
  characterCode: z.string().optional(),
  indent: z.number().optional(),
  numberType: textBulletNumberTypeSchema.optional(),
  style: z.string().optional(),
});
const textBulletSchema = z.union([z.boolean(), textBulletDetailSchema]);
const textFillSchema = z.object({
  color: z.string(),
});
const textHyperlinkSchema = z.object({
  url: z.string(),
  slide: z.number().optional(),
  tooltip: z.string().optional(),
});
const textOptionsSchema = z.object({
  // Position and size
  x: z.number().optional(),
  y: z.number().optional(),
  w: z.number().optional(),
  h: z.number().optional(),
  // Text formatting
  color: z.string().optional(),
  fontFace: z.string().optional(),
  fontSize: fontSizeSchema.optional(),
  bold: z.boolean().optional(),
  italic: z.boolean().optional(),
  underline: textUnderlineSchema.optional(),
  align: textAlignSchema.optional(),
  valign: textValignSchema.optional(),
  // Advanced settings
  bullet: textBulletSchema.optional(),
  fill: textFillSchema.optional(),
  hyperlink: textHyperlinkSchema.optional(),
});
const textContentSchema = z.object({
  text: z.string(),
  options: textOptionsSchema.optional(),
});
const textContentsSchema = z.array(textContentSchema);
const slideSchema = z.object({
  background: backgroundSchema.optional(),
  color: z.string().optional(),
  slideNumber: slideNumberSchema.optional(),
  texts: textContentsSchema.optional(),
});
const stateSchema = z.object({
  metadata: z.object({
    name: z.string(),
    title: z.string().optional(),
    subject: z.string().optional(),
    createdAt: z.string().datetime(),
    updatedAt: z.string().datetime(),
  }),
  slides: z.array(slideSchema),
});

/**
 * State
 */

type Slide = z.infer<typeof slideSchema>;
type State = z.infer<typeof stateSchema>;

function createState(params: {
  name: string;
  title?: string;
  subject?: string;
}): State {
  const now = new Date().toISOString();
  return stateSchema.parse({
    metadata: {
      name: params.name,
      title: params.title,
      subject: params.subject,
      createdAt: now,
      updatedAt: now,
    },
    slides: [],
  });
}

function setUpdatedAt(state: State): State {
  state.metadata.updatedAt = new Date().toISOString();
  return state;
}

function addSlide(state: State, slide: Slide): State {
  state.slides.push(slide);
  return setUpdatedAt(state);
}

function removeSlide(state: State, slideIndex: number): State {
  if (slideIndex < 0 || slideIndex >= state.slides.length) {
    throw new Error(`Invalid slide index: ${slideIndex}`);
  }
  state.slides.splice(slideIndex, 1);
  return setUpdatedAt(state);
}

function replaceSlide(state: State, slideIndex: number, slide: Slide): State {
  if (slideIndex < 0 || slideIndex >= state.slides.length) {
    throw new Error(`Invalid slide index: ${slideIndex}`);
  }
  state.slides[slideIndex] = slide;
  return setUpdatedAt(state);
}

/**
 * State Hydration
 */

const STATE_DIR = path.join(STORAGE_DIR, ".state");

async function ensureStateDir(): Promise<void> {
  try {
    await fs.access(STATE_DIR);
  } catch {
    await fs.mkdir(STATE_DIR, { recursive: true });
  }
}

function getStatePath(name: string): string {
  return path.join(STATE_DIR, `${name}.json`);
}

async function readState(name: string): Promise<State> {
  const statePath = getStatePath(name);
  const stateJson = await fs.readFile(statePath, "utf-8");
  return stateSchema.parse(JSON.parse(stateJson));
}

async function writeState(state: State): Promise<string> {
  const statePath = getStatePath(state.metadata.name);
  await fs.writeFile(statePath, JSON.stringify(state, null, 2));
  return statePath;
}

/**
 * PowerPoint
 */

function getPptxPath(outDir: string, name: string): string {
  return path.join(outDir, `${name}.pptx`);
}

function getPdfPath(outDir: string, name: string): string {
  return path.join(outDir, `${name}.pdf`);
}

function getPngPathGlob(outDir: string, name: string): Glob {
  return new Glob(path.join(outDir, `${name}_slide_*.png`));
}

function getPngPathTmpl(outDir: string, name: string): string {
  return path.join(outDir, `${name}_slide_%d.png`);
}

function getPngPath(outDir: string, name: string, slideIndex: number): string {
  return path.join(outDir, `${name}_slide_${slideIndex}.png`);
}

function createPptxFromState(state: State): PptxGenJS {
  const pptx = new PptxGenJS();

  if (state.metadata.title) {
    pptx.title = state.metadata.title;
  }
  if (state.metadata.subject) {
    pptx.subject = state.metadata.subject;
  }

  for (const slide of state.slides) {
    const pptxSlide = pptx.addSlide();

    if (slide.background) {
      pptxSlide.background = { ...slide.background };
    }
    if (slide.color) {
      pptxSlide.color = slide.color;
    }
    if (slide.slideNumber) {
      pptxSlide.slideNumber = { ...slide.slideNumber };
    }
    if (slide.texts) {
      for (const text of slide.texts) {
        if (Array.isArray(text)) {
          pptxSlide.addText(
            text.map((t) => ({
              text: t.text,
              options: t.options || {},
            }))
          );
        } else {
          pptxSlide.addText(text.text, text.options || {});
        }
      }
    }
  }

  return pptx;
}

async function writePptxFromState(
  state: State,
  outDir: string
): Promise<string> {
  const outPath = getPptxPath(outDir, state.metadata.name);
  const pptx = createPptxFromState(state);
  await pptx.writeFile({ fileName: outPath });
  return outPath;
}

async function writeSlidePngFromState(
  state: State,
  slideIndex: number,
  outDir: string
): Promise<string> {
  const pptxPath = await writePptxFromState(state, outDir);
  const pdfPath = getPdfPath(outDir, state.metadata.name);
  const pngPath = getPngPath(outDir, state.metadata.name, slideIndex);
  await $`soffice --headless --convert-to pdf --outdir ${outDir} ${pptxPath}`;
  await $`magick ${pdfPath}[${slideIndex}] ${pngPath}`;
  await fs.unlink(pdfPath);
  return pngPath;
}

async function writeAllSlidePngFromState(
  state: State,
  outDir: string
): Promise<string[]> {
  const pptxPath = await writePptxFromState(state, outDir);
  const pdfPath = getPdfPath(outDir, state.metadata.name);
  const pngPathTmpl = getPngPathTmpl(outDir, state.metadata.name);
  const pngPathGlob = getPngPathGlob(outDir, state.metadata.name);
  await $`soffice --headless --convert-to pdf --outdir ${outDir} ${pptxPath}`;
  await $`magick ${pdfPath} ${pngPathTmpl}`;
  await fs.unlink(pdfPath);
  return Array.fromAsync(pngPathGlob.scan());
}

/**
 * Utilities
 */

function description(lines: string[]): string {
  return lines.join("\n");
}

function okText(text: string): CallToolResult {
  return { content: [{ type: "text", text }] };
}

function errText(text: string): CallToolResult {
  return { content: [{ type: "text", text }], isError: true };
}

/**
 * Guidelines
 */

const SLIDE_CREATION_GUIDELINE = `
Follow guidelines below when creating presentation slides:

1. Basic Slide Dimensions
- Slide layout: 10 x 5.625 inches
- Coordinate system: Origin (0,0) at top-left, specified in inches (number) or percentages (string)

2. Margins and Safe Areas
- Vertical margins: >= 0.5 inches
- Horizontal margins: >= 0.5 inches
- Usable content area: 9.0 x 4.625 inches
- Bottom safe area: avoid placing content within 0.75 inches (≈13%) from bottom

3. Basic Layout (Inch Specification)
- Title position: 0.5 <= y <= 1.5
- Subtitle position: 1.75 <= y <= 2.25
- Main content start: y = 2.5
- Content lower limit: y = 4.875
- Standard textbox width: w = 8.5

5. Layout Principles
- Slide numbers: Always enable slide numbers
- Item spacing: minimum 0.5 inches (≈9%)
- Font sizes scaled to inch measurements (title: 40-48pt, body: 24-28pt)
`.trim();

/**
 * Main
 */

const server = new McpServer({
  name: "PowerPoint MCP Server",
  version: "0.1.0",
});

server.tool(
  "guide_slide_creation",
  description([
    "Provides guidelines for creating presentation slides.",
    "You must reference this tool when adding or replacing slides.",
  ]),
  async () => okText(SLIDE_CREATION_GUIDELINE)
);

server.tool(
  "create_presentation",
  description([
    "Creates a new presentation file.",
    "You can specify a `title` and `subject`.",
    "`name` must consist of alphanumeric characters, underscores, or hyphens.",
  ]),
  {
    name: z.string(),
    title: z.string().optional(),
    subject: z.string().optional(),
  },
  async ({ name, title, subject }) => {
    try {
      const state = createState({ name, title, subject });
      const statePath = await writeState(state);
      return okText(`Presentation created: ${statePath}`);
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      return errText(`Failed to create presentation: ${message}`);
    }
  }
);

server.tool(
  "add_slide",
  description([
    "Adds a new slide to the presentation.",
    "You must reference the `slide_guidelines` tool when adding a slide.",
  ]),
  {
    name: z.string(),
    background: backgroundSchema.optional(),
    color: z.string().optional(),
    slideNumber: slideNumberSchema.optional(),
    texts: textContentsSchema.optional(),
  },
  async ({ name, background, color, slideNumber, texts }) => {
    try {
      const state = await readState(name);

      const newSlide = slideSchema.safeParse({
        background,
        color,
        slideNumber,
        texts,
      });
      if (!newSlide.success) {
        return errText(`Failed to add slide: ${newSlide.error.message}`);
      }

      const modifiedState = addSlide(state, newSlide.data);
      await writeState(modifiedState);

      return okText(
        `Slide ${state.slides.length} added to ${state.metadata.name}`
      );
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      return errText(`Failed to add slide: ${message}`);
    }
  }
);

server.tool(
  "get_slide",
  "Gets the state of the specified slide.",
  {
    name: z.string(),
    slideIndex: z.number().int().min(0),
  },
  async ({ name, slideIndex }) => {
    try {
      const state = await readState(name);
      if (slideIndex < 0 || slideIndex >= state.slides.length) {
        return errText(`Invalid slide index: ${slideIndex}`);
      }

      const slide = state.slides[slideIndex];
      return okText(JSON.stringify(slide, null, 2));
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      return errText(`Failed to get slide: ${message}`);
    }
  }
);

server.tool(
  "get_slides",
  "Gets the state of all slides.",
  {
    name: z.string(),
  },
  async ({ name }) => {
    try {
      const state = await readState(name);
      return okText(JSON.stringify(state.slides, null, 2));
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      return errText(`Failed to get slide list: ${message}`);
    }
  }
);

server.tool(
  "remove_slide",
  "Removes the slide at the specified index.",
  {
    name: z.string(),
    slideIndex: z.number().int().min(0),
  },
  async ({ name, slideIndex }) => {
    try {
      const state = await readState(name);
      if (slideIndex < 0 || slideIndex >= state.slides.length) {
        return errText(`Invalid slide index: ${slideIndex}`);
      }

      const modifiedState = removeSlide(state, slideIndex);
      await writeState(modifiedState);

      return okText(`Slide ${slideIndex} removed from ${state.metadata.name}`);
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      return errText(`Failed to remove slide: ${message}`);
    }
  }
);

server.tool(
  "replace_slide",
  description([
    "Replaces the slide at the specified index with new content.",
    "You must reference the `slide_guidelines` tool when replacing a slide.",
  ]),
  {
    name: z.string(),
    slideIndex: z.number().int().min(0),
    background: backgroundSchema.optional(),
    color: z.string().optional(),
    slideNumber: slideNumberSchema.optional(),
    texts: textContentsSchema.optional(),
  },
  async ({ name, slideIndex, background, color, slideNumber, texts }) => {
    try {
      const state = await readState(name);

      const newSlide = slideSchema.safeParse({
        background,
        color,
        slideNumber,
        texts,
      });
      if (!newSlide.success) {
        return errText(`Failed to replace slide: ${newSlide.error.message}`);
      }

      const modifiedState = replaceSlide(state, slideIndex, newSlide.data);
      await writeState(modifiedState);

      return okText(`Slide ${slideIndex} replaced in ${state.metadata.name}`);
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      return errText(`Failed to replace slide: ${message}`);
    }
  }
);

server.tool(
  "export_presentation_as_pptx",
  description([
    "Exports the presentation as a PPTX file.",
    "`outDir` must be an absolute path.",
  ]),
  {
    name: z.string(),
    outDir: z.string(),
  },
  async ({ name, outDir }) => {
    try {
      const state = await readState(name);
      const outPath = await writePptxFromState(state, outDir);
      return okText(`Saved PPTX file: ${outPath}`);
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      return errText(`Failed to save PPTX file: ${message}`);
    }
  }
);

server.tool(
  "export_slide_as_png",
  description([
    "Exports the specified slide as a PNG image.",
    "`outDir` must be an absolute path.",
  ]),
  {
    name: z.string(),
    slideIndex: z.number().int().min(0),
    outDir: z.string(),
  },
  async ({ name, slideIndex, outDir }) => {
    try {
      const state = await readState(name);
      if (slideIndex < 0 || slideIndex >= state.slides.length) {
        return errText(`Invalid slide index: ${slideIndex}`);
      }

      const pngPath = await writeSlidePngFromState(state, slideIndex, outDir);
      return okText(`Saved slide: ${pngPath}`);
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      return errText(`Failed to save slide: ${message}`);
    }
  }
);

server.tool(
  "export_slides_as_png",
  description([
    "Exports all slides as PNG images.",
    "`outDir` must be an absolute path.",
  ]),
  {
    name: z.string(),
    outDir: z.string(),
  },
  async ({ name, outDir }) => {
    try {
      const state = await readState(name);
      const pngPaths = await writeAllSlidePngFromState(state, outDir);
      return okText(`Saved slides: ${pngPaths.join(", ")}`);
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      return errText(`Failed to save slides: ${message}`);
    }
  }
);

await ensureStateDir();

const transport = new StdioServerTransport();
await server.connect(transport);
