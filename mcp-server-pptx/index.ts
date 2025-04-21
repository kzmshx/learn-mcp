import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { $ } from "bun";
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
 * State
 */

const backgroundSchema = z.object({
  color: z.string().optional(),
  transparency: z.number().min(0).max(100).optional(),
});
type Background = z.infer<typeof backgroundSchema>;

const slideNumberSchema = z.object({
  x: z.number(),
  y: z.number(),
  color: z.string().optional(),
  fontFace: z.string().optional(),
  fontSize: z.number().min(8).max(256).optional(),
});
type SlideNumber = z.infer<typeof slideNumberSchema>;

const textOptionUnderlineStyleSchema = z.enum([
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
type TextOptionUnderlineStyle = z.infer<typeof textOptionUnderlineStyleSchema>;

const textOptionUnderlineSchema = z.object({
  style: textOptionUnderlineStyleSchema.optional(),
  color: z.string().optional(),
});
type TextOptionUnderline = z.infer<typeof textOptionUnderlineSchema>;

const textOptionBulletTypeSchema = z.enum(["number", "bullet"]);
type TextOptionBulletType = z.infer<typeof textOptionBulletTypeSchema>;

const textOptionBulletNumberTypeSchema = z.enum([
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
type TextOptionBulletNumberType = z.infer<
  typeof textOptionBulletNumberTypeSchema
>;

const textOptionBulletDetailSchema = z.object({
  type: textOptionBulletTypeSchema.optional(),
  characterCode: z.string().optional(),
  indent: z.number().optional(),
  numberType: textOptionBulletNumberTypeSchema.optional(),
  style: z.string().optional(),
});
type TextOptionBulletDetail = z.infer<typeof textOptionBulletDetailSchema>;

const textOptionBulletSchema = z.union([
  z.boolean(),
  textOptionBulletDetailSchema,
]);
type TextOptionBullet = z.infer<typeof textOptionBulletSchema>;

const textOptionsSchema = z.object({
  // Position and size
  x: z.number().optional(),
  y: z.number().optional(),
  w: z.number().optional(),
  h: z.number().optional(),

  // Text formatting
  color: z.string().optional(),
  fontFace: z.string().optional(),
  fontSize: z.number().min(8).max(256).optional(),
  bold: z.boolean().optional(),
  italic: z.boolean().optional(),
  underline: textOptionUnderlineSchema.optional(),
  align: z.enum(["left", "center", "right"]).optional(),
  valign: z.enum(["top", "middle", "bottom"]).optional(),

  // Advanced settings
  bullet: textOptionBulletSchema.optional(),
  fill: z
    .object({
      color: z.string(),
    })
    .optional(),
  hyperlink: z
    .object({
      url: z.string().optional(),
      slide: z.number().optional(),
      tooltip: z.string().optional(),
    })
    .optional(),
});
type TextOptions = z.infer<typeof textOptionsSchema>;

const textContentSchema = z.object({
  text: z.string(),
  options: textOptionsSchema.optional(),
});
type TextContent = z.infer<typeof textContentSchema>;

const slideTextSchema = z.array(
  z.union([textContentSchema, z.array(textContentSchema)])
);
type SlideText = z.infer<typeof slideTextSchema>;

const slideSchema = z.object({
  background: backgroundSchema.optional(),
  color: z.string().optional(),
  slideNumber: slideNumberSchema.optional(),
  texts: slideTextSchema.optional(),
});
type Slide = z.infer<typeof slideSchema>;

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

function getPptxPath(name: string): string {
  return path.join(STORAGE_DIR, `${name}.pptx`);
}

function getPdfPath(name: string): string {
  return path.join(STORAGE_DIR, `${name}.pdf`);
}

function getPngPathTemplate(name: string): string {
  return path.join(STORAGE_DIR, `${name}_slide_%d.png`);
}

function getPngPath(name: string, slideIndex: number): string {
  return path.join(STORAGE_DIR, `${name}_slide_${slideIndex}.png`);
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

async function writePptxFromState(state: State): Promise<string> {
  const pptx = createPptxFromState(state);
  const pptxPath = getPptxPath(state.metadata.name);
  await pptx.writeFile({ fileName: pptxPath });
  return pptxPath;
}

async function convertPptxSlideToPng(
  name: string,
  slideIndex: number
): Promise<void> {
  const pptxPath = getPptxPath(name);
  const pdfPath = getPdfPath(name);
  const pdfDir = path.dirname(pdfPath);
  const pngPathTemplate = getPngPathTemplate(name);

  console.error(pptxPath);
  console.error(pdfPath);
  console.error(pdfDir);
  console.error(pngPathTemplate);
  console.error(slideIndex);

  await $`soffice --headless --convert-to pdf --outdir "${pdfDir}" "${pptxPath}"`;
  await $`convert -density 600 -resize 960x540 "${pdfPath}"[${slideIndex}] "${pngPathTemplate}"`;
  await fs.unlink(pdfPath);
}

/**
 * Main
 */

const server = new McpServer({
  name: "PowerPoint MCP Server",
  version: "0.1.0",
});

server.tool(
  "create_presentation",
  {
    name: z.string(),
    title: z.string().optional(),
    subject: z.string().optional(),
  },
  async ({ name, title, subject }) => {
    try {
      const state = createState({ name, title, subject });
      const statePath = await writeState(state);

      return {
        content: [{ type: "text", text: `Created presentation: ${statePath}` }],
      };
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);

      return {
        content: [
          { type: "text", text: `Failed to create presentation: ${message}` },
        ],
        isError: true,
      };
    }
  }
);

server.tool(
  "save_as_pptx",
  {
    name: z.string(),
  },
  async ({ name }) => {
    try {
      const state = await readState(name);
      const pptxFilePath = await writePptxFromState(state);

      return {
        content: [{ type: "text", text: `Saved PPTX file: ${pptxFilePath}` }],
      };
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);

      return {
        content: [
          { type: "text", text: `Failed to save PPTX file: ${message}` },
        ],
        isError: true,
      };
    }
  }
);

server.tool(
  "add_slide",
  {
    name: z.string(),
    background: backgroundSchema.optional(),
    color: z.string().optional(),
    slideNumber: slideNumberSchema.optional(),
    texts: slideTextSchema.optional(),
  },
  async ({ name, background, color, slideNumber, texts }) => {
    try {
      const state = await readState(name);
      const newSlide = slideSchema.parse({
        background,
        color,
        slideNumber,
        texts,
      });
      const modifiedState = addSlide(state, newSlide);
      await writeState(modifiedState);

      return {
        content: [
          {
            type: "text",
            text: `Added slide ${state.slides.length} to presentation: ${state.metadata.name}`,
          },
        ],
      };
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);

      return {
        content: [{ type: "text", text: `Failed to add slide: ${message}` }],
        isError: true,
      };
    }
  }
);

server.tool(
  "replace_slide",
  {
    name: z.string(),
    slideIndex: z.number().int().min(0),
    background: backgroundSchema.optional(),
    color: z.string().optional(),
    slideNumber: slideNumberSchema.optional(),
    texts: slideTextSchema.optional(),
  },
  async ({ name, slideIndex, background, color, slideNumber, texts }) => {
    try {
      const state = await readState(name);
      const newSlide = slideSchema.parse({
        background,
        color,
        slideNumber,
        texts,
      });
      const modifiedState = replaceSlide(state, slideIndex, newSlide);
      await writeState(modifiedState);

      return {
        content: [
          {
            type: "text",
            text: `Replaced slide ${slideIndex} in presentation: ${state.metadata.name}`,
          },
        ],
      };
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);

      return {
        content: [
          { type: "text", text: `Failed to replace slide: ${message}` },
        ],
        isError: true,
      };
    }
  }
);

server.tool(
  "get_slide_as_png",
  {
    name: z.string(),
    slideIndex: z.number().int().min(0),
  },
  async ({ name, slideIndex }) => {
    try {
      const state = await readState(name);
      if (slideIndex < 0 || slideIndex >= state.slides.length) {
        throw new Error(`Invalid slide index: ${slideIndex}`);
      }

      await writePptxFromState(state);
      await convertPptxSlideToPng(name, slideIndex);
      const pngPath = getPngPath(name, slideIndex);

      return {
        content: [{ type: "text", text: `PNG file is saved at: ${pngPath}` }],
      };
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);

      return {
        content: [
          { type: "text", text: `Failed to get slide as PNG: ${message}` },
        ],
        isError: true,
      };
    }
  }
);

await ensureStateDir();

const transport = new StdioServerTransport();
await server.connect(transport);
