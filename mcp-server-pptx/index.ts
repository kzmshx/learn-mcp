import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
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

const slideNumberSchema = z.object({
  x: z.number(),
  y: z.number(),
  color: z.string().optional(),
  fontFace: z.string().optional(),
  fontSize: z.number().min(8).max(256).optional(),
});

const slideSchema = z.object({
  background: backgroundSchema.optional(),
  color: z.string().optional(),
  slideNumber: slideNumberSchema.optional(),
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

type Background = z.infer<typeof backgroundSchema>;
type SlideNumber = z.infer<typeof slideNumberSchema>;
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

function getStateFilePath(name: string): string {
  return path.join(STATE_DIR, `${name}.json`);
}

async function readState(name: string): Promise<State> {
  const stateFilePath = getStateFilePath(name);
  const stateJson = await fs.readFile(stateFilePath, "utf-8");
  return stateSchema.parse(JSON.parse(stateJson));
}

async function writeState(state: State): Promise<void> {
  const stateFilePath = getStateFilePath(state.metadata.name);
  await fs.writeFile(stateFilePath, JSON.stringify(state, null, 2));
}

/**
 * PowerPoint
 */

function getPptxFilePath(name: string): string {
  return path.join(STORAGE_DIR, `${name}.pptx`);
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
  }

  return pptx;
}

/**
 * Main
 */

const server = new McpServer({
  name: "PowerPoint MCP Server",
  version: "0.1.0",
});

server.tool(
  "presentation_create",
  {
    name: z.string(),
    title: z.string().optional(),
    subject: z.string().optional(),
  },
  async ({ name, title, subject }) => {
    try {
      const state = createState({ name, title, subject });
      await writeState(state);
      return {
        content: [
          {
            type: "text",
            text: `Created presentation state: ${getStateFilePath(name)}`,
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Failed to create presentation: ${
              error instanceof Error ? error.message : String(error)
            }`,
          },
        ],
        isError: true,
      };
    }
  }
);

server.tool(
  "presentation_flush_pptx",
  {
    name: z.string(),
  },
  async ({ name }) => {
    try {
      const state = await readState(name);
      const pptx = createPptxFromState(state);
      const pptxFilePath = getPptxFilePath(name);
      await pptx.writeFile({ fileName: pptxFilePath });

      return {
        content: [
          {
            type: "text",
            text: `Generated PPTX file: ${pptxFilePath}`,
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Failed to generate PPTX file: ${
              error instanceof Error ? error.message : String(error)
            }`,
          },
        ],
        isError: true,
      };
    }
  }
);

server.tool(
  "slide_add",
  {
    name: z.string(),
    background: backgroundSchema.optional(),
    color: z.string().optional(),
    slideNumber: slideNumberSchema.optional(),
  },
  async ({ name, background, color, slideNumber }) => {
    try {
      const state = await readState(name);

      const newSlide = slideSchema.parse({
        background,
        color,
        slideNumber,
      });

      const slideAddedState = addSlide(state, newSlide);
      await writeState(slideAddedState);

      return {
        content: [
          {
            type: "text",
            text: `Added slide ${state.slides.length} to presentation: ${state.metadata.name}`,
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Failed to add slide: ${
              error instanceof Error ? error.message : String(error)
            }`,
          },
        ],
        isError: true,
      };
    }
  }
);

await ensureStateDir();

const transport = new StdioServerTransport();
await server.connect(transport);
