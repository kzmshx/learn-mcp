import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import * as fs from "fs/promises";
import * as path from "path";
import PptxGenJS from "pptxgenjs";
import { z } from "zod";

const envSchema = z.object({
  STORAGE_DIR: z.string().min(1, "STORAGE_DIR is required"),
});
const env = envSchema.parse(process.env);
const STORAGE_DIR = path.resolve(env.STORAGE_DIR);
const STATE_DIR = path.join(STORAGE_DIR, ".state");

const stateSchema = z.object({
  metadata: z.object({
    filename: z.string(),
    title: z.string().optional(),
    subject: z.string().optional(),
    createdAt: z.string().datetime(),
    updatedAt: z.string().datetime(),
  }),
  slides: z.array(
    z.object({
      index: z.number().int().nonnegative(),
    })
  ),
});

type State = z.infer<typeof stateSchema>;

function initState(params: {
  filename: string;
  title?: string;
  subject?: string;
}): State {
  const now = new Date().toISOString();
  return stateSchema.parse({
    metadata: {
      filename: params.filename,
      title: params.title,
      subject: params.subject,
      createdAt: now,
      updatedAt: now,
    },
    slides: [],
  });
}

function getStateFilePath(filename: string): string {
  return path.join(STATE_DIR, `${filename}.json`);
}

async function readState(filename: string): Promise<State> {
  const stateFilePath = getStateFilePath(filename);
  const stateJson = await fs.readFile(stateFilePath, "utf-8");
  return stateSchema.parse(JSON.parse(stateJson));
}

async function writeState(state: State): Promise<void> {
  const stateFilePath = getStateFilePath(state.metadata.filename);
  await fs.writeFile(stateFilePath, JSON.stringify(state, null, 2));
}

async function ensureStateDir(): Promise<void> {
  try {
    await fs.access(STATE_DIR);
  } catch {
    await fs.mkdir(STATE_DIR, { recursive: true });
  }
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
    filename: z.string(),
    title: z.string().optional(),
    subject: z.string().optional(),
  },
  async ({ filename, title, subject }) => {
    try {
      const state = initState({ filename, title, subject });
      await writeState(state);
      return {
        content: [
          {
            type: "text",
            text: `Created presentation state: ${getStateFilePath(filename)}`,
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
    filename: z.string(),
  },
  async ({ filename }) => {
    try {
      const state = await readState(filename);
      const pptx = new PptxGenJS();

      // Set presentation properties
      if (state.metadata.title) {
        pptx.title = state.metadata.title;
      }
      if (state.metadata.subject) {
        pptx.subject = state.metadata.subject;
      }

      // Add slides (currently empty as per schema)
      for (const slide of state.slides) {
        pptx.addSlide();
      }

      // Save PPTX file
      const pptxPath = path.join(STORAGE_DIR, filename);
      await pptx.writeFile({ fileName: pptxPath });

      return {
        content: [
          {
            type: "text",
            text: `Generated PPTX file: ${pptxPath}`,
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

await ensureStateDir();

const transport = new StdioServerTransport();
await server.connect(transport);
