import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import * as path from "path";
import { z } from "zod";

const envSchema = z.object({
  STORAGE_DIR: z.string().min(1, "STORAGE_DIR is required"),
});
const env = envSchema.parse(process.env);
const storageDir = path.resolve(env.STORAGE_DIR);

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
    const presentationPath = path.join(storageDir, filename);
    // TODO: 実装を追加
    return {
      content: [
        {
          type: "text",
          text: `Created presentation: ${presentationPath}`,
        },
      ],
    };
  }
);

server.tool(
  "presentation_flush_pptx",
  {
    filename: z.string(),
  },
  async ({ filename }) => {
    const presentationPath = path.join(storageDir, filename);
    // TODO: 実装を追加
    return {
      content: [
        {
          type: "text",
          text: `Generated PPTX file: ${presentationPath}`,
        },
      ],
    };
  }
);

const transport = new StdioServerTransport();
await server.connect(transport);
