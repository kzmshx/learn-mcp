import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import cuid2 from "@paralleldrive/cuid2";
import { z } from "zod";

function createIdFactory(length: number, fingerprint?: string) {
  return cuid2.init({
    length,
    fingerprint,
  });
}

const server = new McpServer({
  name: "Cuid2 MCP Server",
  version: "0.1.0",
});

server.tool(
  "generate_cuid2",
  {
    length: z.number().min(1).max(128),
    count: z.number().min(1).max(100),
    fingerprint: z.string().optional(),
  },
  async ({ length, count, fingerprint }) => {
    const createId = createIdFactory(length, fingerprint);

    return {
      content: [
        {
          type: "text",
          text: Array.from({ length: count }, () => createId()).join("\n"),
        },
      ],
    };
  }
);

const transport = new StdioServerTransport();
await server.connect(transport);
