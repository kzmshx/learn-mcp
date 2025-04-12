import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import * as chrono from "chrono-node";
import { Temporal } from "temporal-polyfill";
import { z } from "zod";

function formatTimeResponse(instant: Temporal.Instant, timezone: string) {
  const tz = timezone || Temporal.Now.timeZoneId();
  const zdt = instant.toZonedDateTimeISO(tz);
  const response = {
    iso: instant.toString(),
    timezone: tz.toString(),
    timestamp: Number(instant.epochMilliseconds),
    parts: {
      year: zdt.year,
      month: zdt.month,
      day: zdt.day,
      hour: zdt.hour,
      minute: zdt.minute,
      second: zdt.second,
      millisecond: zdt.millisecond,
    },
  };
  return JSON.stringify(response, null, 2);
}

const server = new McpServer({
  name: "Time MCP Server",
  version: "0.1.0",
});

server.tool(
  "get_current_time",
  {
    timezone: z.string().optional(),
  },
  async ({ timezone }) => {
    const now = Temporal.Now.instant();
    return {
      content: [
        { type: "text", text: formatTimeResponse(now, timezone || "") },
      ],
    };
  }
);

server.tool(
  "calculate_time",
  {
    query: z.string(),
    timezone: z.string().optional(),
  },
  async ({ query, timezone }) => {
    const parsed = chrono.parseDate(query, {
      instant: new Date(),
      timezone,
    });
    if (!parsed) {
      throw new Error("Could not parse the time query");
    }

    const instant = Temporal.Instant.fromEpochMilliseconds(parsed.getTime());
    return {
      content: [
        {
          type: "text",
          text: formatTimeResponse(instant, timezone || ""),
        },
      ],
    };
  }
);

const transport = new StdioServerTransport();
await server.connect(transport);
