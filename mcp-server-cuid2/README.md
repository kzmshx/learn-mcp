# mcp-server-cuid2

MCP server for generating CUID2 identifiers.

## Features

- **Generate CUID2**: Generate one or more CUID2 identifiers with customizable length and optional fingerprint

## Installation

To add this MCP server to your environment, add the following to your MCP config file:

```json
{
  "mcpServers": {
    "kzmshx.mcp-server-cuid2": {
      "autoApprove": [],
      "args": ["/path/to/start.sh"],
      "command": "sh"
    }
  }
}
```

## Tools

### generate_cuid2

```ts
/**
 * Generate one or more CUID2 identifiers
 *
 * @param {number} length - Length of each CUID2 identifier (min: 1, max: 128)
 * @param {number} count - Number of identifiers to generate (min: 1, max: 100)
 * @param {string} [fingerprint] - Custom fingerprint for ID generation
 * @returns {string} A newline-separated string containing the generated CUID2 identifiers
 */
type GenerateCuid2 = (
  length: number;
  count: number;
  fingerprint?: string;
) => string;
```
