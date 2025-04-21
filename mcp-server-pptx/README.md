# mcp-server-pptx

MCP server for generating and managing PowerPoint presentations using PptxGenJS.

## Installation

To add this MCP server to your environment, add the following to your MCP config file:

```json
{
  "mcpServers": {
    "kzmshx.mcp-server-pptx": {
      "autoApprove": [],
      "args": ["/path/to/start.sh"],
      "command": "sh",
      "env": {
        "STORAGE_DIR": "/path/to/pptx/storage"
      }
    }
  }
}
```

## v0.1.0

### Tools

#### `presentation_add`

Create a new presentation file

```ts
function presentation_create(params: {
  // Name of the presentation file
  name: string;
  // Title text
  title?: string;
  // Subject text
  subject?: string;
}) => void;
```

#### `presentation_flush_pptx`

Flush the presentation file to the storage directory

```ts
function presentation_flush_pptx(params: {
  name: string;
}) => void;
```

<!--

## Future

### Resources

#### `pptx://schema`

JSON schema for the presentation state.

#### `state:///<filename>`

Current state of the presentation.

#### `state:///<filename>/<slide_index>`

Current state of the slide at the given index.

### Tools

#### `presentation_add`

#### `presentation_delete`

#### `presentation_get_as_png`

#### `presentation_get_as_pptx`

#### `slide_add`

#### `slide_delete`

#### `slide_get_as_png`

#### `slide_get_as_pptx`

#### `slide_update_master`

#### `slide_update`

-->
