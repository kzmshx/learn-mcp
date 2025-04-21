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
      "config": {
        "storageDir": "/path/to/pptx/storage"
      }
    }
  }
}
```

## v0.1.0

### Tools

#### `presentation_add`

```ts
/**
 * Create a new presentation file
 */
function presentation_create(params: {
  /**
   * Name of the presentation file
   */
  filename: string;
  /**
   * Title text
   */
  title?: string;
  /**
   * Subject text
   */
  subject?: string;
}) => void;
```

#### `presentation_get_as_pptx`

```ts
/**
 * Save the presentation fle as a PPTX file
 */
function presentation_flush_pptx(params: {
  filename: string;
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
