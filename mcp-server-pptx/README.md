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

#### `create_presentation`

Create a new presentation file

```ts
function create_presentation(params: {
  // Name of the presentation file
  name: string;
  // Title text
  title?: string;
  // Subject text
  subject?: string;
}) => void;
```

#### `save_as_pptx`

Save the presentation as a PPTX file

```ts
function save_as_pptx(params: {
  // Name of the presentation file
  name: string;
}) => void;
```

#### `add_slide`

Add a new slide to the presentation

```ts
function add_slide(params: {
  // Name of the presentation file
  name: string;
  // Background properties for the slide
  background?: Background;
  // Default text color for the slide in hex format
  color?: string;
  // Slide number properties
  slideNumber?: SlideNumber;
}) => void;
```

#### Shared Types

```ts
type Background = {
  // Background color in hex format (e.g., "F1F1F1")
  color?: string;
  // Background transparency (0-100)
  transparency?: number;
};

type SlideNumber = {
  // Horizontal position in inches (number) or percentage (string)
  // @example 1.0 or "50%"
  x: number | string;
  // Vertical position in inches (number) or percentage (string)
  // @example 1.0 or "90%"
  y: number | string;
  // Color in hex format (default: "000000")
  color?: string;
  // Font face (e.g., "Arial")
  fontFace?: string;
  // Font size (8-256)
  fontSize?: number;
};
```

<!--

## Future

### Tools

#### `presentation_delete`

#### `presentation_get_as_png`

#### `slide_add`

#### `slide_delete`

#### `slide_get_as_png`

#### `slide_get_as_pptx`

#### `slide_update_master`

#### `slide_update`

-->
