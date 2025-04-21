# mcp-server-pptx

MCP server for generating and managing PowerPoint presentations using PptxGenJS.

## Prerequisites

- [LibreOffice](https://ja.libreoffice.org/)
  - `brew install --cask libreoffice`
- [ImageMagick](https://imagemagick.org/)
  - `brew install imagemagick`
- [Ghostscript](https://www.ghostscript.com/)
  - `brew install ghostscript`

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

Create a new presentation file.

```ts
function create_presentation(params: {
  name: string; // Name of the presentation file
  title?: string; // Title text
  subject?: string; // Subject text
}) => void;
```

#### `save_as_pptx`

Save the presentation as a PPTX file.

```ts
function save_as_pptx(params: {
  name: string; // Name of the presentation file
}) => void;
```

#### `add_slide`

Add a new slide to the presentation.

```ts
function add_slide(params: {
  name: string; // Name of the presentation file
  background?: Background;  // Background properties for the slide
  color?: string; // Default text color for the slide in hex format
  slideNumber?: SlideNumber; // Slide number properties
  texts?: SlideText[]; // Array of text elements to add to the slide
}) => void;
```

#### `replace_slide`

Replace a slide in the presentation.

```ts
function replace_slide(params: {
  name: string; // Name of the presentation file
  background?: Background; // Background properties for the slide
  color?: string; // Default text color for the slide in hex format
  slideNumber?: SlideNumber; // Slide number properties
  texts?: SlideText[]; // Array of text elements to add to the slide
}) => void;
```

#### `get_slide_as_png`

Get a slide as a PNG image.

```ts
function get_slide_as_png(params: {
  name: string; // Name of the presentation file
  slideIndex: number; // Index of the slide to get
}) => string;
```

#### Shared Types

```ts
type Background = {
  color?: string; // Background color in hex format (e.g., "F1F1F1")
  transparency?: number; // Background transparency (0-100)
};

type SlideNumber = {
  x: number | string; // Horizontal position in inches (number) or percentage (string) (e.g., 1.0 or "50%")
  y: number | string; // Vertical position in inches (number) or percentage (string) (e.g., 1.0 or "90%")
  color?: string; // Color in hex format (default: "000000")
  fontFace?: string; // Font face (e.g., "Arial")
  fontSize?: number; // Font size (8-256)
};

type TextContent = {
  text: string; // Text content
  options?: TextOptions; // Text formatting options
};

type TextOptions = {
  // Position and size
  x?: number | string; // inches or percentage (e.g., 1.0 or "50%")
  y?: number | string; // inches or percentage (e.g., 1.0 or "50%")
  w?: number | string; // inches or percentage (e.g., 2.0 or "30%")
  h?: number | string; // inches or percentage (e.g., 1.0 or "10%")

  // Text formatting
  color?: string; // hex color code (e.g., "0088CC")
  fontFace?: string; // font name (e.g., "Arial")
  fontSize?: number; // font size (8-256)
  bold?: boolean; // bold text
  italic?: boolean; // italic text
  underline?: boolean; // underlined text
  align?: "left" | "center" | "right"; // horizontal alignment
  valign?: "top" | "middle" | "bottom"; // vertical alignment

  // Advanced settings
  bullet?:
    | boolean
    | {
        // bullet point settings
        type?: "number" | string;
        code?: string;
        style?: string;
      };
  fill?: {
    // background color settings
    color: string;
  };
  hyperlink?: {
    // hyperlink settings
    url?: string;
    slide?: string;
    tooltip?: string;
  };
};

type SlideText = TextContent | TextContent[];
```
