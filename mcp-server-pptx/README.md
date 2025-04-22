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

#### `guide_slide_creation`

Get guidelines for creating a presentation slide.

#### `create_presentation`

Create a new presentation file.

- Parameters
  - `name` (string): Name of the presentation
  - `title` (string | undefined): Title text
  - `subject` (string | undefined): Subject text
- Returns
  - Created presentation file path

#### `add_slide`

Add a new slide to the presentation.

- Parameters
  - `name` (string): Name of the presentation
  - `background` (Background | undefined): Background properties for the slide
  - `color` (string | undefined): Default text color for the slide in hex format
  - `slideNumber` (SlideNumber | undefined): Slide number properties
  - `texts` (SlideText[] | undefined): Array of text elements to add to the slide
- Returns
  - Added slide index

#### `get_slide`

Gets the state of the specified slide.

- Parameters
  - `name` (string): Name of the presentation
  - `slideIndex` (number): Index of the slide to get
- Returns
  - State of the specified slide

#### `get_slides`

Gets the state of all slides.

- Parameters
  - `name` (string): Name of the presentation
- Returns
  - State of all slides

#### `remove_slide`

Removes the slide at the specified index.

- Parameters
  - `name` (string): Name of the presentation
  - `slideIndex` (number): Index of the slide to remove
- Returns
  - Removed slide index

#### `replace_slide`

Replace a slide in the presentation.

- Parameters
  - `name` (string): Name of the presentation
  - `slideIndex` (number): Index of the slide to replace
  - `background` (Background | undefined): Background properties for the slide
  - `color` (string | undefined): Default text color for the slide in hex format
  - `slideNumber` (SlideNumber | undefined): Slide number properties
  - `texts` (SlideText[] | undefined): Array of text elements to add to the slide
- Returns
  - Replaced slide index

#### `export_presentation_as_pptx`

Exports the presentation as a PPTX file.

- Parameters
  - `name`: Name of the presentation
  - `outDir`: Directory to save the PPTX file
- Returns
  - Exported PPTX file path

#### `export_slide_as_png`

Exports the specified slide as a PNG image.

- Parameters
  - `name`: Name of the presentation
  - `slideIndex`: Index of the slide to export
  - `outDir`: Directory to save the PNG file
- Returns
  - Exported PNG file path

#### `export_slides_as_png`

Exports all slides as PNG images.

- Parameters
  - `name`: Name of the presentation
  - `outDir`: Directory to save the PNG files
- Returns
  - Exported PNG file paths

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
