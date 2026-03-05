# DOCX MCP Server

A comprehensive Model Context Protocol (MCP) server for processing Microsoft Word (.docx) documents with full formatting support.

## Features

This MCP server provides advanced DOCX document processing capabilities using the powerful `mammoth` library:

- **Text Extraction**: Extract plain text with word count
- **HTML Conversion**: Convert to HTML with preserved formatting
- **Structure Analysis**: Analyze document structure, headings, and formatting elements
- **Image Extraction**: Extract embedded images (as base64 or save to files)
- **Markdown Conversion**: Convert to Markdown format
- **Rich Formatting Support**: Handles bold, italic, lists, headings, and more

## Available Tools

### 1. `extract_text`

Extract plain text content from a DOCX file.

**Parameters:**

- `file_path` (string): Path to the .docx file

**Returns:**

- Plain text content
- Processing messages
- Word count

### 2. `convert_to_html`

Convert DOCX file to HTML with formatting preserved.

**Parameters:**

- `file_path` (string): Path to the .docx file
- `include_styles` (boolean, optional): Include inline styles (default: true)

**Returns:**

- HTML content with formatting
- Processing messages
- Warnings and errors

### 3. `analyze_structure`

Analyze document structure, headings, and formatting elements.

**Parameters:**

- `file_path` (string): Path to the .docx file

**Returns:**

- Document statistics (characters, words, paragraphs, headings)
- Structure analysis (headings with levels)
- Formatting analysis (bold, italic, lists count)
- Processing messages

### 4. `extract_images`

Extract and list images from a DOCX file.

**Parameters:**

- `file_path` (string): Path to the .docx file
- `output_dir` (string, optional): Directory to save extracted images

**Returns:**

- Total image count
- Image details (src, alt text, base64 status)
- Output directory information
- Processing messages

### 5. `convert_to_markdown`

Convert DOCX file to Markdown format.

**Parameters:**

- `file_path` (string): Path to the .docx file

**Returns:**

- Markdown content
- Word count
- Processing messages

## HTTP Server Mode

In addition to MCP stdio mode, this server can run as a standalone HTTP server for direct API access.

### Running in HTTP Mode

Set the `USE_HTTP` environment variable to `true`:

```bash
USE_HTTP=true npm start
# or
USE_HTTP=true node build/index.js
```

The server will start on port 3000 (configurable via `PORT` environment variable).

### HTTP API Endpoints

All MCP tools are available as HTTP POST endpoints with multipart file upload:

- `POST /extract_text` - Extract text
- `POST /convert_to_html` - Convert to HTML
- `POST /analyze_structure` - Analyze structure
- `POST /extract_images` - Extract images
- `POST /convert_to_markdown` - Convert to Markdown

**Request Format:** Multipart form-data with:
- `file`: The .docx file to process
- Additional parameters as form fields (for endpoints that need them)

**Response:** JSON with the tool's return data.

Example usage with curl:

```bash
# Extract text
curl -X POST -F "file=@document.docx" http://localhost:3000/extract_text

# Convert to HTML with custom styling
curl -X POST -F "file=@document.docx" -F "include_styles=false" http://localhost:3000/convert_to_html

# Extract images to a directory
curl -X POST -F "file=@document.docx" -F "output_dir=./images" http://localhost:3000/extract_images
```

## Installation

```bash
npm install
npm run build
```

## Usage

The server runs on stdio and communicates via JSON-RPC 2.0 protocol.

### Example Usage with MCP Client

```json
{
  "jsonrpc": "2.0",
  "id": 1,
  "method": "tools/call",
  "params": {
    "name": "analyze_structure",
    "arguments": {
      "file_path": "/path/to/document.docx"
    }
  }
}
```

### Example Usage with Roo

```json
{
  "file_path": "/path/to/document.docx"
}
```

## Supported Features

- ✅ **Text Extraction**: Plain text with word counting
- ✅ **Rich Formatting**: Bold, italic, underline, strikethrough
- ✅ **Document Structure**: Headings (H1-H6), paragraphs
- ✅ **Lists**: Ordered and unordered lists with items
- ✅ **Images**: Extraction as base64 or file export
- ✅ **Tables**: Basic table structure (via HTML conversion)
- ✅ **Links**: Hyperlinks preservation
- ✅ **Styles**: Custom style mapping support
- ✅ **Error Handling**: Comprehensive error reporting
- ✅ **Multiple Formats**: HTML, Markdown, plain text output

## Advanced Features

### Custom Style Mapping

The `convert_to_html` tool supports custom style mapping for better semantic HTML output:

```javascript
// Example style mappings
"p[style-name='Heading 1'] => h1:fresh"
"r[style-name='Strong'] => strong"
"r[style-name='Emphasis'] => em"
```

### Image Handling

- **Base64 Embedding**: Images can be embedded as base64 data URLs
- **File Export**: Images can be extracted to a specified directory
- **Metadata**: Alt text and content type preservation

### Document Analysis

Provides comprehensive document analysis including:

- Character and word counts
- Paragraph and heading counts
- Formatting element statistics
- Document structure hierarchy

## Development

Install dependencies:

```bash
npm install
```

Build the server:

```bash
npm run build
```

For development with auto-rebuild:

```bash
npm run watch
```

## Installation for Claude Desktop

To use with Claude Desktop, add the server config:

On MacOS: `~/Library/Application Support/Claude/claude_desktop_config.json`
On Windows: `%APPDATA%/Claude/claude_desktop_config.json`

```json
{
  "mcpServers": {
    "docx-format-server": {
      "command": "/path/to/docx-format-server/build/index.js"
    }
  }
}
```

## Dependencies

- `@modelcontextprotocol/sdk`: MCP protocol implementation
- `mammoth`: Advanced DOCX processing library
- `zod`: Schema validation
- `typescript`: TypeScript support

## Error Handling

All tools include comprehensive error handling with detailed error messages for:

- File not found errors
- Invalid file format
- Processing errors
- Permission issues

## Debugging

Since MCP servers communicate over stdio, debugging can be challenging. We recommend using the [MCP Inspector](https://github.com/modelcontextprotocol/inspector), which is available as a package script:

```bash
npm run inspector
```

The Inspector will provide a URL to access debugging tools in your browser.

## Version History

- **v0.3.0**: Added HTTP server mode with REST API endpoints alongside MCP stdio mode
- **v0.2.0**: Complete rewrite with mammoth library, added 5 comprehensive tools
- **v0.1.0**: Basic text extraction with docx-parser (deprecated)

## License

ISC License
