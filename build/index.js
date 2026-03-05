#!/usr/bin/env node
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { z } from 'zod';
import * as fs from 'fs';
import * as path from 'path';
import { createRequire } from 'module';
import express from 'express';
import process from 'process';
// Create require for CommonJS modules
const require = createRequire(import.meta.url);
const mammoth = require('mammoth');
// Shared functions for DOCX processing
async function extractText(file_path) {
    const absolutePath = path.resolve(file_path);
    if (!fs.existsSync(absolutePath)) {
        throw new Error(`File not found: ${absolutePath}`);
    }
    const result = await mammoth.extractRawText({ path: absolutePath });
    return {
        text: result.value,
        messages: result.messages,
        word_count: result.value.split(/\s+/).filter((word) => word.length > 0).length,
    };
}
async function convertToHtml(file_path, include_styles = true) {
    const absolutePath = path.resolve(file_path);
    if (!fs.existsSync(absolutePath)) {
        throw new Error(`File not found: ${absolutePath}`);
    }
    const options = include_styles ? {} : {
        styleMap: [
            "p[style-name='Heading 1'] => h1:fresh",
            "p[style-name='Heading 2'] => h2:fresh",
            "p[style-name='Heading 3'] => h3:fresh",
            "r[style-name='Strong'] => strong",
            "r[style-name='Emphasis'] => em",
        ],
    };
    const result = await mammoth.convertToHtml({ path: absolutePath }, options);
    return {
        html: result.value,
        messages: result.messages,
        warnings: result.messages.filter((m) => m.type === 'warning'),
        errors: result.messages.filter((m) => m.type === 'error'),
    };
}
async function analyzeStructure(file_path) {
    const absolutePath = path.resolve(file_path);
    if (!fs.existsSync(absolutePath)) {
        throw new Error(`File not found: ${absolutePath}`);
    }
    const htmlResult = await mammoth.convertToHtml({ path: absolutePath });
    const html = htmlResult.value;
    const textResult = await mammoth.extractRawText({ path: absolutePath });
    const text = textResult.value;
    const headings = (html.match(/<h[1-6][^>]*>.*?<\/h[1-6]>/gi) || []).map((h) => ({
        level: parseInt(h.match(/<h([1-6])/)[1]),
        text: h.replace(/<[^>]*>/g, '').trim(),
    }));
    const paragraphs = (html.match(/<p[^>]*>.*?<\/p>/gi) || []).length;
    const strongElements = (html.match(/<strong[^>]*>.*?<\/strong>/gi) || []).length;
    const emElements = (html.match(/<em[^>]*>.*?<\/em>/gi) || []).length;
    const lists = (html.match(/<[uo]l[^>]*>.*?<\/[uo]l>/gi) || []).length;
    const listItems = (html.match(/<li[^>]*>.*?<\/li>/gi) || []).length;
    return {
        document_stats: {
            total_characters: text.length,
            total_words: text.split(/\s+/).filter((word) => word.length > 0).length,
            total_paragraphs: paragraphs,
            total_headings: headings.length,
        },
        structure: {
            headings,
            heading_levels: [...new Set(headings.map((h) => h.level))].sort(),
        },
        formatting: {
            bold_elements: strongElements,
            italic_elements: emElements,
            lists,
            list_items: listItems,
        },
        messages: htmlResult.messages,
    };
}
async function extractImages(file_path, output_dir) {
    const absolutePath = path.resolve(file_path);
    if (!fs.existsSync(absolutePath)) {
        throw new Error(`File not found: ${absolutePath}`);
    }
    const options = {
        convertImage: mammoth.images.imgElement(function (image) {
            if (output_dir) {
                const outputPath = path.resolve(output_dir);
                if (!fs.existsSync(outputPath)) {
                    fs.mkdirSync(outputPath, { recursive: true });
                }
                const imagePath = path.join(outputPath, `image_${Date.now()}_${Math.random().toString(36).substr(2, 9)}.${image.contentType.split('/')[1]}`);
                return image.read().then(function (imageBuffer) {
                    fs.writeFileSync(imagePath, imageBuffer);
                    return { src: imagePath, alt: image.altText || 'Extracted image' };
                });
            }
            else {
                return image.read().then(function (imageBuffer) {
                    return {
                        src: `data:${image.contentType};base64,${imageBuffer.toString('base64')}`,
                        alt: image.altText || 'Embedded image',
                        size: imageBuffer.length,
                    };
                });
            }
        }),
    };
    const result = await mammoth.convertToHtml({ path: absolutePath }, options);
    const images = (result.value.match(/<img[^>]*>/gi) || []).map((img) => {
        const srcMatch = img.match(/src="([^"]*)"/);
        const altMatch = img.match(/alt="([^"]*)"/);
        return {
            src: srcMatch ? srcMatch[1] : '',
            alt: altMatch ? altMatch[1] : '',
            is_base64: srcMatch ? srcMatch[1].startsWith('data:') : false,
        };
    });
    return {
        total_images: images.length,
        images,
        output_directory: output_dir || 'Images embedded as base64',
        messages: result.messages,
    };
}
async function convertToMarkdown(file_path) {
    const absolutePath = path.resolve(file_path);
    if (!fs.existsSync(absolutePath)) {
        throw new Error(`File not found: ${absolutePath}`);
    }
    const htmlResult = await mammoth.convertToHtml({ path: absolutePath });
    let html = htmlResult.value;
    let markdown = html
        .replace(/<h1[^>]*>(.*?)<\/h1>/gi, '# $1\n\n')
        .replace(/<h2[^>]*>(.*?)<\/h2>/gi, '## $1\n\n')
        .replace(/<h3[^>]*>(.*?)<\/h3>/gi, '### $1\n\n')
        .replace(/<h4[^>]*>(.*?)<\/h4>/gi, '#### $1\n\n')
        .replace(/<h5[^>]*>(.*?)<\/h5>/gi, '##### $1\n\n')
        .replace(/<h6[^>]*>(.*?)<\/h6>/gi, '###### $1\n\n')
        .replace(/<strong[^>]*>(.*?)<\/strong>/gi, '**$1**')
        .replace(/<b[^>]*>(.*?)<\/b>/gi, '**$1**')
        .replace(/<em[^>]*>(.*?)<\/em>/gi, '*$1*')
        .replace(/<i[^>]*>(.*?)<\/i>/gi, '*$1*')
        .replace(/<ul[^>]*>/gi, '')
        .replace(/<\/ul>/gi, '\n')
        .replace(/<ol[^>]*>/gi, '')
        .replace(/<\/ol>/gi, '\n')
        .replace(/<li[^>]*>(.*?)<\/li>/gi, '- $1\n')
        .replace(/<p[^>]*>(.*?)<\/p>/gi, '$1\n\n')
        .replace(/<br[^>]*>/gi, '\n')
        .replace(/<[^>]*>/g, '')
        .replace(/\n{3,}/g, '\n\n')
        .trim();
    return {
        markdown,
        word_count: markdown.split(/\s+/).filter((word) => word.length > 0).length,
        messages: htmlResult.messages,
    };
}
const server = new McpServer({
    name: 'docx-format-server',
    version: '0.2.0',
});
// Tool to extract text content from DOCX files
server.tool('extract_text', 'Extract plain text content from a DOCX file', {
    file_path: z.string().describe('Path to the .docx file'),
}, async ({ file_path }) => {
    try {
        const result = await extractText(file_path);
        return {
            content: [
                {
                    type: 'text',
                    text: JSON.stringify(result, null, 2),
                },
            ],
        };
    }
    catch (error) {
        return {
            content: [
                {
                    type: 'text',
                    text: `Error extracting text: ${error.message}`,
                },
            ],
            isError: true,
        };
    }
});
// Tool to convert DOCX to HTML with formatting preserved
server.tool('convert_to_html', 'Convert DOCX file to HTML with formatting preserved', {
    file_path: z.string().describe('Path to the .docx file'),
    include_styles: z
        .boolean()
        .optional()
        .describe('Include inline styles (default: true)'),
}, async ({ file_path, include_styles = true }) => {
    try {
        const result = await convertToHtml(file_path, include_styles);
        return {
            content: [
                {
                    type: 'text',
                    text: JSON.stringify(result, null, 2),
                },
            ],
        };
    }
    catch (error) {
        return {
            content: [
                {
                    type: 'text',
                    text: `Error converting to HTML: ${error.message}`,
                },
            ],
            isError: true,
        };
    }
});
// Tool to analyze document structure and formatting
server.tool('analyze_structure', 'Analyze document structure, headings, and formatting elements', {
    file_path: z.string().describe('Path to the .docx file'),
}, async ({ file_path }) => {
    try {
        const result = await analyzeStructure(file_path);
        return {
            content: [
                {
                    type: 'text',
                    text: JSON.stringify(result, null, 2),
                },
            ],
        };
    }
    catch (error) {
        return {
            content: [
                {
                    type: 'text',
                    text: `Error analyzing structure: ${error.message}`,
                },
            ],
            isError: true,
        };
    }
});
// Tool to extract images from DOCX
server.tool('extract_images', 'Extract and list images from a DOCX file', {
    file_path: z.string().describe('Path to the .docx file'),
    output_dir: z
        .string()
        .optional()
        .describe('Directory to save extracted images (optional)'),
}, async ({ file_path, output_dir }) => {
    try {
        const result = await extractImages(file_path, output_dir);
        return {
            content: [
                {
                    type: 'text',
                    text: JSON.stringify(result, null, 2),
                },
            ],
        };
    }
    catch (error) {
        return {
            content: [
                {
                    type: 'text',
                    text: `Error extracting images: ${error.message}`,
                },
            ],
            isError: true,
        };
    }
});
// Tool to convert DOCX to Markdown
server.tool('convert_to_markdown', 'Convert DOCX file to Markdown format', {
    file_path: z.string().describe('Path to the .docx file'),
}, async ({ file_path }) => {
    try {
        const result = await convertToMarkdown(file_path);
        return {
            content: [
                {
                    type: 'text',
                    text: JSON.stringify(result, null, 2),
                },
            ],
        };
    }
    catch (error) {
        return {
            content: [
                {
                    type: 'text',
                    text: `Error converting to Markdown: ${error.message}`,
                },
            ],
            isError: true,
        };
    }
});
// Check if HTTP mode is enabled
const useHttp = process.env.USE_HTTP === 'true';
if (useHttp) {
    // HTTP Server Mode
    const app = express();
    app.use(express.json());
    // HTTP endpoints mirroring MCP tools
    app.post('/extract_text', async (req, res) => {
        try {
            const { file_path } = req.body;
            const result = await extractText(file_path);
            res.json(result);
        }
        catch (error) {
            const message = error instanceof Error ? error.message : String(error);
            res.status(500).json({ error: `Error extracting text: ${message}` });
        }
    });
    app.post('/convert_to_html', async (req, res) => {
        try {
            const { file_path, include_styles = true } = req.body;
            const result = await convertToHtml(file_path, include_styles);
            res.json(result);
        }
        catch (error) {
            const message = error instanceof Error ? error.message : String(error);
            res.status(500).json({ error: `Error converting to HTML: ${message}` });
        }
    });
    app.post('/analyze_structure', async (req, res) => {
        try {
            const { file_path } = req.body;
            const result = await analyzeStructure(file_path);
            res.json(result);
        }
        catch (error) {
            const message = error instanceof Error ? error.message : String(error);
            res.status(500).json({ error: `Error analyzing structure: ${message}` });
        }
    });
    app.post('/extract_images', async (req, res) => {
        try {
            const { file_path, output_dir } = req.body;
            const result = await extractImages(file_path, output_dir);
            res.json(result);
        }
        catch (error) {
            const message = error instanceof Error ? error.message : String(error);
            res.status(500).json({ error: `Error extracting images: ${message}` });
        }
    });
    app.post('/convert_to_markdown', async (req, res) => {
        try {
            const { file_path } = req.body;
            const result = await convertToMarkdown(file_path);
            res.json(result);
        }
        catch (error) {
            const message = error instanceof Error ? error.message : String(error);
            res.status(500).json({ error: `Error converting to Markdown: ${message}` });
        }
    });
    const PORT = process.env.PORT || 3000;
    app.listen(PORT, () => {
        console.log(`Streamable HTTP DOCX server running on port ${PORT}`);
    });
}
else {
    // MCP Stdio Mode
    const transport = new StdioServerTransport();
    await server.connect(transport);
    console.error('Advanced DOCX MCP server running on stdio');
}
