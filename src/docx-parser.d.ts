declare module 'docx-parser' {
  export interface Paragraph {
    text: string
    bold?: boolean
  }

  export function parse(buffer: Buffer): Promise<Paragraph[]>
}
