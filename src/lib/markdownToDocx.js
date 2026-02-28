import { unified } from 'unified'
import remarkParse from 'remark-parse'
import remarkGfm from 'remark-gfm'
import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  ImageRun,
  Table,
  TableRow,
  TableCell,
  HeadingLevel,
  AlignmentType,
  BorderStyle,
  ShadingType,
  WidthType,
  LevelFormat,
  TableOfContents,
  Header,
  Footer,
  PageNumber,
  ExternalHyperlink,
  convertInchesToTwip,
  VerticalAlign,
  PageOrientation,
} from 'docx'

// ─────────────────────────────────────────────────────────────────────────────
// Image pre-fetch: collect all image URLs → fetch in parallel → return Map
// Offline-first: only data URLs and local/same-origin URLs are fetched.
// Any external internet URL is intentionally skipped.
// ─────────────────────────────────────────────────────────────────────────────

// Max content width in points (6.5 inch content area at 72pt/in)
const MAX_IMG_WIDTH_PT = 450

const SUPPORTED_TYPES = new Set(['png', 'jpg', 'jpeg', 'gif', 'bmp', 'tiff'])

function collectImageUrls(node, urls = new Set()) {
  if (node.type === 'image' && node.url) urls.add(node.url)
  if (node.children) node.children.forEach(c => collectImageUrls(c, urls))
  return urls
}

function getImageDimensions(src) {
  return new Promise((resolve) => {
    const img = new Image()
    img.onload = () => resolve({ w: img.naturalWidth, h: img.naturalHeight })
    img.onerror = () => resolve({ w: 400, h: 300 })
    img.src = src
  })
}

function mimeToType(mime) {
  const sub = mime.split('/')[1]?.split(';')[0]?.toLowerCase() || 'png'
  return sub === 'jpeg' ? 'jpg' : sub
}

function isLoopbackHost(hostname = '') {
  const host = hostname.toLowerCase()
  return host === 'localhost' || host === '127.0.0.1' || host === '::1' || host.endsWith('.localhost')
}

function isLocalOnlyUrl(url) {
  try {
    if (url.startsWith('data:')) return true
    if (url.startsWith('blob:')) return true

    const resolved = new URL(url, window.location.href)
    if (resolved.protocol !== 'http:' && resolved.protocol !== 'https:') return false

    return resolved.origin === window.location.origin || isLoopbackHost(resolved.hostname)
  } catch (_) {
    return false
  }
}

async function prefetchImages(tree) {
  const imageMap = new Map() // url → { data: ArrayBuffer, type: string, widthPt: number, heightPt: number }
  const urls = collectImageUrls(tree)

  await Promise.all([...urls].map(async (url) => {
    try {
      let data, type, srcForDims

      if (url.startsWith('data:')) {
        // Data URL — works offline
        const commaIdx = url.indexOf(',')
        const header = url.slice(0, commaIdx)
        const b64 = url.slice(commaIdx + 1)
        type = mimeToType(header.replace('data:', ''))
        if (!SUPPORTED_TYPES.has(type)) return
        const binary = atob(b64)
        const bytes = new Uint8Array(binary.length)
        for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i)
        data = bytes.buffer
        srcForDims = url
      } else {
        // Local/same-origin URL only (external internet URLs are blocked)
        if (!isLocalOnlyUrl(url)) return

        const resp = await fetch(url)
        if (!resp.ok) return
        type = mimeToType(resp.headers.get('content-type') || 'image/png')
        if (!SUPPORTED_TYPES.has(type)) return
        data = await resp.arrayBuffer()
        srcForDims = URL.createObjectURL(new Blob([data]))
      }

      const { w, h } = await getImageDimensions(srcForDims)
      // Convert pixels → points (assumes 96 DPI screen)
      const wPt = w * 72 / 96
      const hPt = h * 72 / 96
      // Scale down if wider than content area
      const scale = wPt > MAX_IMG_WIDTH_PT ? MAX_IMG_WIDTH_PT / wPt : 1
      imageMap.set(url, {
        data,
        type,
        widthPt: Math.round(wPt * scale),
        heightPt: Math.round(hPt * scale),
      })
    } catch (_) {
      // Unsupported/external URL or load error — fall back to text placeholder
    }
  }))

  return imageMap
}

// ─────────────────────────────────────────────────────────────────────────────
// Utility: extract plain text from any MDAST node tree
// ─────────────────────────────────────────────────────────────────────────────
function extractText(node) {
  if (!node) return ''
  if (node.value != null) return node.value
  if (node.children) return node.children.map(extractText).join('')
  return ''
}

// ─────────────────────────────────────────────────────────────────────────────
// Inline run factory helpers
// ─────────────────────────────────────────────────────────────────────────────
function makeRun(text, overrides = {}) {
  return new TextRun({ text, ...overrides })
}

// ─────────────────────────────────────────────────────────────────────────────
// Process inline MDAST nodes → array of TextRun / ExternalHyperlink / ImageRun
// options: { bold, italic, strike, color, size }
// ─────────────────────────────────────────────────────────────────────────────
function processInline(nodes, theme, opts = {}, imageMap = new Map()) {
  if (!nodes || nodes.length === 0) return []
  const runs = []

  for (const node of nodes) {
    const base = {
      font: theme.font,
      size: opts.size || 22,
      color: opts.color || theme.text,
      bold: opts.bold || false,
      italics: opts.italic || false,
      strike: opts.strike || false,
    }

    switch (node.type) {
      case 'text':
        runs.push(new TextRun({ text: node.value, ...base }))
        break

      case 'strong':
        runs.push(...processInline(node.children, theme, { ...opts, bold: true }, imageMap))
        break

      case 'emphasis':
        runs.push(...processInline(node.children, theme, { ...opts, italic: true }, imageMap))
        break

      case 'delete':
        runs.push(...processInline(node.children, theme, { ...opts, strike: true }, imageMap))
        break

      case 'inlineCode':
        runs.push(new TextRun({
          text: node.value,
          font: theme.codeFont,
          size: 20,
          color: theme.accent,
        }))
        break

      case 'link': {
        const linkRuns = processInline(node.children, theme, {
          ...opts,
          color: theme.secondary,
        }, imageMap)
        // linkRuns might already be ExternalHyperlinks if nested; filter to TextRun only
        const textRunsOnly = linkRuns.filter(r => r instanceof TextRun)
        if (textRunsOnly.length === 0) {
          textRunsOnly.push(new TextRun({ text: extractText(node), font: theme.font, size: opts.size || 22, color: theme.secondary }))
        }
        runs.push(new ExternalHyperlink({ children: textRunsOnly, link: node.url || '#' }))
        break
      }

      case 'image': {
        const img = imageMap.get(node.url)
        if (img) {
          runs.push(new ImageRun({
            data: img.data,
            transformation: { width: img.widthPt, height: img.heightPt },
            type: img.type,
            altText: { title: node.alt || '', description: node.alt || '', name: node.alt || '' },
          }))
        } else {
          runs.push(new TextRun({
            text: `[Image: ${node.alt || node.url || 'image'}]`,
            italics: true,
            color: theme.muted,
            size: 20,
            font: theme.font,
          }))
        }
        break
      }

      case 'break':
        runs.push(new TextRun({ text: '', break: 1 }))
        break

      case 'html':
        // strip HTML nodes
        break

      default:
        // Fallback: recurse into children if present
        if (node.children) {
          runs.push(...processInline(node.children, theme, opts, imageMap))
        }
        break
    }
  }

  return runs
}

// ─────────────────────────────────────────────────────────────────────────────
// Paragraph / heading factory helpers
// ─────────────────────────────────────────────────────────────────────────────
function h1(text, theme) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    children: [new TextRun({ text, bold: true, color: theme.primary, size: 48, font: theme.font })],
    spacing: { before: 400, after: 120 },
  })
}

function h2(text, theme) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    children: [new TextRun({ text, bold: true, color: theme.primary, size: 36, font: theme.font })],
    spacing: { before: 320, after: 100 },
  })
}

function h3(text, theme) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    children: [new TextRun({ text, bold: true, color: theme.secondary, size: 28, font: theme.font })],
    spacing: { before: 240, after: 80 },
  })
}

function h4(text, theme) {
  return new Paragraph({
    children: [new TextRun({ text, bold: true, color: theme.text, size: 24, font: theme.font })],
    spacing: { before: 200, after: 60 },
  })
}

function p(runs, theme) {
  return new Paragraph({
    children: runs,
    spacing: { before: 0, after: 160, line: 276, lineRule: 'auto' },
    style: 'Normal',
  })
}

function spacer(twips = 200) {
  return new Paragraph({ children: [], spacing: { before: twips, after: 0 } })
}

function pgBreak() {
  return new Paragraph({ children: [], pageBreakBefore: true })
}

function hr(theme) {
  return new Paragraph({
    children: [],
    border: {
      bottom: { style: BorderStyle.SINGLE, size: 6, color: theme.border, space: 1 },
    },
    spacing: { before: 240, after: 240 },
  })
}

// ─────────────────────────────────────────────────────────────────────────────
// Code block
// ─────────────────────────────────────────────────────────────────────────────
function codeBlock(lines, lang, theme) {
  const paras = []

  if (lang) {
    paras.push(new Paragraph({
      children: [new TextRun({
        text: ` ${lang} `,
        font: theme.codeFont,
        size: 16,
        bold: true,
        color: theme.white,
        highlight: 'none',
      })],
      shading: { type: ShadingType.CLEAR, color: 'auto', fill: theme.accent },
      spacing: { before: 160, after: 0 },
    }))
  } else {
    paras.push(spacer(160))
  }

  for (const line of lines) {
    paras.push(new Paragraph({
      children: [new TextRun({
        text: line === '' ? ' ' : line,
        font: theme.codeFont,
        size: 18,
        color: theme.text,
      })],
      shading: { type: ShadingType.CLEAR, color: 'auto', fill: theme.bgLight },
      spacing: { before: 0, after: 0 },
      indent: { left: 240 },
    }))
  }

  // Bottom padding row
  paras.push(new Paragraph({
    children: [new TextRun({ text: ' ', font: theme.codeFont, size: 12 })],
    shading: { type: ShadingType.CLEAR, color: 'auto', fill: theme.bgLight },
    spacing: { before: 0, after: 160 },
  }))

  return paras
}

// ─────────────────────────────────────────────────────────────────────────────
// Blockquote callout
// ─────────────────────────────────────────────────────────────────────────────
function blockquoteParagraph(runs, theme) {
  return new Paragraph({
    children: runs,
    border: {
      left: { style: BorderStyle.SINGLE, size: 20, color: theme.info, space: 6 },
    },
    shading: { type: ShadingType.CLEAR, color: 'auto', fill: theme.infoBg },
    indent: { left: 480 },
    spacing: { before: 80, after: 80, line: 276, lineRule: 'auto' },
  })
}

// ─────────────────────────────────────────────────────────────────────────────
// Table builder
// ─────────────────────────────────────────────────────────────────────────────
function makeTable(headers, rows, colWidths, theme) {
  const cellPad = { top: 80, bottom: 80, left: 150, right: 150 }
  const borderOpts = { style: BorderStyle.SINGLE, size: 4, color: theme.border }
  const allBorders = {
    top: borderOpts, bottom: borderOpts, left: borderOpts,
    right: borderOpts, insideHorizontal: borderOpts, insideVertical: borderOpts,
  }

  const headerRow = new TableRow({
    tableHeader: true,
    children: headers.map((header, i) =>
      new TableCell({
        children: [new Paragraph({
          children: [new TextRun({ text: String(header), bold: true, color: theme.white, font: theme.font, size: 20 })],
          alignment: AlignmentType.LEFT,
        })],
        width: { size: colWidths[i], type: WidthType.DXA },
        shading: { type: ShadingType.CLEAR, color: 'auto', fill: theme.primary },
        margins: cellPad,
        verticalAlign: VerticalAlign.CENTER,
      })
    ),
  })

  const dataRows = rows.map((row, rowIdx) =>
    new TableRow({
      children: row.map((cell, cellIdx) =>
        new TableCell({
          children: [new Paragraph({
            children: [new TextRun({ text: String(cell), font: theme.font, size: 20, color: theme.text })],
          })],
          width: { size: colWidths[cellIdx] || Math.floor(9360 / row.length), type: WidthType.DXA },
          shading: {
            type: ShadingType.CLEAR,
            color: 'auto',
            fill: rowIdx % 2 === 0 ? 'FFFFFF' : theme.bgLight,
          },
          margins: cellPad,
          verticalAlign: VerticalAlign.CENTER,
        })
      ),
    })
  )

  return new Table({
    rows: [headerRow, ...dataRows],
    width: { size: 9360, type: WidthType.DXA },
    borders: allBorders,
  })
}

// ─────────────────────────────────────────────────────────────────────────────
// List item processing (recursive for nesting)
// ─────────────────────────────────────────────────────────────────────────────
function processListNode(listNode, theme, depth = 0, out = [], imageMap = new Map()) {
  const level = Math.min(depth, 2) // docx numbering supports 0-2 well
  const isOrdered = listNode.ordered

  for (const item of listNode.children) {
    const paragraphs = item.children.filter(c => c.type === 'paragraph')
    const nestedLists = item.children.filter(c => c.type === 'list')

    if (paragraphs.length > 0) {
      let runs = processInline(paragraphs[0].children, theme, {}, imageMap)

      // Task list checkbox prefix
      if (item.checked === true) {
        runs = [new TextRun({ text: '☑ ', font: theme.font, size: 22, color: theme.success }), ...runs]
      } else if (item.checked === false) {
        runs = [new TextRun({ text: '☐ ', font: theme.font, size: 22, color: theme.muted }), ...runs]
      }

      out.push(new Paragraph({
        children: runs,
        numbering: { reference: isOrdered ? 'number-list' : 'bullet-list', level },
        spacing: { before: 40, after: 40 },
      }))
    }

    // Recurse into nested lists
    for (const nested of nestedLists) {
      processListNode(nested, theme, depth + 1, out, imageMap)
    }
  }

  return out
}

// ─────────────────────────────────────────────────────────────────────────────
// Main MDAST → docx elements walker
// ─────────────────────────────────────────────────────────────────────────────
function mdastToDocxElements(node, theme, imageMap = new Map()) {
  const out = []

  switch (node.type) {
    case 'root':
      for (const child of node.children) {
        out.push(...mdastToDocxElements(child, theme, imageMap))
      }
      break

    case 'heading': {
      const text = extractText(node)
      switch (node.depth) {
        case 1: out.push(h1(text, theme)); break
        case 2: out.push(h2(text, theme)); break
        case 3: out.push(h3(text, theme)); break
        default: out.push(h4(text, theme)); break
      }
      break
    }

    case 'paragraph': {
      const runs = processInline(node.children, theme, {}, imageMap)
      if (runs.length === 0) break
      // Standalone image paragraph: center it
      const isImageOnly = node.children.every(c => c.type === 'image')
      if (isImageOnly) {
        out.push(new Paragraph({
          children: runs,
          alignment: AlignmentType.CENTER,
          spacing: { before: 160, after: 160 },
        }))
      } else {
        out.push(p(runs, theme))
      }
      break
    }

    case 'code': {
      const lang = node.lang || ''
      const lines = node.value.split('\n')
      out.push(...codeBlock(lines, lang, theme))
      break
    }

    case 'blockquote': {
      for (const child of node.children) {
        if (child.type === 'paragraph') {
          const runs = processInline(child.children, theme, { color: theme.info }, imageMap)
          out.push(blockquoteParagraph(runs, theme))
        } else {
          out.push(...mdastToDocxElements(child, theme, imageMap))
        }
      }
      break
    }

    case 'list':
      out.push(...processListNode(node, theme, 0, [], imageMap))
      break

    case 'table': {
      const [headerRow, ...dataRows] = node.children
      const headers = headerRow.children.map(cell => extractText(cell))
      const rows = dataRows.map(row => row.children.map(cell => extractText(cell)))
      const numCols = Math.max(headers.length, 1)
      const colWidth = Math.floor(9360 / numCols)
      out.push(makeTable(headers, rows, Array(numCols).fill(colWidth), theme))
      out.push(spacer(120))
      break
    }

    case 'thematicBreak':
      out.push(hr(theme))
      break

    case 'html':
      // Skip raw HTML blocks
      break

    default:
      if (node.children) {
        for (const child of node.children) {
          out.push(...mdastToDocxElements(child, theme, imageMap))
        }
      }
      break
  }

  return out
}

// ─────────────────────────────────────────────────────────────────────────────
// Numbering config for the document
// ─────────────────────────────────────────────────────────────────────────────
function buildNumberingConfig() {
  const bulletLevels = ['•', '◦', '▪'].map((text, level) => ({
    level,
    format: LevelFormat.BULLET,
    text,
    alignment: AlignmentType.LEFT,
    style: {
      run: { font: 'Arial' },
      paragraph: {
        indent: {
          left: convertInchesToTwip(0.5 + level * 0.25),
          hanging: convertInchesToTwip(0.25),
        },
      },
    },
  }))

  const numberLevels = [1, 2, 3].map((n, level) => ({
    level,
    format: LevelFormat.DECIMAL,
    text: `%${n}.`,
    alignment: AlignmentType.LEFT,
    style: {
      paragraph: {
        indent: {
          left: convertInchesToTwip(0.5 + level * 0.25),
          hanging: convertInchesToTwip(0.25),
        },
      },
    },
  }))

  return {
    config: [
      { reference: 'bullet-list', levels: bulletLevels },
      { reference: 'number-list', levels: numberLevels },
    ],
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Build the full Document (optional TOC/header/footer + content)
// ─────────────────────────────────────────────────────────────────────────────
function buildDocument(contentElements, theme, options = {}) {
  const {
    includeToc = true,
    includeHeader = false,
    includeFooter = true,
    headerTitle = 'Document',
  } = options
  const margin = convertInchesToTwip(1)

  // ── Header ────────────────────────────────────────────────
  const docHeader = new Header({
    children: [
      new Paragraph({
        alignment: AlignmentType.RIGHT,
        children: [new TextRun({ text: headerTitle, italics: true, color: theme.subtle, size: 18, font: theme.font })],
        border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: theme.border, space: 1 } },
      }),
    ],
  })

  // ── Footer ────────────────────────────────────────────────
  const docFooter = new Footer({
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new TextRun({ text: 'Page ', color: theme.subtle, size: 18, font: theme.font }),
          new TextRun({ children: [PageNumber.CURRENT], color: theme.subtle, size: 18, font: theme.font }),
          new TextRun({ text: ' of ', color: theme.subtle, size: 18, font: theme.font }),
          new TextRun({ children: [PageNumber.TOTAL_PAGES], color: theme.subtle, size: 18, font: theme.font }),
        ],
      }),
    ],
  })

  // ── TOC page ──────────────────────────────────────────────
  const tocElements = includeToc
    ? [
        new TableOfContents('Table of Contents', {
          hyperlink: true,
          headingStyleRange: '1-3',
        }),
      ]
    : []

  const documentChildren = []
  if (tocElements.length > 0) {
    documentChildren.push(...tocElements)
    if (contentElements.length > 0) documentChildren.push(pgBreak())
  }
  documentChildren.push(...contentElements)

  // ── Assemble document ─────────────────────────────────────
  return new Document({
    numbering: buildNumberingConfig(),
    styles: {
      default: {
        document: {
          run: { font: theme.font, size: 22, color: theme.text },
          paragraph: { spacing: { line: 276, lineRule: 'auto' } },
        },
      },
      paragraphStyles: [
        {
          id: 'Normal',
          name: 'Normal',
          run: { font: theme.font, size: 22, color: theme.text },
          paragraph: { spacing: { after: 160 } },
        },
      ],
    },
    sections: [
      {
        properties: {
          page: {
            size: {
              orientation: PageOrientation.PORTRAIT,
              width: 12240,
              height: 15840,
            },
            margin: { top: margin, right: margin, bottom: margin, left: margin, footer: 708, header: 708 },
          },
        },
        headers: includeHeader ? { default: docHeader } : undefined,
        footers: includeFooter ? { default: docFooter } : undefined,
        children: documentChildren,
      },
    ],
  })
}

// ─────────────────────────────────────────────────────────────────────────────
// Public API
// ─────────────────────────────────────────────────────────────────────────────
export async function convertMarkdownToDocx(markdownText, theme, options = {}) {
  // Parse markdown to MDAST
  const tree = unified()
    .use(remarkParse)
    .use(remarkGfm)
    .parse(markdownText)

  let firstHeadingTitle = 'Document'
  for (const node of tree.children) {
    if (node.type === 'heading' && node.depth === 1) {
      firstHeadingTitle = extractText(node) || firstHeadingTitle
      break
    }
  }

  // Pre-fetch all images before the synchronous walk
  const imageMap = await prefetchImages(tree)

  // Convert all MDAST nodes to docx elements
  const contentElements = mdastToDocxElements(tree, theme, imageMap)

  // Build and pack the document
  const doc = buildDocument(contentElements, theme, {
    ...options,
    headerTitle: options.headerTitle?.trim() || firstHeadingTitle,
  })
  return Packer.toBlob(doc)
}
