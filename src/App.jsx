import React, { useState, useEffect, useRef, useCallback } from 'react'
import MDEditor from '@uiw/react-md-editor'
import { renderAsync } from 'docx-preview'
import { saveAs } from 'file-saver'
import { themes, themeKeys } from './lib/themes.js'
import { convertMarkdownToDocx } from './lib/markdownToDocx.js'

import '@uiw/react-md-editor/markdown-editor.css'
import '@uiw/react-markdown-preview/markdown.css'

const STORAGE_KEYS = {
  uiTheme: 'md-docx-ui-theme',
  docTheme: 'md-docx-doc-theme',
  sidebarOpen: 'md-docx-sidebar-open',
  splitPct: 'md-docx-split-pct',
  markdown: 'md-docx-markdown',
  docTitle: 'md-docx-doc-title',
  docxLayout: 'md-docx-layout',
}

const MOBILE_BREAKPOINT = 980

// ─────────────────────────────────────────────────────────────────────────────
// Default example markdown
// ─────────────────────────────────────────────────────────────────────────────
const DEFAULT_MARKDOWN = `# My Amazing Document

Welcome to the **MD → DOCX Converter**! Edit markdown on the left and watch the DOCX preview update live on the right.

## Text Formatting

This paragraph has **bold text**, *italic text*, ***bold and italic combined***, ~~strikethrough~~, and \`inline code\`. You can also include [links](https://example.com).

## Lists

**Unordered list with nesting:**

- Item one
  - Nested item 1.1
  - Nested item 1.2
    - Deeply nested 1.2.1
- Item two
- Item three

**Ordered list with nesting:**

1. First step
2. Second step
   1. Sub-step A
   2. Sub-step B
3. Third step

**Task list:**

- [x] Set up the project
- [x] Write the converter
- [ ] Open the document in Word
- [ ] Share with team

## Code Example

\`\`\`python
def fibonacci(n):
    if n <= 1:
        return n
    return fibonacci(n - 1) + fibonacci(n - 2)

# Print first 10 Fibonacci numbers
for i in range(10):
    print(f"F({i}) = {fibonacci(i)}")
\`\`\`

## Data Table

| Feature        | Supported | Notes                          |
|----------------|-----------|--------------------------------|
| Headings H1–H3 | ✓         | With theme colors              |
| Bold / Italic  | ✓         | Including combinations         |
| Tables         | ✓         | Styled with theme              |
| Code blocks    | ✓         | Language label + shading       |
| Task lists     | ✓         | ☑ / ☐ checkboxes              |
| Nested lists   | ✓         | Up to 3 levels                 |
| Blockquotes    | ✓         | Left-border callout style      |
| Strikethrough  | ✓         | \`~~like this~~\`              |
| Links          | ✓         | External hyperlinks            |

## Blockquote

> This is a blockquote. It supports **bold** and *italic* text inside it. Use blockquotes to highlight important notes or quotes from other sources.

## Horizontal Rule

The line below is a thematic break:

---

That's it! Use the **Theme** dropdown in the sidebar to switch color schemes, then click **Download DOCX** to export the document.
`

// ─────────────────────────────────────────────────────────────────────────────
// App
// ─────────────────────────────────────────────────────────────────────────────
export default function App() {
  const [markdown, setMarkdown] = useState(() => {
    const saved = sessionStorage.getItem(STORAGE_KEYS.markdown)
    return saved != null ? saved : DEFAULT_MARKDOWN
  })
  const [selectedTheme, setSelectedTheme] = useState(() => {
    const saved = sessionStorage.getItem(STORAGE_KEYS.docTheme)
    return saved && themeKeys.includes(saved) ? saved : 'professional'
  })
  const [isConverting, setIsConverting] = useState(false)
  const [conversionError, setConversionError] = useState(null)

  // ── New UI state ─────────────────────────────────────────────────────────
  const [sidebarOpen, setSidebarOpen] = useState(() => {
    const saved = sessionStorage.getItem(STORAGE_KEYS.sidebarOpen)
    if (saved != null) return saved === 'true'
    return window.innerWidth > MOBILE_BREAKPOINT
  })
  const [uiTheme, setUiTheme] = useState(() => {
    const saved = sessionStorage.getItem(STORAGE_KEYS.uiTheme)
    return saved === 'light' ? 'light' : 'dark'
  })
  const [docTitle, setDocTitle] = useState(() => {
    const saved = sessionStorage.getItem(STORAGE_KEYS.docTitle)
    return saved != null ? saved : 'My Document'
  })
  const [docxLayout, setDocxLayout] = useState(() => {
    const defaults = { includeToc: true, includeHeader: false, includeFooter: true }
    const saved = sessionStorage.getItem(STORAGE_KEYS.docxLayout)
    if (!saved) return defaults
    try {
      const parsed = JSON.parse(saved)
      return {
        includeToc: typeof parsed.includeToc === 'boolean' ? parsed.includeToc : defaults.includeToc,
        includeHeader: typeof parsed.includeHeader === 'boolean' ? parsed.includeHeader : defaults.includeHeader,
        includeFooter: typeof parsed.includeFooter === 'boolean' ? parsed.includeFooter : defaults.includeFooter,
      }
    } catch (_) {
      return defaults
    }
  })
  const [wordCount, setWordCount] = useState(0)
  const [lineCount, setLineCount] = useState(0)
  const [splitPct, setSplitPct] = useState(() => {
    const saved = Number(sessionStorage.getItem(STORAGE_KEYS.splitPct))
    if (Number.isFinite(saved) && saved >= 25 && saved <= 75) return saved
    return 50
  })
  const [isDragging, setIsDragging] = useState(false)
  const [copyStatus, setCopyStatus] = useState('idle')
  const [isNarrowViewport, setIsNarrowViewport] = useState(
    window.innerWidth <= MOBILE_BREAKPOINT,
  )

  const previewRef = useRef(null)
  const debounceRef = useRef(null)
  const copyResetRef = useRef(null)
  const splitRef = useRef(null)

  // ── Apply UI theme to <html> ─────────────────────────────────────────────
  useEffect(() => {
    document.documentElement.setAttribute('data-ui-theme', uiTheme)
    // react-md-editor reads data-color-mode from the html element
    document.documentElement.setAttribute('data-color-mode', uiTheme)
    sessionStorage.setItem(STORAGE_KEYS.uiTheme, uiTheme)
  }, [uiTheme])

  // ── Word / line counts ───────────────────────────────────────────────────
  useEffect(() => {
    const text = markdown.trim()
    const trimmedLines = markdown.replace(/\s+$/, '')
    setWordCount(text ? text.split(/\s+/).length : 0)
    setLineCount(trimmedLines ? trimmedLines.split('\n').length : 0)
    sessionStorage.setItem(STORAGE_KEYS.markdown, markdown)
  }, [markdown])

  useEffect(() => {
    sessionStorage.setItem(STORAGE_KEYS.docTheme, selectedTheme)
  }, [selectedTheme])

  useEffect(() => {
    sessionStorage.setItem(STORAGE_KEYS.sidebarOpen, String(sidebarOpen))
  }, [sidebarOpen])

  useEffect(() => {
    sessionStorage.setItem(STORAGE_KEYS.splitPct, String(splitPct))
  }, [splitPct])

  useEffect(() => {
    sessionStorage.setItem(STORAGE_KEYS.docTitle, docTitle)
  }, [docTitle])

  useEffect(() => {
    sessionStorage.setItem(STORAGE_KEYS.docxLayout, JSON.stringify(docxLayout))
  }, [docxLayout])

  useEffect(() => {
    const onResize = () => setIsNarrowViewport(window.innerWidth <= MOBILE_BREAKPOINT)
    window.addEventListener('resize', onResize)
    return () => window.removeEventListener('resize', onResize)
  }, [])

  useEffect(() => {
    document.body.classList.toggle('is-dragging', isDragging)
    return () => document.body.classList.remove('is-dragging')
  }, [isDragging])

  useEffect(() => () => {
    if (copyResetRef.current) clearTimeout(copyResetRef.current)
  }, [])

  // ── Live preview with debounce ───────────────────────────────────────────
  const runPreview = useCallback(async (md, themeKey, layout, title) => {
    if (!previewRef.current) return

    if (!md.trim()) {
      setConversionError(null)
      setIsConverting(false)
      previewRef.current.replaceChildren()
      return
    }

    setIsConverting(true)
    setConversionError(null)
    try {
      const blob = await convertMarkdownToDocx(md, themes[themeKey], {
        ...layout,
        headerTitle: title,
      })
      await renderAsync(blob, previewRef.current, null, {
        inWrapper: true,
        ignoreWidth: true,
        ignoreHeight: false,
        ignoreFonts: false,
        breakPages: true,
        ignoreLastRenderedPageBreak: true,
        experimental: false,
        trimXmlDeclaration: true,
        debug: false,
      })
    } catch (err) {
      console.error('Preview error:', err)
      setConversionError(err.message || 'Conversion failed.')
    } finally {
      setIsConverting(false)
    }
  }, [])

  useEffect(() => {
    if (debounceRef.current) clearTimeout(debounceRef.current)
    debounceRef.current = setTimeout(() => {
      runPreview(markdown, selectedTheme, docxLayout, docTitle)
    }, 900)
    return () => {
      if (debounceRef.current) clearTimeout(debounceRef.current)
    }
  }, [markdown, selectedTheme, docxLayout, docTitle, runPreview])

  function flashCopyStatus(status) {
    setCopyStatus(status)
    if (copyResetRef.current) clearTimeout(copyResetRef.current)
    copyResetRef.current = setTimeout(() => {
      setCopyStatus('idle')
    }, 1800)
  }

  async function handleCopyPreview() {
    if (!previewRef.current) return
    const pages = Array.from(previewRef.current.querySelectorAll('section.docx, section.docx-wrapper'))
    if (pages.length === 0) {
      flashCopyStatus('error')
      return
    }

    const styles = Array.from(previewRef.current.querySelectorAll('style'))
      .map((node) => node.textContent || '')
      .filter(Boolean)
      .join('\n')
    const pagesHtml = pages.map((page) => page.outerHTML).join('')
    const plainText = pages
      .map((page) => page.innerText || page.textContent || '')
      .filter(Boolean)
      .join('\n\n')
    const richHtml = `<!doctype html><html><head><meta charset="utf-8">${styles ? `<style>${styles}</style>` : ''}</head><body>${pagesHtml}</body></html>`

    try {
      if (navigator.clipboard?.write && window.ClipboardItem) {
        const item = new ClipboardItem({
          'text/html': new Blob([richHtml], { type: 'text/html' }),
          'text/plain': new Blob([plainText], { type: 'text/plain' }),
        })
        await navigator.clipboard.write([item])
      } else if (navigator.clipboard?.writeText) {
        await navigator.clipboard.writeText(plainText)
      } else {
        const textarea = document.createElement('textarea')
        textarea.value = plainText
        textarea.setAttribute('readonly', '')
        textarea.style.position = 'fixed'
        textarea.style.opacity = '0'
        document.body.appendChild(textarea)
        textarea.select()
        document.execCommand('copy')
        document.body.removeChild(textarea)
      }
      flashCopyStatus('copied')
    } catch (err) {
      console.error('Copy failed:', err)
      flashCopyStatus('error')
    }
  }

  // ── Download ─────────────────────────────────────────────────────────────
  const handleDownload = async () => {
    if (!markdown.trim()) return

    try {
      const blob = await convertMarkdownToDocx(markdown, themes[selectedTheme], {
        ...docxLayout,
        headerTitle: docTitle,
      })
      saveAs(blob, `${docTitle || 'document'}.docx`)
    } catch (err) {
      console.error('Download error:', err)
      alert(`Download failed: ${err.message}`)
    }
  }

  // ── Drag-handle resize ───────────────────────────────────────────────────
  function updateSplit(clientX) {
    if (!splitRef.current) return
    const rect = splitRef.current.getBoundingClientRect()
    const pct = ((clientX - rect.left) / rect.width) * 100
    setSplitPct(Math.min(75, Math.max(25, pct)))
  }

  function startDrag(e) {
    if (isNarrowViewport) return
    e.preventDefault()
    e.currentTarget.setPointerCapture?.(e.pointerId)
    setIsDragging(true)

    const onMove = (ev) => updateSplit(ev.clientX)
    const stopDrag = () => {
      setIsDragging(false)
      document.removeEventListener('pointermove', onMove)
      document.removeEventListener('pointerup', stopDrag)
      document.removeEventListener('pointercancel', stopDrag)
    }
    document.addEventListener('pointermove', onMove)
    document.addEventListener('pointerup', stopDrag)
    document.addEventListener('pointercancel', stopDrag)
  }

  function handleResizeKeyDown(e) {
    if (isNarrowViewport) return
    if (e.key === 'ArrowLeft') {
      e.preventDefault()
      setSplitPct((pct) => Math.max(25, pct - 2))
    } else if (e.key === 'ArrowRight') {
      e.preventDefault()
      setSplitPct((pct) => Math.min(75, pct + 2))
    }
  }

  // ── Render ───────────────────────────────────────────────────────────────
  return (
    <>
      {/* Top bar */}
      <div className="top-bar">
        <button
          className="top-bar-hamburger"
          type="button"
          onClick={() => setSidebarOpen(o => !o)}
          aria-label={sidebarOpen ? 'Collapse sidebar' : 'Expand sidebar'}
          title={sidebarOpen ? 'Collapse sidebar' : 'Expand sidebar'}
        >
          ☰
        </button>
        <span className="top-bar-title">MD → DOCX</span>
        <input
          className="top-bar-doc-title"
          value={docTitle}
          onChange={e => setDocTitle(e.target.value)}
          title="Document title"
          aria-label="Document title"
          spellCheck={false}
        />
        <div className="top-bar-spacer" />
        <button
          className={`btn-copy${copyStatus === 'copied' ? ' copied' : ''}${copyStatus === 'error' ? ' error' : ''}`}
          type="button"
          onClick={handleCopyPreview}
          disabled={isConverting || !markdown.trim()}
          title="Copy DOCX preview content"
        >
          {copyStatus === 'copied' ? '✓ Copied' : copyStatus === 'error' ? 'Copy failed' : '⎘ Copy'}
        </button>
        <button
          className="btn-download"
          type="button"
          onClick={handleDownload}
          disabled={isConverting || !markdown.trim()}
        >
          ⬇ Download DOCX
        </button>
      </div>

      {/* Main area */}
      <div className="main-area">
        {/* Sidebar */}
        <aside
          className="sidebar"
          style={{ width: sidebarOpen ? 'var(--sidebar-w)' : '0' }}
          aria-hidden={!sidebarOpen}
        >
          <div className="sidebar-inner">
            <div className="sidebar-logo">
              <span className="sidebar-logo-icon">📄</span>
              MD → DOCX
            </div>

            {/* Document section */}
            <div className="sidebar-section">
              <div className="sidebar-section-label">Document</div>
              <input
                className="sidebar-title-input"
                value={docTitle}
                onChange={e => setDocTitle(e.target.value)}
                placeholder="Document title"
                aria-label="Sidebar document title"
                spellCheck={false}
              />
            </div>

            <div className="sidebar-divider" />

            {/* DOCX Theme section */}
            <div className="sidebar-section">
              <div className="sidebar-section-label">DOCX Theme</div>
              <select
                className="sidebar-theme-select"
                value={selectedTheme}
                onChange={e => setSelectedTheme(e.target.value)}
                aria-label="DOCX theme"
              >
                {themeKeys.map(key => (
                  <option key={key} value={key}>{themes[key].name}</option>
                ))}
              </select>
            </div>

            <div className="sidebar-divider" />

            {/* DOCX Layout section */}
            <div className="sidebar-section">
              <div className="sidebar-section-label">DOCX Layout</div>
              <div className="sidebar-layout-options">
                <label className="sidebar-checkbox-row">
                  <input
                    type="checkbox"
                    checked={docxLayout.includeToc}
                    onChange={(e) => setDocxLayout((prev) => ({ ...prev, includeToc: e.target.checked }))}
                  />
                  <span>Table of Contents</span>
                </label>
                <label className="sidebar-checkbox-row">
                  <input
                    type="checkbox"
                    checked={docxLayout.includeHeader}
                    onChange={(e) => setDocxLayout((prev) => ({ ...prev, includeHeader: e.target.checked }))}
                  />
                  <span>Header</span>
                </label>
                <label className="sidebar-checkbox-row">
                  <input
                    type="checkbox"
                    checked={docxLayout.includeFooter}
                    onChange={(e) => setDocxLayout((prev) => ({ ...prev, includeFooter: e.target.checked }))}
                  />
                  <span>Footer (page numbers)</span>
                </label>
              </div>
            </div>

            <div className="sidebar-divider" />

            {/* App Theme section */}
            <div className="sidebar-section">
              <div className="sidebar-section-label">App Theme</div>
              <div className="ui-theme-toggle">
                <button
                  className={`ui-theme-btn${uiTheme === 'dark' ? ' active' : ''}`}
                  onClick={() => setUiTheme('dark')}
                  type="button"
                >
                  🌙 Dark
                </button>
                <button
                  className={`ui-theme-btn${uiTheme === 'light' ? ' active' : ''}`}
                  onClick={() => setUiTheme('light')}
                  type="button"
                >
                  ☀ Light
                </button>
              </div>
            </div>

            <div className="sidebar-divider" />

            {/* Stats section */}
            <div className="sidebar-stats">
              <div className="sidebar-section-label">Stats</div>
              <div className="sidebar-stat-row">
                <span>Words</span>
                <span className="sidebar-stat-val">{wordCount.toLocaleString()}</span>
              </div>
              <div className="sidebar-stat-row">
                <span>Lines</span>
                <span className="sidebar-stat-val">{lineCount.toLocaleString()}</span>
              </div>
            </div>
          </div>
        </aside>
        {isNarrowViewport && sidebarOpen ? (
          <button
            className="sidebar-backdrop"
            type="button"
            aria-label="Close sidebar"
            onClick={() => setSidebarOpen(false)}
          />
        ) : null}

        {/* Split pane */}
        <div className="split-pane" ref={splitRef}>
          {/* Editor pane */}
          <div className="editor-pane" style={{ width: isNarrowViewport ? '100%' : splitPct + '%' }}>
            <div className="pane-label">Markdown</div>
            <MDEditor
              value={markdown}
              onChange={val => setMarkdown(val || '')}
              height="100%"
              preview="edit"
              hideToolbar={false}
              visibleDragbar={false}
              data-color-mode={uiTheme}
            />
          </div>

          {/* Drag handle */}
          <div
            className={`drag-handle${isDragging ? ' dragging' : ''}`}
            onPointerDown={startDrag}
            onDoubleClick={() => setSplitPct(50)}
            onKeyDown={handleResizeKeyDown}
            role="separator"
            tabIndex={isNarrowViewport ? -1 : 0}
            aria-orientation="vertical"
            aria-valuemin={25}
            aria-valuemax={75}
            aria-valuenow={Math.round(splitPct)}
            aria-label="Resize editor and preview"
          />

          {/* Preview pane */}
          <div className="preview-pane">
            <div className="pane-label">
              DOCX Preview
              {isConverting && <span className="converting-dot">●</span>}
            </div>
            <div className="preview-scroll">
              {conversionError ? (
                <div className="preview-state">
                  <div className="error-box">
                    <strong>Conversion error</strong>
                    <br />
                    {conversionError}
                  </div>
                </div>
              ) : isConverting ? (
                <div className="preview-state">
                  <div className="spinner" />
                  <span>Rendering DOCX…</span>
                </div>
              ) : null}
              <div ref={previewRef} className="docx-render-container" />
            </div>
          </div>
        </div>
      </div>
    </>
  )
}
