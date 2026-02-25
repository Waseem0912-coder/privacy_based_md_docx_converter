import React, { useState, useEffect, useRef, useCallback } from 'react'
import MDEditor from '@uiw/react-md-editor'
import { renderAsync } from 'docx-preview'
import { saveAs } from 'file-saver'
import { themes, themeKeys } from './lib/themes.js'
import { convertMarkdownToDocx } from './lib/markdownToDocx.js'

import '@uiw/react-md-editor/markdown-editor.css'
import '@uiw/react-markdown-preview/markdown.css'

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
  const [markdown, setMarkdown] = useState(DEFAULT_MARKDOWN)
  const [selectedTheme, setSelectedTheme] = useState('professional')
  const [isConverting, setIsConverting] = useState(false)
  const [conversionError, setConversionError] = useState(null)

  // ── New UI state ─────────────────────────────────────────────────────────
  const [sidebarOpen, setSidebarOpen] = useState(true)
  const [uiTheme, setUiTheme] = useState('dark')
  const [docTitle, setDocTitle] = useState('My Document')
  const [wordCount, setWordCount] = useState(0)
  const [lineCount, setLineCount] = useState(0)
  const [splitPct, setSplitPct] = useState(50)
  const [isDragging, setIsDragging] = useState(false)

  const previewRef = useRef(null)
  const debounceRef = useRef(null)
  const splitRef = useRef(null)

  // ── Apply UI theme to <html> ─────────────────────────────────────────────
  useEffect(() => {
    document.documentElement.setAttribute('data-ui-theme', uiTheme)
    // react-md-editor reads data-color-mode from the html element
    document.documentElement.setAttribute('data-color-mode', uiTheme)
  }, [uiTheme])

  // ── Word / line counts ───────────────────────────────────────────────────
  useEffect(() => {
    const text = markdown.trim()
    setWordCount(text ? text.split(/\s+/).length : 0)
    setLineCount(markdown.split('\n').length)
  }, [markdown])

  // ── Live preview with debounce ───────────────────────────────────────────
  const runPreview = useCallback(async (md, themeKey) => {
    if (!previewRef.current) return
    setIsConverting(true)
    setConversionError(null)
    try {
      const blob = await convertMarkdownToDocx(md, themes[themeKey])
      await renderAsync(blob, previewRef.current, null, {
        className: 'docx-wrapper',
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
      runPreview(markdown, selectedTheme)
    }, 900)
    return () => {
      if (debounceRef.current) clearTimeout(debounceRef.current)
    }
  }, [markdown, selectedTheme, runPreview])

  // ── Download ─────────────────────────────────────────────────────────────
  const handleDownload = async () => {
    try {
      const blob = await convertMarkdownToDocx(markdown, themes[selectedTheme])
      saveAs(blob, `${docTitle || 'document'}.docx`)
    } catch (err) {
      console.error('Download error:', err)
      alert(`Download failed: ${err.message}`)
    }
  }

  // ── Drag-handle resize ───────────────────────────────────────────────────
  function startDrag(e) {
    e.preventDefault()
    setIsDragging(true)
    const onMove = (ev) => {
      if (!splitRef.current) return
      const rect = splitRef.current.getBoundingClientRect()
      const pct = ((ev.clientX - rect.left) / rect.width) * 100
      setSplitPct(Math.min(75, Math.max(25, pct)))
    }
    const onUp = () => {
      setIsDragging(false)
      document.removeEventListener('mousemove', onMove)
    }
    document.addEventListener('mousemove', onMove)
    document.addEventListener('mouseup', onUp, { once: true })
  }

  // ── Render ───────────────────────────────────────────────────────────────
  return (
    <>
      {/* Top bar */}
      <div className="top-bar">
        <button
          className="top-bar-hamburger"
          onClick={() => setSidebarOpen(o => !o)}
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
          spellCheck={false}
        />
        <div className="top-bar-spacer" />
        <button
          className="btn-download"
          onClick={handleDownload}
          disabled={isConverting}
        >
          ⬇ Download DOCX
        </button>
      </div>

      {/* Main area */}
      <div className="main-area">
        {/* Sidebar */}
        <aside className="sidebar" style={{ width: sidebarOpen ? 'var(--sidebar-w)' : '0' }}>
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
              >
                {themeKeys.map(key => (
                  <option key={key} value={key}>{themes[key].name}</option>
                ))}
              </select>
            </div>

            <div className="sidebar-divider" />

            {/* App Theme section */}
            <div className="sidebar-section">
              <div className="sidebar-section-label">App Theme</div>
              <div className="ui-theme-toggle">
                <button
                  className={`ui-theme-btn${uiTheme === 'dark' ? ' active' : ''}`}
                  onClick={() => setUiTheme('dark')}
                >
                  🌙 Dark
                </button>
                <button
                  className={`ui-theme-btn${uiTheme === 'light' ? ' active' : ''}`}
                  onClick={() => setUiTheme('light')}
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

        {/* Split pane */}
        <div className="split-pane" ref={splitRef}>
          {/* Editor pane */}
          <div className="editor-pane" style={{ width: splitPct + '%' }}>
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
            onMouseDown={startDrag}
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
