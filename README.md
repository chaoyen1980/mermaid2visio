# Mermaid2Visio

![License](https://img.shields.io/badge/license-ISC-blue)
![Build Status](https://github.com/chaoyen1980/mermaid2visio/actions/workflows/ci.yml/badge.svg)
![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20Mac%20%7C%20Linux-lightgrey)
![Status](https://img.shields.io/badge/status-Active-green)

> **Fork notice.** This is a fork of [tropikandy/mermaid2visio](https://github.com/tropikandy/mermaid2visio) that fixes a class of issues which made the upstream output get rejected by Visio ("bad format or corrupted") and render blank in LibreOffice. See **[What's different in this fork](#whats-different-in-this-fork)** below for the full list of fixes.

**Mermaid2Visio** is a modern, cross-platform utility that converts [MermaidJS](https://mermaid.js.org/) diagrams into native, editable Microsoft Visio (`.vsdx`) files.

Unlike legacy tools that rely on a local installation of Visio (COM automation), this project is built on **Node.js** and generates the VSDX XML structure directly. This means it runs on **Windows, macOS, and Linux**, and produces high-fidelity files without requiring Visio to be installed on the machine performing the conversion.

## What's different in this fork

The upstream output had several issues that surfaced in real Visio and LibreOffice: structurally invalid OPC packages, formula expressions written into numeric-only attributes, and shape geometries that LibreOffice rendered as blank. This fork addresses them:

- **Structurally valid VSDX.** Adds the missing OPC parts (`docProps/app.xml`, `docProps/core.xml`, `docProps/custom.xml`, `visio/theme/theme1.xml`, `visio/pages/_rels/page1.xml.rels`), fixes the `FaceName Panos` → `Panose` typo, and enables DEFLATE compression. The package now passes LibreOffice's `libvisio2` validator (the prior package was rejected outright).
- **`V` vs `F` for geometry formulas.** Cells like `<Cell N="X" V="Width*0.5"/>` violate the Visio schema — `V` must be numeric, formulas belong in `F`. Every Geometry/Connection row now goes through an `evalFormula()` helper that emits `V="numeric" F="formula"`. This single fix unblocked shape rendering.
- **`LocPinX` / `LocPinY` on every shape.** Without these, Visio treats `PinX`/`PinY` as the *bottom-left* of the shape instead of the *center*, shifting every shape by half its own size. Fixed for nodes, clusters, and labels.
- **Cluster titles render natively.** Title is placed back inside the cluster shape with `VerticalAlign=0`, `HorzAlign=0`, bold, and small margins — no fragile floating sibling label.
- **Cylinder geometry.** Rewritten as one closed silhouette (top half-ellipse out, sides, bottom half-ellipse out) plus a no-fill inner arc for the 3D rim lip. Renders correctly in both Visio and LibreOffice.
- **Text overflow.** Tight 0.02" margins on every node, plus a parser pass that expands shape height to `lineCount × fontPx × 1.3 + padding` so multi-line labels no longer get clipped.
- **Center-aligned node text.** Mermaid's `foreignObject` labels report `textAlign='left'` by default; this fork forces center-horizontal + middle-vertical on all node shapes.
- **Edge label positioning.** Mermaid places edge labels via `transform=translate(cx, cy)` (center), but the upstream `vsdx.ts` was treating that as top-left and offsetting by `w/2, h/2`. Labels now sit where Mermaid intended.
- **Label overlap.** Post-process pass detects labels whose bboxes intersect any node or another label and shifts them vertically (alternating down/up) until clear.
- **Sequence diagrams.** Added a `sequenceDiagram` parser branch (actor rects + lifelines + message arrows + sequence numbers) so sequence diagrams produce real shapes instead of an empty page.

## Key Advantages

- **No Visio License Required**: Built on a modern Node.js architecture that generates native `.vsdx` XML directly. This means you don't need to buy or install Microsoft Visio to perform conversions.
- **Cross-Platform**: Runs seamlessly on **Windows, macOS, and Linux**.
- **AI-Native Integration**: Includes a built-in **Model Context Protocol (MCP)** server. Connect it to AI agents (like Claude Desktop) to give them the "skill" to generate professional diagrams for you.
- **Smart Glue & Dynamic Routing**: Features an advanced routing engine that creates "Smart Glue" connectors. When you open the file in Visio and move a shape, the lines follow and reroute automatically.
- **Web-Based Visual Editor**: Comes with a local web GUI for instant previewing and one-click downloads.

## Installation

### Prerequisites
- [Node.js](https://nodejs.org/) (v18 or later)

### Setup
```bash
git clone https://github.com/chaoyen1980/mermaid2visio.git
cd mermaid2visio
npm install
npm run build
```

## Usage

### 1. Web GUI (Recommended)
Visual editor with live preview.
```bash
node dist/gui.js
```
Opens `http://localhost:3000` in your browser. Paste Mermaid code, verify the preview, and download the `.vsdx`.

### 2. Command Line (CLI)
Convert files in bulk or via scripts.
```bash
node dist/index.js input.mmd [output.vsdx]
```

### 3. AI Agent Integration (MCP)
Add this tool to your AI assistant (e.g., Claude Desktop) to give it "Visio Skills".

**Configuration (`claude_desktop_config.json`):**
```json
{
  "mcpServers": {
    "mermaid2visio": {
      "command": "node",
      "args": [
        "/absolute/path/to/mermaid2visio/dist/server.js"
      ]
    }
  }
}
```
**Prompt:** *"Generate a system architecture diagram for a cloud app and save it as a Visio file."*

### 4. Windows Context Menu
Right-click any `.mmd` or `.md` file to convert.
1. Run `install_context_menu.bat` as Administrator.
2. Right-click a file -> **Convert to Visio**.

## Supported Features

- **Flowcharts** (`graph TD`, `LR`, etc.)
- **Subgraphs** (mapped to Containers)
- **Shapes**:
  - Rectangle `[]`
  - Rounded `()`
  - Cylinder `[()]` (Database)
  - Rhombus `{}` (Decision)
  - Stadium `([])`
  - Subroutine `[[]]`
  - Circle `(())`
- **Styling**: `fill`, `stroke`, `stroke-width`, `stroke-dasharray`, `color`
- **Interactivity**: Hyperlinks (`click` directive)

## License
ISC