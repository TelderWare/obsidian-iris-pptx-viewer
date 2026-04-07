# Obsidian PPTX Viewer

View PowerPoint (`.pptx`) files directly inside Obsidian with slide navigation and faithful rendering of text, shapes, and images.

## Features

- **Native file view** — `.pptx` files open in a dedicated viewer tab, just like PDFs or images
- **Slide navigation** — Prev/Next buttons + keyboard arrows (← → ↑ ↓ Space)
- **Text rendering** — Preserves font size, bold, italic, underline, text color, alignment
- **Image support** — Embedded PNG, JPEG, GIF images displayed in position
- **Shape rendering** — Background fills, borders, rotation
- **Responsive scaling** — Slides scale to fit the available pane

## Building

```bash
# Install dependencies
npm install

# Build the plugin
npm run build

# Or watch for changes during development
npm run dev
```

## Installation

1. Build the plugin (see above)
2. Create a folder in your vault's plugins directory:
   ```
   <your-vault>/.obsidian/plugins/pptx-viewer/
   ```
3. Copy these three files into that folder:
   - `main.js` (the compiled output)
   - `manifest.json`
   - `styles.css`
4. Restart Obsidian (or reload without cache: Ctrl/Cmd+Shift+R)
5. Go to **Settings → Community Plugins** and enable **PPTX Viewer**

## Usage

Simply place a `.pptx` file anywhere in your vault and click it in the file explorer. It will open in the PPTX Viewer automatically.

## Limitations

- **Slide layouts/masters**: Theme colors from slide masters are not fully resolved (scheme colors fall back to defaults)
- **Charts**: Embedded Excel charts are not rendered
- **SmartArt**: SmartArt diagrams are not rendered
- **Animations/transitions**: Not supported (static rendering only)
- **EMF/WMF images**: These legacy Windows formats may not display in all environments

## License

MIT
