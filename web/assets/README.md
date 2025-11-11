# Icon Assets

This folder should contain icon files for the Excel add-in manifest.

## Required Files

- `icon-16.png` (16x16 pixels)
- `icon-32.png` (32x32 pixels)
- `icon-64.png` (64x64 pixels)
- `icon-80.png` (80x80 pixels)

## Quick Solution

For a prototype, you can use simple emoji icons or text:

1. Go to https://favicon.io/emoji-favicons/
2. Generate favicons with "📊" emoji
3. Download and rename to the sizes above

Or create simple PNG files with:
- OpenGov logo
- "OG" text
- Any company branding

## Example using ImageMagick

```bash
# Create simple colored squares (temporary)
convert -size 16x16 xc:#0078d4 icon-16.png
convert -size 32x32 xc:#0078d4 icon-32.png
convert -size 64x64 xc:#0078d4 icon-64.png
convert -size 80x80 xc:#0078d4 icon-80.png
```

These icons are referenced in `manifest.xml` and should be accessible at:
- https://YOUR-APP.onrender.com/assets/icon-16.png
- https://YOUR-APP.onrender.com/assets/icon-32.png
- etc.

