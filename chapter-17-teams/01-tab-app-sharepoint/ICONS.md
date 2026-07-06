# Teams app icons

Teams requires **two PNGs** at the root of the .zip. They cannot be omitted; the upload fails without them.

## Specifications

| File | Size | Format |
|---|---|---|
| `color.png` | 192×192 px | Color, PNG, no transparency outside the icon |
| `outline.png` | 32×32 px | Monochrome (white on transparent), PNG with transparency |

## How to generate them quickly

With ImageMagick (or any editor):

```bash
# color.png — rellena de color corporativo con texto
convert -size 192x192 xc:'#107c10' -fill white -gravity center \
  -pointsize 72 -annotate +0+0 'SP' color.png

# outline.png — silueta monocroma transparente
convert -size 32x32 xc:none -fill white -gravity center \
  -pointsize 16 -annotate +0+0 'SP' outline.png
```

Or use the official [Microsoft Teams App Icon Generator](https://developer.microsoft.com/microsoft-teams/app-icon-tool).

> **Honest note**: icons are a step people underestimate. The `outline.png` **must** be a transparent monochrome (white on transparent), not a scaled-down color.png — Teams displays it in dark themes and breaks if it has an opaque background.