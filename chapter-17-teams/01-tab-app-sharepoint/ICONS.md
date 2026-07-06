# Iconos de la app de Teams

Teams requiere **dos PNG** en la raíz del .zip. No se pueden omitir; la subida falla sin ellos.

## Especificaciones

| Archivo | Tamaño | Formato |
|---|---|---|
| `color.png` | 192×192 px | Color, PNG, sin transparencia fuera del icono |
| `outline.png` | 32×32 px | Monocromo (blanco sobre transparente), PNG con transparencia |

## Cómo generarlos rápido

Con ImageMagick (o cualquier editor):

```bash
# color.png — rellena de color corporativo con texto
convert -size 192x192 xc:'#107c10' -fill white -gravity center \
  -pointsize 72 -annotate +0+0 'SP' color.png

# outline.png — silueta monocroma transparente
convert -size 32x32 xc:none -fill white -gravity center \
  -pointsize 16 -annotate +0+0 'SP' outline.png
```

O usa el [Microsoft Teams App Icon Generator](https://developer.microsoft.com/microsoft-teams/app-icon-tool) oficial.

> **Honesto**: los iconos son un paso que la gente subestima. El `outline.png` **debe** ser monocromo transparente (blanco sobre transparente), no un color.png reducido — Teams lo muestra en temas oscuros y rompe si tiene fondo opaco.