#!/usr/bin/env bash
# Empaqueta la app de Teams en un .zip válido (3 archivos en la raíz, sin carpeta contenedora).
# Uso: bash package.sh
set -euo pipefail

APP_NAME="sp-projects-tab"
SRC_DIR="$(cd "$(dirname "$0")" && pwd)"
OUT="$SRC_DIR/$APP_NAME.zip"

# Validar que existan los 3 archivos requeridos
for f in manifest.json color.png outline.png; do
  if [ ! -f "$SRC_DIR/$f" ]; then
    echo "FALTA: $f — lee ICONS.md para crear los PNG." >&2
    exit 1
  fi
done

rm -f "$OUT"
( cd "$SRC_DIR" && zip -j -q "$OUT" manifest.json color.png outline.png )
echo "Empaquetado: $OUT"
echo "Sube el .zip en Teams → Apps → Manage your apps → Upload an app → Upload a custom app"