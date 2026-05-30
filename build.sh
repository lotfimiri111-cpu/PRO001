#!/bin/bash
# مذكرتي Pro v28 — Build Script
# No set -e: optional steps shouldn't abort the build

echo "======================================"
echo "  مذكرتي Pro v28 — Build"
echo "  Python: $(python3 --version)"
echo "======================================"

echo ""
echo "==> [1/5] System packages..."
apt-get update -qq 2>/dev/null && \
  apt-get install -y -qq \
    fontconfig fonts-noto-core \
    libreoffice-impress libreoffice-calc libreoffice-writer \
    poppler-utils curl \
    2>/dev/null && echo "    ✅ System packages installed" || \
  echo "    ⚠️  apt-get failed (normal on some platforms)"

echo ""
echo "==> [2/5] Arabic font (Cairo)..."
FONT_DIR="${HOME}/.fonts/cairo"
mkdir -p "$FONT_DIR" 2>/dev/null || { FONT_DIR="/tmp/fonts/cairo"; mkdir -p "$FONT_DIR"; }

if fc-list 2>/dev/null | grep -qi "cairo"; then
  echo "    ✅ Cairo already present"
else
  FONT_URL="https://github.com/google/fonts/raw/main/ofl/cairo/Cairo%5Bslnt%2Cwght%5D.ttf"
  if curl -fsSL --max-time 30 "$FONT_URL" -o "$FONT_DIR/Cairo.ttf" 2>/dev/null; then
    fc-cache -fv "$FONT_DIR" 2>/dev/null || true
    echo "    ✅ Cairo downloaded"
  else
    # Fallback: Amiri
    curl -fsSL --max-time 30 \
      "https://github.com/google/fonts/raw/main/ofl/amiri/Amiri-Regular.ttf" \
      -o "$FONT_DIR/Amiri-Regular.ttf" 2>/dev/null && \
      fc-cache -fv "$FONT_DIR" 2>/dev/null && \
      echo "    ✅ Amiri fallback installed" || \
      echo "    ⚠️  No Arabic font — Calibri fallback"
  fi
fi

echo ""
echo "==> [3/5] LibreOffice check..."
if command -v soffice &>/dev/null; then
  echo "    ✅ soffice: $(soffice --version 2>/dev/null || echo 'found')"
else
  echo "    ⚠️  soffice not found — preview uses Pillow fallback"
fi

if command -v pdftoppm &>/dev/null; then
  echo "    ✅ pdftoppm found"
else
  echo "    ⚠️  pdftoppm not found"
fi

echo ""
echo "==> [4/5] Python dependencies..."
pip install --no-cache-dir -r requirements.txt

echo ""
echo "==> [5/5] Smoke test..."
python3 -c "
from engine.pipeline import get_pipeline
from core.models import PresentationRequest
pipeline = get_pipeline()
req = PresentationRequest.from_dict({'titleAr':'test','studentName':'test'})
r = pipeline.build(req)
assert r.success, f'Build failed: {r.error}'
print(f'    ✅ Pipeline OK — {r.slide_count} slides in {r.elapsed:.2f}s — font={pipeline._font}')
" || echo "    ❌ Smoke test failed"

echo ""
echo "======================================"
echo "  Build Complete — v28"
echo "  Python  : $(python3 --version)"
echo "  soffice : $(command -v soffice 2>/dev/null || echo 'not found')"
echo "  pdftoppm: $(command -v pdftoppm 2>/dev/null || echo 'not found')"
echo "  Cairo   : $(fc-list 2>/dev/null | grep -ic cairo) file(s)"
echo "======================================"
