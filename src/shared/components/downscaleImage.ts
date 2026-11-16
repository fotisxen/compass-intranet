// Photos uploaded straight from a camera or phone are often 3000+ px wide,
// while a card shows them at a few hundred. Browsers shrink such a huge image
// in one jump, which looks rough, and it is also a slow multi-MB download to
// decode. Here the picture is shrunk in careful steps on a canvas to roughly
// the size it is shown at (twice that, for sharp high-DPI screens) and the
// result is used instead. If anything goes wrong the original URL is kept.

const cache: Record<string, string> = {};
const pending: Record<string, Promise<string>> = {};

/** The already-shrunk version of a picture, or the picture itself while it isn't ready. */
export function shrunkOrOriginal(url: string | undefined): string | undefined {
  return url ? cache[url] || url : url;
}

function load(url: string): Promise<HTMLImageElement> {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => resolve(img);
    img.onerror = () => reject(new Error('image failed'));
    img.src = url;
  });
}

function draw(source: CanvasImageSource, width: number, height: number): HTMLCanvasElement {
  const canvas = document.createElement('canvas');
  canvas.width = width;
  canvas.height = height;
  const ctx = canvas.getContext('2d');
  if (!ctx) {
    throw new Error('no canvas');
  }
  ctx.imageSmoothingEnabled = true;
  ctx.imageSmoothingQuality = 'high';
  ctx.drawImage(source, 0, 0, width, height);
  return canvas;
}

async function shrink(url: string, maxWidth: number): Promise<string> {
  const img = await load(url);
  // Already about the right size (or a vector/small image): nothing to gain.
  if (!img.naturalWidth || img.naturalWidth <= maxWidth * 1.25) {
    return url;
  }
  const targetW = maxWidth;
  const targetH = Math.max(1, Math.round(img.naturalHeight * (targetW / img.naturalWidth)));
  let current: CanvasImageSource = img;
  let w = img.naturalWidth;
  let h = img.naturalHeight;
  // Halve until within 2x of the target, then one last step to the exact size.
  while (w / 2 > targetW) {
    w = Math.round(w / 2);
    h = Math.round(h / 2);
    current = draw(current, w, h);
  }
  const finalCanvas = draw(current, targetW, targetH);
  const blob: Blob = await new Promise((resolve, reject) => {
    finalCanvas.toBlob(b => (b ? resolve(b) : reject(new Error('no blob'))), 'image/jpeg', 0.92);
  });
  return URL.createObjectURL(blob);
}

/** Resolves to a sharper, lighter copy of the picture (or the original URL if that isn't possible). Never rejects. */
export function downscaleImage(url: string | undefined, maxWidth: number): Promise<string | undefined> {
  if (!url) {
    return Promise.resolve(url);
  }
  const key = url;
  if (cache[key]) {
    return Promise.resolve(cache[key]);
  }
  if (!pending[key]) {
    pending[key] = shrink(url, maxWidth)
      .then(result => {
        cache[key] = result;
        return result;
      })
      .catch(() => {
        // Another site, blocked, or not decodable — the browser can still show the original.
        cache[key] = url;
        return url;
      });
  }
  return pending[key];
}
