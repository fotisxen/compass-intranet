/** Longest side, in pixels, a posted photo is scaled down to — keeps each photo around 150–300 KB. */
const MAX_SIDE = 1280;
const JPEG_QUALITY = 0.82;

export interface IResizedPhoto {
  /** For the preview thumbnail; revoke with URL.revokeObjectURL when removed. */
  previewUrl: string;
  fileName: string;
  data: ArrayBuffer;
}

function loadImage(url: string): Promise<HTMLImageElement> {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => resolve(img);
    img.onerror = () => reject(new Error('This file is not a picture the browser can read.'));
    img.src = url;
  });
}

/** Scales a picked photo down and re-encodes it as JPEG in the browser, before it is uploaded. */
export async function resizePhoto(file: File, index: number): Promise<IResizedPhoto> {
  const sourceUrl = URL.createObjectURL(file);
  try {
    const img = await loadImage(sourceUrl);
    const scale = Math.min(1, MAX_SIDE / Math.max(img.naturalWidth, img.naturalHeight));
    const width = Math.max(1, Math.round(img.naturalWidth * scale));
    const height = Math.max(1, Math.round(img.naturalHeight * scale));

    const canvas = document.createElement('canvas');
    canvas.width = width;
    canvas.height = height;
    const context = canvas.getContext('2d');
    if (!context) {
      throw new Error('This browser cannot process pictures.');
    }
    // White behind transparent PNGs, since JPEG has no transparency.
    context.fillStyle = '#ffffff';
    context.fillRect(0, 0, width, height);
    context.drawImage(img, 0, 0, width, height);

    const blob: Blob = await new Promise((resolve, reject) => {
      canvas.toBlob(b => (b ? resolve(b) : reject(new Error('Could not process the picture.'))), 'image/jpeg', JPEG_QUALITY);
    });
    const data = await new Promise<ArrayBuffer>((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result as ArrayBuffer);
      reader.onerror = () => reject(new Error('Could not read the picture.'));
      reader.readAsArrayBuffer(blob);
    });
    return { previewUrl: URL.createObjectURL(blob), fileName: `photo-${Date.now()}-${index}.jpg`, data };
  } finally {
    URL.revokeObjectURL(sourceUrl);
  }
}
