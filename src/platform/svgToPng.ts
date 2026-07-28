/**
 * SVG → PNG 位图转换（浏览器 canvas 路径,Web 与 Tauri WebView 通用）。
 * 用于 Office 导出嵌图与图表 PNG 下载。
 */

export async function svgElementToPng(svg: SVGSVGElement, scale = 2): Promise<Uint8Array> {
  const vb = svg.viewBox.baseVal;
  const w = (vb && vb.width) || svg.clientWidth || 960;
  const h = (vb && vb.height) || svg.clientHeight || 300;
  const clone = svg.cloneNode(true) as SVGSVGElement;
  clone.setAttribute('width', String(w));
  clone.setAttribute('height', String(h));
  clone.setAttribute('xmlns', 'http://www.w3.org/2000/svg');
  const src = new XMLSerializer().serializeToString(clone);
  return svgMarkupToPng(src, w, h, scale);
}

export function prepareSvgMarkup(markup: string, w: number, h: number): string {
  let src = markup;
  if (!src.includes('xmlns=')) src = src.replace('<svg', '<svg xmlns="http://www.w3.org/2000/svg"');
  if (!/<svg[^>]*\swidth=/.test(src)) src = src.replace('<svg', `<svg width="${w}" height="${h}"`);
  return src;
}

function loadSvgImage(src: string): Promise<HTMLImageElement> {
  return new Promise<HTMLImageElement>((resolve, reject) => {
    const image = new Image();
    image.onload = () => resolve(image);
    image.onerror = () => reject(new Error('SVG 位图化失败'));
    image.src = src;
  });
}

export async function svgMarkupToPng(markup: string, w: number, h: number, scale = 2): Promise<Uint8Array> {
  const src = prepareSvgMarkup(markup, w, h);
  const url = URL.createObjectURL(new Blob([src], { type: 'image/svg+xml;charset=utf-8' }));
  try {
    // 部分 Windows WebView2 对 blob: SVG 的图像解码不稳定；失败时改用 data: SVG。
    // 两种来源都只含本地生成的 SVG，不依赖网络资源。
    let img: HTMLImageElement;
    try {
      img = await loadSvgImage(url);
    } catch {
      img = await loadSvgImage(`data:image/svg+xml;charset=utf-8,${encodeURIComponent(src)}`);
    }
    const canvas = document.createElement('canvas');
    canvas.width = Math.round(w * scale);
    canvas.height = Math.round(h * scale);
    const ctx = canvas.getContext('2d');
    if (!ctx) throw new Error('浏览器不支持图形报告画布');
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
    const blob = await new Promise<Blob>((resolve, reject) =>
      canvas.toBlob((b) => (b ? resolve(b) : reject(new Error('PNG 编码失败'))), 'image/png'),
    );
    return new Uint8Array(await blob.arrayBuffer());
  } finally {
    URL.revokeObjectURL(url);
  }
}
