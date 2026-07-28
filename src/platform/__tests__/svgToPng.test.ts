import { describe, expect, it } from 'vitest';
import { prepareSvgMarkup } from '../svgToPng';

describe('Office 图形位图化源', () => {
  it('为帕累托等含中文文字的静态 SVG 补齐尺寸和命名空间', () => {
    const source = prepareSvgMarkup('<svg viewBox="0 0 960 300"><text>缺陷帕累托图</text></svg>', 960, 300);

    expect(source).toContain('xmlns="http://www.w3.org/2000/svg"');
    expect(source).toContain('width="960"');
    expect(source).toContain('height="300"');
    expect(source).toContain('缺陷帕累托图');
  });
});
