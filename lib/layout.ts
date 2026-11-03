import type { Box, Template } from "@/lib/templates";
import type { Slide } from "@/lib/types";

// 글상자 안쪽 여백 (인치). 미리보기와 PPT 파일이 같은 값을 씀
export const TEXT_INSET = 0.1;
export const LINE_HEIGHT = 1.25;
export const PARA_GAP = 0.3; // 글머리표 사이 간격 (글자 크기의 배수)
export const BULLET_INDENT = 1.5; // 글머리표 들여쓰기 (글자 크기의 배수)
const MIN_SIZE = 10;

// 한글·전각 글자는 1em, 영문·숫자는 약 0.55em으로 어림잡은 글자 폭
function textWidthEm(text: string): number {
  let em = 0;
  for (const ch of text) em += ch.charCodeAt(0) > 0x2e80 ? 1 : 0.55;
  return em;
}

/**
 * 문단들이 상자 안에 들어가는 가장 큰 글자 크기(pt)를 찾음.
 * 템플릿을 바꿔 상자가 좁아져도 글자 내용은 그대로 두고 크기만 줄이기 위함 (PRD must ②).
 */
export function fitFontSize(
  paragraphs: string[],
  box: Box,
  baseSize: number,
  { indent = false, minSize = MIN_SIZE, keepWords = false } = {},
): number {
  const availW = (box.w - TEXT_INSET * 2) * 72;
  const availH = (box.h - TEXT_INSET * 2) * 72;
  for (let size = baseSize; size > minSize; size--) {
    const lineW = availW - (indent ? size * BULLET_INDENT : 0);
    // keepWords: 짧은 글자("1단계")가 단어 한가운데서 줄이 바뀌지 않도록, 낱말이 한 줄에 들어갈 때까지 크기를 더 줄임
    if (keepWords && paragraphs.some((p) => p.split(/\s+/).some((w) => textWidthEm(w) * size > lineW))) continue;
    let height = 0;
    for (const p of paragraphs) {
      const lines = Math.max(1, Math.ceil((textWidthEm(p) * size) / lineW));
      height += lines * size * LINE_HEIGHT;
    }
    height += Math.max(0, paragraphs.length - 1) * size * PARA_GAP;
    if (height <= availH) return size;
  }
  return minSize;
}

export type NodeLayout = Box & { label: string; shape: "box" | "circle"; color: string; size: number };
export type ArrowLayout = { x: number; y: number; w: number };

export type SlideLayout = {
  title: { box: Box; size: number; color: string; bold: boolean };
  body: { box: Box; size: number; color: string; bold: boolean };
  nodes: NodeLayout[];
  arrows: ArrowLayout[];
};

const ARROW_GAP = 0.4;

// 원 안에 꼭 맞는 정사각형 (한 변이 지름의 1/√2)
const inscribed = (b: Box): Box => {
  const w = b.w * Math.SQRT1_2;
  const h = b.h * Math.SQRT1_2;
  return { x: b.x + (b.w - w) / 2, y: b.y + (b.h - h) / 2, w, h };
};

export function computeSlideLayout(slide: Slide, t: Template): SlideLayout {
  const nodesIn = slide.diagram?.nodes ?? [];
  const hasDiagram = nodesIn.length > 0;
  const bodyBox = hasDiagram ? t.bodyWithDiagram : t.body;

  const nodes: NodeLayout[] = [];
  const arrows: ArrowLayout[] = [];
  if (hasDiagram) {
    const d = t.diagram;
    const n = nodesIn.length;
    const slotW = (d.w - ARROW_GAP * (n - 1)) / n;
    nodesIn.forEach((node, i) => {
      const slotX = d.x + i * (slotW + ARROW_GAP);
      const w = node.shape === "circle" ? Math.min(slotW, d.h) : slotW;
      const h = node.shape === "circle" ? w : Math.min(d.h, 1.1);
      const box = { x: slotX + (slotW - w) / 2, y: d.y + (d.h - h) / 2, w, h };
      nodes.push({
        ...box,
        label: node.label,
        shape: node.shape,
        color: node.color ?? t.accent,
        // 원은 글자가 들어가는 자리가 바깥 틀에 꼭 맞는 정사각형(약 70%)이라서 그 크기로 글자 크기를 맞춤
        size: fitFontSize([node.label], node.shape === "circle" ? inscribed(box) : box, 14, { minSize: 8, keepWords: true }),
      });
      if (i < n - 1) {
        arrows.push({ x: slotX + slotW + 0.05, y: d.y + d.h / 2, w: ARROW_GAP - 0.1 });
      }
    });
  }

  return {
    title: {
      box: t.title,
      size: fitFontSize([slide.title], t.title, t.titleSize),
      color: slide.style?.titleColor ?? t.titleColor,
      bold: slide.style?.titleBold ?? true,
    },
    body: {
      box: bodyBox,
      size: fitFontSize(slide.bullets, bodyBox, t.bodySize, { indent: true }),
      color: slide.style?.bodyColor ?? t.bodyColor,
      bold: slide.style?.bodyBold ?? false,
    },
    nodes,
    arrows,
  };
}
