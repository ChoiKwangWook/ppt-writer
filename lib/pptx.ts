import { computeSlideLayout, LINE_HEIGHT, PARA_GAP, TEXT_INSET } from "@/lib/layout";
import { SLIDE_W, type Box, type Template } from "@/lib/templates";
import type { Deck } from "@/lib/types";

// 미리보기와 같은 배치 계산(computeSlideLayout)으로 .pptx를 만듦.
// PowerPoint에 넣을 때는 base64 글자(buildPptxBase64)로, 예전 웹 화면은 파일 내려받기(downloadPptx)로 씀

// PowerPoint 기본 16:9 슬라이드 크기(인치). 템플릿 좌표는 10×5.625인치 기준이라
// 위치·크기·글자 크기를 모두 같은 비율(S배)로 키워 모양이 달라지지 않게 함
const PPT_W = 12192000 / 914400; // PowerPoint 기본 가로 크기(EMU)와 정확히 같게 맞춤
const PPT_H = 7.5;
const S = PPT_W / SLIDE_W;

const hex = (c: string) => c.replace("#", "");
const MARGIN_PT = TEXT_INSET * 72 * S;
const scaleBox = (b: Box): Box => ({ x: b.x * S, y: b.y * S, w: b.w * S, h: b.h * S });
const scaleSize = (pt: number) => Math.round(pt * S * 10) / 10;

async function buildPresentation(deck: Deck, t: Template) {
  const { default: PptxGenJS } = await import("pptxgenjs");
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: "PPT_WRITER_16x9", width: PPT_W, height: PPT_H });
  pptx.layout = "PPT_WRITER_16x9";

  for (const slide of deck.slides) {
    const L = computeSlideLayout(slide, t);
    const s = pptx.addSlide();
    s.background = { color: hex(t.background) };

    for (const d of t.decorations) {
      s.addShape(pptx.ShapeType.rect, {
        ...scaleBox(d),
        fill: { color: hex(d.color) },
        line: { color: hex(d.color), width: 0 },
      });
    }

    s.addText(slide.title, {
      ...scaleBox(L.title.box),
      fontFace: t.fontFace,
      fontSize: scaleSize(L.title.size),
      bold: L.title.bold,
      color: hex(L.title.color),
      align: t.titleAlign,
      valign: "middle",
      margin: MARGIN_PT,
    });

    if (slide.bullets.length) {
      s.addText(
        slide.bullets.map((b) => ({ text: b, options: { bullet: true, breakLine: true } })),
        {
          ...scaleBox(L.body.box),
          fontFace: t.fontFace,
          fontSize: scaleSize(L.body.size),
          bold: L.body.bold,
          color: hex(L.body.color),
          valign: "top",
          margin: MARGIN_PT,
          lineSpacingMultiple: LINE_HEIGHT / 1.2,
          paraSpaceAfter: Math.round(L.body.size * S * PARA_GAP),
        },
      );
    }

    for (const n of L.nodes) {
      s.addText(n.label, {
        shape: n.shape === "circle" ? pptx.ShapeType.ellipse : pptx.ShapeType.roundRect,
        ...scaleBox(n),
        rectRadius: 0.15 * S,
        fill: { color: hex(n.color) },
        line: { color: hex(n.color), width: 0 },
        fontFace: t.fontFace,
        fontSize: scaleSize(n.size),
        bold: true,
        color: "FFFFFF",
        align: "center",
        valign: "middle",
        margin: 4 * S,
      });
    }
    for (const a of L.arrows) {
      s.addShape(pptx.ShapeType.line, {
        x: a.x * S, y: a.y * S, w: a.w * S, h: 0,
        line: { color: hex(L.body.color), width: 2 * S, endArrowType: "triangle" },
      });
    }
  }

  return pptx;
}

// PowerPoint에 넣을 .pptx 내용을 base64 글자로 돌려줌 (insertSlidesFromBase64에 그대로 넘김)
export async function buildPptxBase64(deck: Deck, t: Template): Promise<string> {
  const pptx = await buildPresentation(deck, t);
  return (await pptx.write({ outputType: "base64" })) as string;
}

// 예전 웹 화면의 파일 내려받기. 웹 화면 코드와 함께 PLAN 작업 13에서 정리함
export async function downloadPptx(deck: Deck, t: Template) {
  const pptx = await buildPresentation(deck, t);
  const name = (deck.slides[0]?.title || "발표자료").replace(/[\\/:*?"<>|]/g, "").slice(0, 50);
  await pptx.writeFile({ fileName: `${name}.pptx` });
}
