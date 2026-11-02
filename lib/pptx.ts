import { computeSlideLayout, LINE_HEIGHT, PARA_GAP, TEXT_INSET } from "@/lib/layout";
import type { Template } from "@/lib/templates";
import type { Deck } from "@/lib/types";

// 브라우저에서 .pptx 파일을 만들어 내려받음 (미리보기와 같은 배치 계산을 사용)

const hex = (c: string) => c.replace("#", "");
const MARGIN_PT = TEXT_INSET * 72;

export async function downloadPptx(deck: Deck, t: Template) {
  const { default: PptxGenJS } = await import("pptxgenjs");
  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_16x9";

  for (const slide of deck.slides) {
    const L = computeSlideLayout(slide, t);
    const s = pptx.addSlide();
    s.background = { color: hex(t.background) };

    for (const d of t.decorations) {
      s.addShape(pptx.ShapeType.rect, {
        x: d.x, y: d.y, w: d.w, h: d.h,
        fill: { color: hex(d.color) },
        line: { color: hex(d.color), width: 0 },
      });
    }

    s.addText(slide.title, {
      ...L.title.box,
      fontFace: t.fontFace,
      fontSize: L.title.size,
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
          ...L.body.box,
          fontFace: t.fontFace,
          fontSize: L.body.size,
          bold: L.body.bold,
          color: hex(L.body.color),
          valign: "top",
          margin: MARGIN_PT,
          lineSpacingMultiple: LINE_HEIGHT / 1.2,
          paraSpaceAfter: Math.round(L.body.size * PARA_GAP),
        },
      );
    }

    for (const n of L.nodes) {
      s.addText(n.label, {
        shape: n.shape === "circle" ? pptx.ShapeType.ellipse : pptx.ShapeType.roundRect,
        x: n.x, y: n.y, w: n.w, h: n.h,
        rectRadius: 0.15,
        fill: { color: hex(n.color) },
        line: { color: hex(n.color), width: 0 },
        fontFace: t.fontFace,
        fontSize: n.size,
        bold: true,
        color: "FFFFFF",
        align: "center",
        valign: "middle",
        margin: 4,
      });
    }
    for (const a of L.arrows) {
      s.addShape(pptx.ShapeType.line, {
        x: a.x, y: a.y, w: a.w, h: 0,
        line: { color: hex(L.body.color), width: 2, endArrowType: "triangle" },
      });
    }
  }

  const name = (deck.slides[0]?.title || "발표자료").replace(/[\\/:*?"<>|]/g, "").slice(0, 50);
  await pptx.writeFile({ fileName: `${name}.pptx` });
}
