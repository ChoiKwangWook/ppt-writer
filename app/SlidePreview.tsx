import type { CSSProperties } from "react";
import { BULLET_INDENT, computeSlideLayout, LINE_HEIGHT, PARA_GAP, TEXT_INSET } from "@/lib/layout";
import { SLIDE_H, SLIDE_W, type Box, type Template } from "@/lib/templates";
import type { Slide } from "@/lib/types";

// 슬라이드 폭(10인치 = 720pt)을 기준으로 pt를 화면 크기에 맞춰 바꿈 (100cqw = 슬라이드 폭)
const pt = (n: number) => `calc(${n} * 100cqw / 720)`;
const inch = (n: number) => `calc(${n} * 100cqw / ${SLIDE_W})`;

function place(b: Box): CSSProperties {
  return {
    position: "absolute",
    left: `${(b.x / SLIDE_W) * 100}%`,
    top: `${(b.y / SLIDE_H) * 100}%`,
    width: `${(b.w / SLIDE_W) * 100}%`,
    height: `${(b.h / SLIDE_H) * 100}%`,
  };
}

export default function SlidePreview({ slide, template: t }: { slide: Slide; template: Template }) {
  const L = computeSlideLayout(slide, t);

  return (
    <div
      className="relative aspect-video w-full overflow-hidden rounded-md shadow-sm ring-1 ring-black/10"
      style={{ containerType: "inline-size", background: t.background, fontFamily: `"${t.fontFace}", "맑은 고딕", sans-serif` }}
    >
      {t.decorations.map((d, i) => (
        <div key={i} style={{ ...place(d), background: d.color }} />
      ))}

      <div
        style={{
          ...place(L.title.box),
          padding: inch(TEXT_INSET),
          display: "flex",
          alignItems: "center",
          justifyContent: t.titleAlign === "center" ? "center" : "flex-start",
          textAlign: t.titleAlign,
          fontSize: pt(L.title.size),
          fontWeight: L.title.bold ? 700 : 400,
          color: L.title.color,
          lineHeight: LINE_HEIGHT,
        }}
      >
        {slide.title}
      </div>

      {slide.bullets.length > 0 && (
        <ul
          style={{
            ...place(L.body.box),
            padding: inch(TEXT_INSET),
            paddingLeft: `calc(${inch(TEXT_INSET)} + ${BULLET_INDENT}em)`,
            listStyle: "disc",
            fontSize: pt(L.body.size),
            fontWeight: L.body.bold ? 700 : 400,
            color: L.body.color,
            lineHeight: LINE_HEIGHT,
          }}
        >
          {slide.bullets.map((b, i) => (
            <li key={i} style={{ marginBottom: `${PARA_GAP}em` }}>
              {b}
            </li>
          ))}
        </ul>
      )}

      {L.nodes.map((n, i) => (
        <div
          key={i}
          style={{
            ...place(n),
            background: n.color,
            borderRadius: n.shape === "circle" ? "50%" : inch(0.15),
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            textAlign: "center",
            padding: pt(4),
            fontSize: pt(n.size),
            fontWeight: 700,
            color: "#FFFFFF",
            lineHeight: 1.2,
          }}
        >
          {n.label}
        </div>
      ))}
      {L.arrows.map((a, i) => (
        <svg
          key={i}
          viewBox="0 0 100 20"
          preserveAspectRatio="none"
          style={{ ...place({ x: a.x, y: a.y - 0.1, w: a.w, h: 0.2 }), overflow: "visible" }}
        >
          <line x1="0" y1="10" x2="78" y2="10" stroke={L.body.color} strokeWidth="6" />
          <polygon points="70,0 100,10 70,20" fill={L.body.color} />
        </svg>
      ))}
    </div>
  );
}
