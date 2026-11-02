import { LIMITS, type Deck, type DiagramNode, type Slide, type SlideStyle } from "@/lib/types";

// AI나 브라우저가 보낸 슬라이드 데이터를 정해진 모양과 한도로 다듬음.
// 한도를 넘어 잘라낸 경우 사용자에게 알릴 문장을 notices에 담음 (PRD 예외 상황: 한도 초과)

const HEX = /^#?([0-9a-fA-F]{6})$/;

function color(value: unknown): string | undefined {
  const m = typeof value === "string" ? HEX.exec(value.trim()) : null;
  return m ? `#${m[1].toUpperCase()}` : undefined;
}

function text(value: unknown, max: number): string {
  return typeof value === "string" ? value.trim().slice(0, max) : "";
}

function style(raw: unknown): SlideStyle | undefined {
  if (!raw || typeof raw !== "object") return undefined;
  const r = raw as Record<string, unknown>;
  const s: SlideStyle = {};
  const titleColor = color(r.titleColor);
  const bodyColor = color(r.bodyColor);
  if (titleColor) s.titleColor = titleColor;
  if (bodyColor) s.bodyColor = bodyColor;
  if (typeof r.titleBold === "boolean") s.titleBold = r.titleBold;
  if (typeof r.bodyBold === "boolean") s.bodyBold = r.bodyBold;
  return Object.keys(s).length ? s : undefined;
}

export function sanitizeDeck(raw: unknown): { deck: Deck; notices: string[] } {
  const notices = new Set<string>();
  const slidesIn = Array.isArray((raw as Deck | null)?.slides) ? (raw as Deck).slides : [];
  if (slidesIn.length > LIMITS.slides) {
    notices.add(`슬라이드는 최대 ${LIMITS.slides}장까지 만들 수 있어서 ${LIMITS.slides}장까지만 담았어요.`);
  }

  const slides: Slide[] = slidesIn.slice(0, LIMITS.slides).map((s) => {
    const r = (s ?? {}) as Record<string, unknown>;
    const bulletsIn = Array.isArray(r.bullets) ? r.bullets : [];
    if (bulletsIn.length > LIMITS.bullets) {
      notices.add(`슬라이드 한 장의 글머리표는 최대 ${LIMITS.bullets}개라서 넘치는 항목은 뺐어요.`);
    }
    const slide: Slide = {
      title: text(r.title, 120),
      bullets: bulletsIn.slice(0, LIMITS.bullets).map((b) => text(b, 300)).filter(Boolean),
    };
    const st = style(r.style);
    if (st) slide.style = st;

    const nodesIn = Array.isArray((r.diagram as { nodes?: unknown } | undefined)?.nodes)
      ? ((r.diagram as { nodes: unknown[] }).nodes)
      : [];
    if (nodesIn.length > LIMITS.diagramNodes) {
      notices.add(`다이어그램 하나의 도형은 최대 ${LIMITS.diagramNodes}개까지라서 ${LIMITS.diagramNodes}개까지만 넣었어요.`);
    }
    const nodes: DiagramNode[] = nodesIn.slice(0, LIMITS.diagramNodes).map((n) => {
      const nr = (n ?? {}) as Record<string, unknown>;
      const node: DiagramNode = { label: text(nr.label, 40), shape: nr.shape === "circle" ? "circle" : "box" };
      const c = color(nr.color);
      if (c) node.color = c;
      return node;
    });
    if (nodes.length) slide.diagram = { nodes };
    return slide;
  });

  return { deck: { slides }, notices: [...notices] };
}
