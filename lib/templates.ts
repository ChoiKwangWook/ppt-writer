import type { TemplateId } from "@/lib/types";

// 모든 좌표는 인치 단위 (16:9 슬라이드 = 가로 10 x 세로 5.625)
export const SLIDE_W = 10;
export const SLIDE_H = 5.625;

export type Box = { x: number; y: number; w: number; h: number };

export type Template = {
  id: TemplateId;
  name: string;
  description: string;
  fontFace: string;
  background: string;
  titleColor: string;
  bodyColor: string;
  accent: string;
  titleSize: number;
  bodySize: number;
  titleAlign: "left" | "center";
  title: Box;
  body: Box;
  // 다이어그램이 있으면 본문은 위쪽, 다이어그램은 아래쪽에 놓음
  bodyWithDiagram: Box;
  diagram: Box;
  decorations: (Box & { color: string })[];
};

export const TEMPLATES: Template[] = [
  {
    id: "report",
    name: "보고서형",
    description: "흰 배경에 남색 제목, 차분한 보고 자료",
    fontFace: "Malgun Gothic",
    background: "#FFFFFF",
    titleColor: "#1F3864",
    bodyColor: "#333333",
    accent: "#2E75B6",
    titleSize: 28,
    bodySize: 18,
    titleAlign: "left",
    title: { x: 0.5, y: 0.35, w: 9, h: 0.9 },
    body: { x: 0.5, y: 1.4, w: 9, h: 3.8 },
    bodyWithDiagram: { x: 0.5, y: 1.4, w: 9, h: 1.9 },
    diagram: { x: 0.5, y: 3.5, w: 9, h: 1.6 },
    decorations: [
      { x: 0, y: 0, w: 0.15, h: SLIDE_H, color: "#2E75B6" },
      { x: 0.5, y: 1.25, w: 9, h: 0.03, color: "#BDD7EE" },
    ],
  },
  {
    id: "meeting",
    name: "회의형",
    description: "색 띠 제목과 큰 글씨, 회의에서 보기 쉬운 자료",
    fontFace: "Malgun Gothic",
    background: "#F4F6F8",
    titleColor: "#FFFFFF",
    bodyColor: "#263238",
    accent: "#00796B",
    titleSize: 26,
    bodySize: 20,
    titleAlign: "left",
    title: { x: 0.4, y: 0.1, w: 9.2, h: 0.9 },
    body: { x: 0.6, y: 1.35, w: 8.8, h: 3.8 },
    bodyWithDiagram: { x: 0.6, y: 1.35, w: 8.8, h: 1.9 },
    diagram: { x: 0.6, y: 3.45, w: 8.8, h: 1.7 },
    decorations: [
      { x: 0, y: 0, w: SLIDE_W, h: 1.1, color: "#00796B" },
      { x: 0, y: 5.425, w: SLIDE_W, h: 0.2, color: "#80CBC4" },
    ],
  },
  {
    id: "education",
    name: "교육형",
    description: "어두운 배경에 가운데 큰 제목, 강의·교육 자료",
    fontFace: "Malgun Gothic",
    background: "#1E1E2E",
    titleColor: "#FFD166",
    bodyColor: "#F1F1F1",
    accent: "#EF8354",
    titleSize: 32,
    bodySize: 20,
    titleAlign: "center",
    title: { x: 0.5, y: 0.3, w: 9, h: 1.0 },
    body: { x: 1, y: 1.6, w: 8, h: 3.6 },
    bodyWithDiagram: { x: 1, y: 1.6, w: 8, h: 1.8 },
    diagram: { x: 1, y: 3.6, w: 8, h: 1.6 },
    decorations: [{ x: 4, y: 1.35, w: 2, h: 0.06, color: "#FFD166" }],
  },
];

export function getTemplate(id: TemplateId): Template {
  return TEMPLATES.find((t) => t.id === id) ?? TEMPLATES[0];
}
