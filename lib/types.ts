// 화면·서버·PPT 내보내기가 함께 쓰는 슬라이드 데이터 모양

export type DiagramNode = {
  label: string;
  shape: "box" | "circle";
  color?: string; // "#RRGGBB"
};

export type Diagram = { nodes: DiagramNode[] };

// 템플릿 기본값 위에 덮어쓰는 슬라이드별 서식 (공통 변경으로 바뀌는 부분)
export type SlideStyle = {
  titleColor?: string;
  titleBold?: boolean;
  bodyColor?: string;
  bodyBold?: boolean;
};

export type Slide = {
  title: string;
  bullets: string[];
  style?: SlideStyle;
  diagram?: Diagram;
};

export type Deck = { slides: Slide[] };

export type TemplateId = "report" | "meeting" | "education";

// PRD 5번 "AI가 지킬 규칙"의 한도
export const LIMITS = { slides: 15, bullets: 5, diagramNodes: 8 } as const;
