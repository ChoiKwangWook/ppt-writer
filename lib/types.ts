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

// 모든 한도를 여기 한 곳에서만 정함 (DESIGN 4.4·6장). 프롬프트와 sanitize.ts가 함께 씀
export const LIMITS = {
  // PRD 5번 "AI가 지킬 규칙"의 개수 한도
  slides: 15,
  bullets: 5,
  diagramNodes: 8,
  // DESIGN 6장 보안: AI가 지나치게 긴 글을 돌려주지 못하게 하는 글자 수 한도
  titleChars: 120,
  bulletChars: 300,
  nodeLabelChars: 40,
} as const;
