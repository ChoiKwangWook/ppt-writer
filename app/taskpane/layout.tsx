import OfficeGate from "./OfficeGate";

// 작업창은 한 화면만 쓰고 페이지 이동이 없어서, 개발 서버의 "즉시 화면 이동" 검사를 끔
export const instant = false;

// 작업창 페이지에서만 Office.js를 불러옴 (다른 페이지에는 넣지 않음)
export default function TaskpaneLayout({ children }: LayoutProps<"/taskpane">) {
  return <OfficeGate>{children}</OfficeGate>;
}
