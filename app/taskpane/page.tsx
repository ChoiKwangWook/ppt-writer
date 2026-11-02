import type { Metadata } from "next";
import ConnectionStatus from "./ConnectionStatus";

export const metadata: Metadata = { title: "PPT 작성기 작업창" };

// PowerPoint와 연결되기 전까지 OfficeGate가 이 페이지를 일부러 그리지 않으므로 "즉시 화면 이동" 검사를 끔
export const instant = false;

export default function TaskpanePage() {
  return <ConnectionStatus />;
}
