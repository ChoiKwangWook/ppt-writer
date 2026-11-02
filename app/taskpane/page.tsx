import type { Metadata } from "next";

export const metadata: Metadata = { title: "PPT 작성기 작업창" };

// PLAN 작업 1: PowerPoint에 작업창이 뜨는지만 확인하는 빈 페이지. Office.js 연결과 화면은 작업 2·7에서 붙임
export default function TaskpanePage() {
  return (
    <main className="flex min-h-screen flex-col gap-2 bg-white p-4 text-zinc-900">
      <h1 className="text-lg font-bold">PPT 작성기</h1>
      <p className="text-sm text-zinc-600">작업창이 열렸어요. 채팅 기능은 다음 작업에서 연결해요.</p>
    </main>
  );
}
