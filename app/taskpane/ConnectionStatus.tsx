"use client";

import { useEffect, useState } from "react";

// PLAN 작업 2: PowerPoint와 실제로 연결됐는지 보여 주는 임시 화면. 작업 7에서 채팅 화면(PptChat)으로 바뀜
export default function ConnectionStatus() {
  const [slideCount, setSlideCount] = useState<number | null>(null);

  useEffect(() => {
    PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load("items");
      await context.sync();
      setSlideCount(slides.items.length);
    }).catch(() => setSlideCount(-1));
  }, []);

  return (
    <main className="flex flex-col gap-2 bg-white p-4 text-zinc-900">
      <h1 className="text-lg font-bold">PPT 작성기</h1>
      <p className="rounded-md bg-emerald-50 p-3 text-sm text-emerald-800">PowerPoint와 연결됐어요.</p>
      <p className="text-sm text-zinc-600">
        {slideCount === null
          ? "슬라이드 수를 확인하고 있어요…"
          : slideCount < 0
            ? "슬라이드 정보를 읽지 못했어요."
            : `지금 파일의 슬라이드: ${slideCount}장`}
      </p>
      <p className="text-xs text-zinc-500">채팅 기능은 다음 작업에서 연결해요.</p>
    </main>
  );
}
