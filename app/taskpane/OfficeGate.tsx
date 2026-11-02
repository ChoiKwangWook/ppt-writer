"use client";

import Script from "next/script";
import { useState } from "react";

// PowerPoint 2024에서 되는 기능 중 이 앱이 꼭 필요한 단계: 슬라이드 넣기·지우기(1.2), 표식(1.3)
const REQUIRED_API = { name: "PowerPointApi", version: "1.3" };

// Office.js가 PowerPoint 데스크톱에서 history.pushState·replaceState를 지워 Next.js 화면 이동이 깨지는 알려진 문제가 있어,
// Office.js를 불러오기 전에 보관했다가 불러온 뒤 되돌림
const savedHistory =
  typeof window !== "undefined"
    ? { pushState: window.history.pushState, replaceState: window.history.replaceState }
    : null;

type Status = "loading" | "outside" | "unsupported" | "error" | "ready";

const MESSAGES: Record<Exclude<Status, "ready">, string> = {
  loading: "PowerPoint와 연결하고 있어요…",
  outside: "PowerPoint에서 [PPT 작성기] 추가 기능으로 열어 주세요.",
  unsupported: "PowerPoint 2024 또는 Microsoft 365가 필요해요. 지금 버전에서는 PPT 작성기를 쓸 수 없어요.",
  error: "PowerPoint 연결 도구(Office.js)를 불러오지 못했어요. 인터넷 연결을 확인한 뒤 작업창을 다시 열어 주세요.",
};

export default function OfficeGate({ children }: { children: React.ReactNode }) {
  const [status, setStatus] = useState<Status>("loading");

  async function handleLoad() {
    if (savedHistory) {
      window.history.pushState ??= savedHistory.pushState;
      window.history.replaceState ??= savedHistory.replaceState;
    }
    const info = await Office.onReady();
    if (info.host !== Office.HostType.PowerPoint) {
      setStatus("outside");
    } else if (!Office.context.requirements.isSetSupported(REQUIRED_API.name, REQUIRED_API.version)) {
      setStatus("unsupported");
    } else {
      setStatus("ready");
    }
  }

  return (
    <>
      <Script
        src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"
        onLoad={handleLoad}
        onError={() => setStatus("error")}
      />
      {status === "ready" ? (
        children
      ) : (
        <main className="flex min-h-screen flex-col gap-2 bg-white p-4 text-zinc-900">
          <h1 className="text-lg font-bold">PPT 작성기</h1>
          <p
            role={status === "loading" ? "status" : "alert"}
            className={`text-sm ${status === "loading" ? "text-zinc-600" : "rounded-md bg-amber-50 p-3 text-amber-900"}`}
          >
            {MESSAGES[status]}
          </p>
        </main>
      )}
    </>
  );
}
