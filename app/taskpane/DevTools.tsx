"use client";

import { useEffect, useState } from "react";
import {
  readOverview,
  removeAddinSlides,
  renderDeck,
  syncFromPresentation,
  type Overview,
  type RenderOutcome,
  type SyncResult,
} from "@/lib/office";
import { TEMPLATES } from "@/lib/templates";
import type { Deck } from "@/lib/types";

// PLAN 작업 4·5 확인용 임시 패널 (개발용). 채팅 화면이 생기는 작업 7에서 없앰
const SAMPLE: Deck = {
  slides: [
    { title: "분기 업무 보고", bullets: ["목표 및 성과", "주요 지표 분석", "문제점 및 개선안", "향후 계획"] },
    {
      title: "추진 절차",
      bullets: ["단계별로 진행합니다"],
      diagram: {
        nodes: [
          { label: "기획", shape: "box" },
          { label: "개발", shape: "circle" },
          { label: "배포", shape: "box" },
        ],
      },
    },
    { title: "마무리", bullets: ["질문과 의견을 받습니다", "다음 일정을 안내합니다"], style: { titleColor: "#C00000" } },
  ],
};

// 작업이 걸린 시간(초)을 함께 돌려줌 (SC-1 측정에 PowerPoint 반영 시간을 따로 기록하기 위함)
async function timed(job: () => Promise<RenderOutcome>): Promise<{ result: RenderOutcome; seconds: string }> {
  const start = performance.now();
  const result = await job();
  return { result, seconds: ((performance.now() - start) / 1000).toFixed(1) };
}

const join = (numbers: number[]) => (numbers.length ? numbers.join("·") + "번" : "없음");

export default function DevTools() {
  const [overview, setOverview] = useState<Overview | null>(null);
  const [sync, setSync] = useState<SyncResult | null>(null);
  const [message, setMessage] = useState("");
  const [busy, setBusy] = useState(false);

  async function refresh() {
    try {
      const [o, s] = await Promise.all([readOverview(), syncFromPresentation()]);
      setOverview(o);
      setSync(s);
    } catch {
      setMessage("슬라이드 현황을 읽지 못했어요.");
    }
  }

  useEffect(() => {
    // 작업창을 열 때 한 번 현황과 저장된 데이터를 읽어 옴
    Promise.all([readOverview(), syncFromPresentation()])
      .then(([o, s]) => {
        setOverview(o);
        setSync(s);
      })
      .catch(() => setMessage("슬라이드 현황을 읽지 못했어요."));
  }, []);

  async function run(label: string, job: () => Promise<RenderOutcome>) {
    setBusy(true);
    setMessage("");
    try {
      const { result: r, seconds } = await timed(job);
      setMessage(
        `${label}: ${r.inserted}장 넣음 (${seconds}초)` +
          (r.leftoverOld ? ` · 예전 슬라이드 ${r.leftoverOld}장이 남았어요` : "") +
          (r.saveFailed ? " · 파일 안 저장에 실패했어요(PowerPoint에서 Ctrl+S를 눌러 주세요)" : " · 파일 안에 저장함"),
      );
    } catch (e) {
      setMessage((e as Error).message);
    } finally {
      await refresh();
      setBusy(false);
    }
  }

  const addinCount = overview?.slides.filter((s) => s.tag).length ?? 0;
  const templateName = TEMPLATES.find((t) => t.id === sync?.templateId)?.name;

  return (
    <section className="flex flex-col gap-2 border-t border-zinc-200 bg-zinc-50 p-4 text-xs text-zinc-800">
      <h2 className="text-sm font-bold">개발용 확인 (작업 4·5)</h2>

      <div className="flex flex-col gap-1.5">
        {TEMPLATES.map((t) => (
          <button
            key={t.id}
            disabled={busy}
            onClick={() => run(`${t.name}으로 넣기`, () => renderDeck(SAMPLE, t))}
            className="rounded-md bg-zinc-900 px-3 py-2 font-semibold text-white disabled:opacity-40"
          >
            샘플 3장을 {t.name}으로 넣기
          </button>
        ))}
        <button
          disabled={busy || addinCount === 0}
          onClick={() => run("지우기", removeAddinSlides)}
          className="rounded-md border border-zinc-400 bg-white px-3 py-2 font-semibold disabled:opacity-40"
        >
          추가 기능 슬라이드만 모두 지우기
        </button>
        <button disabled={busy} onClick={refresh} className="rounded-md border border-zinc-300 bg-white px-3 py-1.5 disabled:opacity-40">
          현황 새로 읽기
        </button>
      </div>

      {message && (
        <p role="status" className="rounded-md bg-emerald-50 p-2 text-emerald-800">
          {message}
        </p>
      )}

      {overview && (
        <div>
          <p className="font-semibold">
            슬라이드 {overview.slides.length}장 (추가 기능 {addinCount}장) · 슬라이드 마스터 {overview.masters}개
          </p>
          <ul className="mt-1 space-y-0.5">
            {overview.slides.map((s) => (
              <li key={s.number}>
                {s.number}번 · {s.tag ? `추가 기능 (${s.tag})` : "사람이 만든 슬라이드"}
              </li>
            ))}
          </ul>
        </div>
      )}

      {sync && (
        <div className="rounded-md border border-zinc-300 bg-white p-2">
          <p className="font-semibold">파일 안에 저장된 데이터와 화면 맞추기</p>
          {sync.templateId ? (
            <ul className="mt-1 space-y-0.5">
              <li>저장된 템플릿: {templateName}</li>
              <li>
                데이터의 슬라이드: {sync.deck.slides.length}장 → 화면 {join(sync.positions)}
              </li>
              <li>사람이 만든 슬라이드: 화면 {join(sync.humanSlides)}</li>
              <li>제목: {sync.deck.slides.map((s) => s.title).join(" / ") || "없음"}</li>
              <li className={sync.changed ? "font-semibold text-amber-700" : ""}>
                {sync.changed
                  ? "PowerPoint에서 지우거나 옮긴 것을 반영해 데이터를 맞췄어요"
                  : "저장된 데이터와 화면이 같아요"}
              </li>
            </ul>
          ) : (
            <p className="mt-1">저장된 데이터가 없어요.</p>
          )}
        </div>
      )}
    </section>
  );
}
