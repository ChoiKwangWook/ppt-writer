"use client";

import { useEffect, useRef, useState } from "react";
import { renderDeck, syncFromPresentation } from "@/lib/office";
import { getTemplate, TEMPLATES } from "@/lib/templates";
import type { Deck, TemplateId } from "@/lib/types";
import SlidePreview from "./SlidePreview";

// PowerPoint 오른쪽 작업창(폭 약 320px)에서 쓰는 화면: 템플릿 선택 → 채팅. 슬라이드 결과는 PowerPoint 화면에서 봄

type Message = { role: "user" | "assistant"; content: string; local?: boolean }; // local: 화면에만 보이는 안내 (AI에게 보내지 않음)
type Retry =
  | { kind: "chat"; message: string }
  | { kind: "render"; deck: Deck; templateId: TemplateId; nextPrev: Deck | null | undefined };

const TIMEOUT_MS = 30_000;
const FAIL_TEXT = "AI 응답을 받지 못했어요. 잠시 후 다시 시도해 주세요.";
const SYNC_FAIL = "PowerPoint와 연결하지 못했어요. 잠시 후 다시 시도해 주세요.";
const EXAMPLES = [
  "신입사원 오리엔테이션 자료 5장 만들어줘",
  "모든 제목을 파란색 굵은 글씨로 바꿔줘",
  "2번 슬라이드에 3단계 흐름도를 넣어줘",
];

export default function PptChat() {
  const [templateId, setTemplateId] = useState<TemplateId | null>(null);
  const [deck, setDeck] = useState<Deck>({ slides: [] });
  const [prevDeck, setPrevDeck] = useState<Deck | null>(null); // 되돌리기용 직전 내용 (1단계, 작업창을 닫으면 사라짐)
  const [messages, setMessages] = useState<Message[]>([]);
  const [input, setInput] = useState("");
  const [busy, setBusy] = useState<string | null>(null); // 하고 있는 일의 안내 문구, 쉬는 중이면 null
  const [error, setError] = useState<{ text: string; retry: Retry | null } | null>(null);
  const chatEnd = useRef<HTMLDivElement>(null);

  const template = templateId ? getTemplate(templateId) : null;

  useEffect(() => {
    chatEnd.current?.scrollIntoView({ behavior: "smooth" });
  }, [messages, busy, error]);

  // 작업창을 열 때, 파일 안에 저장된 추가 기능 슬라이드가 있으면 바로 채팅 단계로 이어서 시작함
  useEffect(() => {
    syncFromPresentation()
      .then((s) => {
        if (s.templateId && s.deck.slides.length > 0) {
          setDeck(s.deck);
          setTemplateId(s.templateId);
        }
      })
      .catch(() => {});
  }, []);

  const say = (content: string) => setMessages((m) => [...m, { role: "assistant", content, local: true }]);

  // 슬라이드 데이터를 PowerPoint에 반영하고, 성공했을 때만 화면 상태를 바꿈 (실패하면 이전 상태 그대로)
  async function apply(newDeck: Deck, id: TemplateId, nextPrev?: Deck | null): Promise<boolean> {
    try {
      const r = await renderDeck(newDeck, getTemplate(id));
      setDeck(newDeck);
      setTemplateId(id);
      if (nextPrev !== undefined) setPrevDeck(nextPrev);
      if (r.saveFailed) say("파일 안에 저장하지 못했어요. PowerPoint에서 Ctrl+S로 저장해 주세요.");
      if (r.leftoverOld) say("예전 슬라이드 일부를 지우지 못했어요. 다음 수정 때 정리할게요.");
      return true;
    } catch (e) {
      setError({ text: (e as Error).message, retry: { kind: "render", deck: newDeck, templateId: id, nextPrev } });
      return false;
    }
  }

  async function send(text: string, isRetry = false) {
    const message = text.trim();
    if (!message || busy || !template) return;

    // 다시 시도할 때는 이미 화면에 있는 마지막 요청을 다시 보냄
    const history = (isRetry ? messages.slice(0, -1) : messages).filter((m) => !m.local);
    if (!isRetry) {
      setMessages([...messages, { role: "user", content: message }]);
      setInput("");
    }
    setError(null);
    setBusy("슬라이드를 만들고 있어요…");

    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), TIMEOUT_MS);
    try {
      // 요청을 보내기 직전에 PowerPoint 화면과 맞춰서, 사용자가 지우거나 옮긴 슬라이드와 번호를 AI에게 알림
      const synced = await syncFromPresentation().catch(() => {
        throw new Error(SYNC_FAIL);
      });
      if (synced.changed) {
        setDeck(synced.deck);
        say("PowerPoint에서 지우거나 옮긴 추가 기능 슬라이드를 반영했어요.");
      }

      const res = await fetch("/api/chat", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          message,
          deck: synced.deck,
          history,
          positions: synced.positions,
          humanSlides: synced.humanSlides,
        }),
        signal: controller.signal,
      });
      const data = await res.json().catch(() => ({}));
      if (!res.ok) throw new Error(data.error || FAIL_TEXT);

      setMessages((m) => [...m, { role: "assistant", content: data.reply }]);
      if (data.deck) await apply(data.deck, template.id, synced.deck);
    } catch (e) {
      const aborted = e instanceof DOMException && e.name === "AbortError";
      // 인터넷이 끊기면 fetch가 TypeError("Failed to fetch" 등 영어 원문)를 던지므로 한국어 안내로 바꿈
      const offline = e instanceof TypeError;
      setError({
        text: aborted
          ? "30초 안에 응답이 오지 않았어요. 잠시 후 다시 시도해 주세요."
          : offline
            ? FAIL_TEXT
            : (e as Error).message,
        retry: { kind: "chat", message },
      });
    } finally {
      clearTimeout(timer);
      setBusy(null);
    }
  }

  async function retry() {
    const r = error?.retry;
    if (!r || busy) return;
    setError(null);
    if (r.kind === "chat") {
      await send(r.message, true);
      return;
    }
    setBusy("PowerPoint에 반영하는 중…");
    await apply(r.deck, r.templateId, r.nextPrev);
    setBusy(null);
  }

  // 템플릿 바꾸기: AI를 쓰지 않고, 같은 내용을 새 디자인으로 PowerPoint에 다시 넣기만 함
  async function switchTemplate(id: TemplateId) {
    if (busy || id === templateId) return;
    if (!deck.slides.length) {
      setTemplateId(id);
      return;
    }
    setError(null);
    setBusy("템플릿을 바꾸는 중…");
    try {
      const synced = await syncFromPresentation().catch(() => {
        throw new Error(SYNC_FAIL);
      });
      await apply(synced.deck, id);
    } catch (e) {
      setError({ text: (e as Error).message, retry: null });
    } finally {
      setBusy(null);
    }
  }

  // 되돌리기: 채팅으로 바뀐 내용을 직전 상태로 1단계 되돌림 (템플릿은 그대로)
  async function undo() {
    if (!prevDeck || busy || !template) return;
    setError(null);
    setBusy("되돌리는 중…");
    if (await apply(prevDeck, template.id, null)) say("직전 상태로 되돌렸어요.");
    setBusy(null);
  }

  // 템플릿을 아직 고르지 않았을 때
  if (!template) {
    return (
      <main className="flex min-h-screen flex-col gap-3 bg-white p-4 text-zinc-900">
        <h1 className="text-xl font-extrabold">PPT 작성기</h1>
        <p className="text-sm font-semibold">1. 시작할 템플릿을 골라 주세요</p>
        <div className="flex flex-col gap-3">
          {TEMPLATES.map((t) => (
            <button
              key={t.id}
              onClick={() => setTemplateId(t.id)}
              className="rounded-xl p-2 text-left shadow-sm ring-1 ring-zinc-200 transition hover:ring-2 hover:ring-zinc-900"
            >
              <SlidePreview
                template={t}
                slide={{ title: `${t.name} 템플릿`, bullets: ["채팅으로 내용을 입력하면", "슬라이드가 만들어져요"] }}
              />
              <div className="mt-2 text-sm font-semibold">{t.name}</div>
              <div className="text-xs text-zinc-600">{t.description}</div>
            </button>
          ))}
        </div>
      </main>
    );
  }

  return (
    <main className="flex h-screen flex-col bg-white text-zinc-900">
      <header className="flex flex-col gap-1.5 border-b border-zinc-200 p-2.5">
        <div className="flex flex-wrap items-center gap-1">
          <span className="mr-1 text-xs text-zinc-500">템플릿</span>
          {TEMPLATES.map((t) => (
            <button
              key={t.id}
              onClick={() => switchTemplate(t.id)}
              disabled={!!busy}
              className={`flex items-center gap-1 rounded-full border px-2 py-1 text-xs transition disabled:opacity-50 ${
                templateId === t.id ? "border-zinc-900 bg-zinc-900 text-white" : "border-zinc-300 bg-white hover:bg-zinc-50"
              }`}
            >
              <span className="h-2.5 w-2.5 rounded-full ring-1 ring-black/20" style={{ background: t.accent }} />
              {t.name}
            </button>
          ))}
        </div>
        <div className="flex items-center justify-between text-xs text-zinc-600">
          <span>추가 기능 슬라이드 {deck.slides.length}장</span>
          <button
            onClick={undo}
            disabled={!prevDeck || !!busy}
            className="rounded-md border border-zinc-300 px-2 py-1 hover:bg-zinc-50 disabled:opacity-40"
          >
            되돌리기
          </button>
        </div>
      </header>

      <div className="flex-1 space-y-2 overflow-y-auto p-3" aria-live="polite">
        {messages.length === 0 && (
          <div className="space-y-2">
            <p className="text-xs text-zinc-600">만들고 싶은 발표 자료의 주제와 내용을 알려 주세요. 예를 들면:</p>
            {EXAMPLES.map((ex) => (
              <button
                key={ex}
                onClick={() => setInput(ex)}
                className="block w-full rounded-lg border border-zinc-200 px-3 py-2 text-left text-xs hover:bg-zinc-50"
              >
                {ex}
              </button>
            ))}
          </div>
        )}
        {messages.map((m, i) => (
          <div key={i} className={`flex ${m.role === "user" ? "justify-end" : "justify-start"}`}>
            <div
              className={`max-w-[90%] whitespace-pre-wrap rounded-2xl px-3 py-2 text-sm ${
                m.role === "user" ? "bg-zinc-900 text-white" : "bg-zinc-100 text-zinc-900"
              }`}
            >
              {m.content}
            </div>
          </div>
        ))}
        {busy && <div className="text-xs text-zinc-500">{busy}</div>}
        {error && (
          <div className="rounded-lg border border-red-200 bg-red-50 p-2.5 text-xs text-red-800">
            {error.text}
            {error.retry && (
              <button
                onClick={retry}
                className="ml-2 rounded-md bg-red-700 px-2 py-1 text-xs font-semibold text-white hover:bg-red-800"
              >
                다시 시도
              </button>
            )}
          </div>
        )}
        <div ref={chatEnd} />
      </div>

      <form
        className="border-t border-zinc-200 p-2.5"
        onSubmit={(e) => {
          e.preventDefault();
          send(input);
        }}
      >
        <p className="mb-1.5 rounded-md bg-zinc-100 px-2 py-1.5 text-[11px] leading-4 text-zinc-700">
          ⓘ 추가 기능이 만든 슬라이드만 고쳐요. 손으로 고친 내용은 다음 수정 때 바뀔 수 있어요. 저장은 PowerPoint에서 Ctrl+S로 해 주세요.
        </p>
        <p className="mb-2 rounded-md bg-amber-50 px-2 py-1.5 text-[11px] text-amber-800">
          ⚠ 대외비·개인정보는 입력하지 마세요.
        </p>
        <div className="flex gap-1.5">
          <textarea
            value={input}
            onChange={(e) => setInput(e.target.value)}
            onKeyDown={(e) => {
              // 한글 조합 중 Enter는 무시해야 같은 글이 두 번 전송되지 않음
              if (e.key === "Enter" && !e.shiftKey && !e.nativeEvent.isComposing) {
                e.preventDefault();
                send(input);
              }
            }}
            rows={2}
            maxLength={2000}
            placeholder="예: 모든 제목을 파란색으로 바꿔줘"
            className="min-w-0 flex-1 resize-none rounded-lg border border-zinc-300 px-2 py-1.5 text-sm outline-none focus:border-zinc-900"
          />
          <button
            type="submit"
            disabled={!!busy || !input.trim()}
            className="rounded-lg bg-zinc-900 px-3 text-sm font-semibold text-white disabled:opacity-40"
          >
            보내기
          </button>
        </div>
      </form>
    </main>
  );
}
