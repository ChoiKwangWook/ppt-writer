"use client";

import { useEffect, useRef, useState } from "react";
import { downloadPptx } from "@/lib/pptx";
import { getTemplate, TEMPLATES } from "@/lib/templates";
import type { Deck, TemplateId } from "@/lib/types";
import SlidePreview from "./SlidePreview";

type Message = { role: "user" | "assistant"; content: string };

const TIMEOUT_MS = 30_000;
const EXAMPLES = [
  "신입사원 오리엔테이션 자료 5장 만들어줘",
  "모든 제목을 파란색 굵은 글씨로 바꿔줘",
  "2번 슬라이드에 3단계 흐름도를 넣어줘",
];

export default function PptChat() {
  const [templateId, setTemplateId] = useState<TemplateId | null>(null);
  const [deck, setDeck] = useState<Deck>({ slides: [] });
  const [prevDeck, setPrevDeck] = useState<Deck | null>(null);
  const [messages, setMessages] = useState<Message[]>([]);
  const [input, setInput] = useState("");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<{ text: string; retry: string } | null>(null);
  const [downloading, setDownloading] = useState(false);
  const chatEnd = useRef<HTMLDivElement>(null);

  useEffect(() => {
    chatEnd.current?.scrollIntoView({ behavior: "smooth" });
  }, [messages, loading, error]);

  const template = templateId ? getTemplate(templateId) : null;

  async function send(text: string, isRetry = false) {
    const message = text.trim();
    if (!message || loading) return;

    // 다시 시도할 때는 이미 화면에 있는 마지막 요청을 다시 보냄
    const history = isRetry ? messages.slice(0, -1) : messages;
    if (!isRetry) {
      setMessages([...messages, { role: "user", content: message }]);
      setInput("");
    }
    setError(null);
    setLoading(true);

    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), TIMEOUT_MS);
    try {
      const res = await fetch("/api/chat", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ message, deck, history }),
        signal: controller.signal,
      });
      const data = await res.json().catch(() => ({}));
      if (!res.ok) throw new Error(data.error || "AI 응답을 받지 못했어요. 잠시 후 다시 시도해 주세요.");

      setMessages((m) => [...m, { role: "assistant", content: data.reply }]);
      if (data.deck) {
        setPrevDeck(deck);
        setDeck(data.deck);
      }
    } catch (e) {
      const aborted = e instanceof DOMException && e.name === "AbortError";
      setError({
        text: aborted ? "30초 안에 응답이 오지 않았어요. 잠시 후 다시 시도해 주세요." : (e as Error).message,
        retry: message,
      });
    } finally {
      clearTimeout(timer);
      setLoading(false);
    }
  }

  function undo() {
    if (!prevDeck) return;
    setDeck(prevDeck);
    setPrevDeck(null);
    setMessages((m) => [...m, { role: "assistant", content: "직전 상태로 되돌렸어요." }]);
  }

  async function download() {
    if (!template || !deck.slides.length) return;
    setDownloading(true);
    try {
      await downloadPptx(deck, template);
    } catch (e) {
      console.error(e);
      setError({ text: "파일을 만들지 못했어요. 다시 시도해 주세요.", retry: "" });
    } finally {
      setDownloading(false);
    }
  }

  return (
    <div className="flex min-h-screen flex-col bg-zinc-100 text-zinc-900">
      <header className="flex flex-wrap items-center justify-between gap-3 border-b border-zinc-200 bg-white px-4 py-3 sm:px-6">
        <div className="text-lg font-bold">PPT 작성기</div>
        <div className="flex flex-wrap items-center gap-2">
          <span className="text-sm text-zinc-500">템플릿</span>
          {TEMPLATES.map((t) => (
            <button
              key={t.id}
              onClick={() => setTemplateId(t.id)}
              className={`flex items-center gap-2 rounded-full border px-3 py-1.5 text-sm transition ${
                templateId === t.id ? "border-zinc-900 bg-zinc-900 text-white" : "border-zinc-300 bg-white hover:bg-zinc-50"
              }`}
            >
              <span className="h-3 w-3 rounded-full ring-1 ring-black/20" style={{ background: t.accent }} />
              {t.name}
            </button>
          ))}
        </div>
      </header>

      {!template ? (
        <main className="mx-auto w-full max-w-5xl flex-1 px-4 py-10 sm:px-6">
          <h1 className="mb-8 text-4xl font-extrabold tracking-tight sm:text-5xl">PPT 작성기</h1>
          <h2 className="text-xl font-bold">1. 시작할 템플릿을 골라 주세요</h2>
          <p className="mt-1 text-sm text-zinc-600">나중에 내용은 그대로 두고 템플릿만 바꿔 볼 수도 있어요.</p>
          <div className="mt-6 grid gap-5 sm:grid-cols-3">
            {TEMPLATES.map((t) => (
              <button
                key={t.id}
                onClick={() => setTemplateId(t.id)}
                className="rounded-xl bg-white p-3 text-left shadow-sm ring-1 ring-zinc-200 transition hover:ring-2 hover:ring-zinc-900"
              >
                <SlidePreview
                  template={t}
                  slide={{ title: `${t.name} 템플릿`, bullets: ["채팅으로 내용을 입력하면", "AI가 슬라이드를 만들어 줘요"] }}
                />
                <div className="mt-3 font-semibold">{t.name}</div>
                <div className="text-sm text-zinc-600">{t.description}</div>
              </button>
            ))}
          </div>
        </main>
      ) : (
        <main className="grid flex-1 gap-4 p-4 sm:p-6 lg:grid-cols-[minmax(320px,400px)_1fr] lg:overflow-hidden">
          <section className="flex min-h-[420px] flex-col rounded-xl bg-white shadow-sm ring-1 ring-zinc-200 lg:h-[calc(100vh-110px)]">
            <div className="flex-1 space-y-3 overflow-y-auto p-4">
              {messages.length === 0 && (
                <div className="space-y-2">
                  <p className="text-sm text-zinc-600">만들고 싶은 발표 자료의 주제와 내용을 알려 주세요. 예를 들면:</p>
                  {EXAMPLES.map((ex) => (
                    <button
                      key={ex}
                      onClick={() => setInput(ex)}
                      className="block w-full rounded-lg border border-zinc-200 px-3 py-2 text-left text-sm hover:bg-zinc-50"
                    >
                      {ex}
                    </button>
                  ))}
                </div>
              )}
              {messages.map((m, i) => (
                <div key={i} className={`flex ${m.role === "user" ? "justify-end" : "justify-start"}`}>
                  <div
                    className={`max-w-[85%] whitespace-pre-wrap rounded-2xl px-3.5 py-2 text-sm ${
                      m.role === "user" ? "bg-zinc-900 text-white" : "bg-zinc-100 text-zinc-900"
                    }`}
                  >
                    {m.content}
                  </div>
                </div>
              ))}
              {loading && <div className="text-sm text-zinc-500">슬라이드를 만들고 있어요…</div>}
              {error && (
                <div className="rounded-lg border border-red-200 bg-red-50 p-3 text-sm text-red-800">
                  {error.text}
                  {error.retry && (
                    <button
                      onClick={() => send(error.retry, true)}
                      className="ml-2 rounded-md bg-red-700 px-2.5 py-1 text-xs font-semibold text-white hover:bg-red-800"
                    >
                      다시 시도
                    </button>
                  )}
                </div>
              )}
              <div ref={chatEnd} />
            </div>

            <form
              className="border-t border-zinc-200 p-3"
              onSubmit={(e) => {
                e.preventDefault();
                send(input);
              }}
            >
              <p className="mb-2 rounded-md bg-amber-50 px-2.5 py-1.5 text-xs text-amber-800">
                ⚠ 대외비·개인정보는 입력하지 마세요.
              </p>
              <div className="flex gap-2">
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
                  className="flex-1 resize-none rounded-lg border border-zinc-300 px-3 py-2 text-sm outline-none focus:border-zinc-900"
                />
                <button
                  type="submit"
                  disabled={loading || !input.trim()}
                  className="rounded-lg bg-zinc-900 px-4 text-sm font-semibold text-white disabled:opacity-40"
                >
                  보내기
                </button>
              </div>
            </form>
          </section>

          <section className="flex min-h-0 flex-col rounded-xl bg-white shadow-sm ring-1 ring-zinc-200 lg:h-[calc(100vh-110px)]">
            <div className="flex flex-wrap items-center justify-between gap-2 border-b border-zinc-200 px-4 py-3">
              <div className="text-sm text-zinc-600">
                미리보기 · {template.name} · {deck.slides.length}장
              </div>
              <div className="flex gap-2">
                <button
                  onClick={undo}
                  disabled={!prevDeck || loading}
                  className="rounded-lg border border-zinc-300 px-3 py-1.5 text-sm hover:bg-zinc-50 disabled:opacity-40"
                >
                  되돌리기
                </button>
                <button
                  onClick={download}
                  disabled={!deck.slides.length || downloading}
                  className="rounded-lg bg-zinc-900 px-3 py-1.5 text-sm font-semibold text-white disabled:opacity-40"
                >
                  {downloading ? "만드는 중…" : ".pptx 다운로드"}
                </button>
              </div>
            </div>
            <div className="flex-1 overflow-y-auto p-4">
              {deck.slides.length === 0 ? (
                <div className="flex h-full min-h-[200px] items-center justify-center text-sm text-zinc-500">
                  왼쪽 채팅에 내용을 입력하면 여기에 슬라이드가 나타나요.
                </div>
              ) : (
                <div className="grid gap-5 xl:grid-cols-2">
                  {deck.slides.map((s, i) => (
                    <div key={i}>
                      <div className="mb-1 text-xs font-semibold text-zinc-500">{i + 1}</div>
                      <SlidePreview slide={s} template={template} />
                    </div>
                  ))}
                </div>
              )}
            </div>
          </section>
        </main>
      )}
    </div>
  );
}
