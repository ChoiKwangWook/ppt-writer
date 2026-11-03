// PowerPoint 안의 슬라이드를 넣고·지우고·표식을 다룸. 작업창(브라우저)에서만 실행됨.
// PowerPoint 2024에서 되는 기능만 씀: 슬라이드 넣기(PowerPointApi 1.2), 표식·슬라이드 마스터(1.3)
// Design Ref: §4.3 — 사람이 만든 슬라이드는 건드리지 않고, 표식이 있는 추가 기능 슬라이드만 지우고 다시 넣음
import { buildPptxBase64 } from "@/lib/pptx";
import { sanitizeDeck } from "@/lib/sanitize";
import { TEMPLATES, type Template } from "@/lib/templates";
import type { Deck, TemplateId } from "@/lib/types";

// 추가 기능이 넣은 슬라이드에 붙이는 표식의 이름 (PowerPoint가 대문자로 저장함)
export const TAG_KEY = "PPTWRITER";
export const FAIL_TEXT = "PowerPoint에 슬라이드를 넣지 못했어요. 다시 시도해 주세요.";
// 슬라이드 데이터를 PowerPoint 파일 안(문서 설정)에 저장할 때 쓰는 이름
const SETTINGS_KEY = "pptWriter";

type SlideInfo = { id: string; tag: string | null };
type Context = PowerPoint.RequestContext;

export type RenderResult = {
  inserted: number; // 새로 넣은 슬라이드 수
  leftoverOld: number; // 지우지 못하고 남은 예전 추가 기능 슬라이드 수 (다음 반영 때 다시 지움)
  renderId: string | null; // 이번에 넣은 슬라이드 표식의 앞부분 (넣은 슬라이드가 없으면 null)
};

// renderDeck·removeAddinSlides의 결과: 파일 안 저장까지 끝났는지 함께 알려 줌
export type RenderOutcome = RenderResult & { saveFailed: boolean };

// 파일 안에 저장하는 내용. 표식 "{renderId}-{순번}"의 순번이 deck.slides의 몇 번째인지를 가리킴
type Saved = { deck: Deck; templateId: TemplateId; renderId: string };

// 파일 안에서 저장된 데이터를 읽음. 내용이 이상하면 믿지 않고 없는 것으로 봄 (sanitizeDeck으로 한 번 더 다듬음)
function loadSaved(): Saved | null {
  const raw = Office.context.document.settings.get(SETTINGS_KEY) as Partial<Saved> | null | undefined;
  if (!raw || typeof raw !== "object") return null;
  const templateId = TEMPLATES.find((t) => t.id === raw.templateId)?.id;
  if (!templateId || typeof raw.renderId !== "string") return null;
  return { deck: sanitizeDeck(raw.deck).deck, templateId, renderId: raw.renderId };
}

// 데이터를 메모리에 넣은 뒤 saveAsync로 파일에 저장 요청함 (set만으로는 파일에 남지 않음)
function saveSettings(saved: Saved): Promise<void> {
  const settings = Office.context.document.settings;
  settings.set(SETTINGS_KEY, saved);
  return new Promise((resolve, reject) => {
    settings.saveAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
      else reject(result.error);
    });
  });
}

// 작업 도중 한 번 실패해도 바로 다시 시도해 봄 (일시적인 PowerPoint 지연 대비)
async function retryOnce<T>(job: () => Promise<T>): Promise<T> {
  try {
    return await job();
  } catch {
    return await job();
  }
}

// 모든 슬라이드의 번호(id)와 표식 값을 순서대로 읽음. 표식이 없으면 사람이 만든 슬라이드
async function readSlides(context: Context): Promise<SlideInfo[]> {
  const slides = context.presentation.slides;
  slides.load("items/id");
  await context.sync();
  const tagLists = slides.items.map((s) => {
    s.tags.load("items/key,items/value");
    return s.tags;
  });
  await context.sync();
  return slides.items.map((s, i) => ({
    id: s.id,
    tag: tagLists[i].items.find((t) => t.key === TAG_KEY)?.value ?? null,
  }));
}

// 사용자가 추가 기능 슬라이드를 복제(Ctrl+D)하면 표식까지 복사됨.
// 같은 표식 값이 둘 이상이면 첫 번째만 추가 기능 슬라이드로 보고, 복제본은 표식을 떼어 사람 슬라이드로 취급함
async function stripDuplicateTags(context: Context, slides: SlideInfo[]): Promise<SlideInfo[]> {
  const seen = new Set<string>();
  const result = slides.map((s) => {
    if (s.tag === null) return s;
    if (seen.has(s.tag)) return { id: s.id, tag: null };
    seen.add(s.tag);
    return s;
  });
  const changed = result.filter((s, i) => slides[i].tag !== null && s.tag === null);
  if (changed.length) {
    changed.forEach((s) => context.presentation.slides.getItem(s.id).tags.delete(TAG_KEY));
    await context.sync();
  }
  return result;
}

// 지정한 슬라이드를 지움. requireTag가 true면 지금도 표식이 있는 슬라이드만 지워서 사람 슬라이드를 지키는 마지막 안전장치로 씀
async function deleteSlides(context: Context, ids: string[], requireTag: boolean) {
  const current = await readSlides(context);
  const targets = current.filter((s) => ids.includes(s.id) && (!requireTag || s.tag !== null));
  targets.forEach((s) => context.presentation.slides.getItem(s.id).delete());
  if (targets.length) await context.sync();
}

/**
 * 추가 기능 슬라이드를 새 .pptx 내용(base64)으로 바꿔 넣음. base64가 null이면 추가 기능 슬라이드만 모두 지움.
 * 순서: 넣기 → 표식 붙이기 → 예전 슬라이드 지우기 (넣기가 실패해도 예전 슬라이드가 그대로 남도록)
 */
export async function replaceAddinSlides(base64: string | null): Promise<RenderResult> {
  return PowerPoint.run(async (context) => {
    const before = await stripDuplicateTags(context, await readSlides(context));
    const oldIds = before.filter((s) => s.tag !== null).map((s) => s.id);
    const beforeIds = new Set(before.map((s) => s.id));
    let insertedIds: string[] = [];
    let renderId: string | null = null;

    if (base64) {
      // 첫 추가 기능 슬라이드 바로 앞에 넣고, 없으면 파일 맨 끝에 붙임 (앞 슬라이드가 없으면 맨 앞에 들어감)
      const firstAddin = before.findIndex((s) => s.tag !== null);
      const targetId = firstAddin === -1 ? before[before.length - 1]?.id : before[firstAddin - 1]?.id;
      context.presentation.insertSlidesFromBase64(base64, {
        formatting: "KeepSourceFormatting",
        ...(targetId ? { targetSlideId: targetId } : {}),
      });
      await context.sync();

      // 넣기 전에 없던 슬라이드가 이번에 들어온 슬라이드
      insertedIds = (await readSlides(context)).filter((s) => !beforeIds.has(s.id)).map((s) => s.id);
      const newRenderId = `r${Date.now().toString(36)}`;
      renderId = newRenderId;
      try {
        await retryOnce(async () => {
          insertedIds.forEach((id, i) =>
            context.presentation.slides.getItem(id).tags.add(TAG_KEY, `${newRenderId}-${i + 1}`),
          );
          await context.sync();
        });
      } catch (error) {
        // 표식을 못 붙이면 표식 없는 새 슬라이드가 남아 영영 못 지우게 되므로, 새로 들어온 것만 도로 지워 원래대로 돌림
        await deleteSlides(context, insertedIds, false);
        throw error;
      }
    }

    // 예전 추가 기능 슬라이드는 표식이 지금도 있는 것만 지움
    let leftoverOld = 0;
    if (oldIds.length) {
      try {
        await retryOnce(() => deleteSlides(context, oldIds, true));
      } catch {
        leftoverOld = oldIds.length;
      }
    }
    return { inserted: insertedIds.length, leftoverOld, renderId };
  });
}

// 슬라이드 데이터와 템플릿으로 PowerPoint의 추가 기능 슬라이드를 새로 넣음 (템플릿 바꾸기·되돌리기도 이 함수로 처리)
export async function renderDeck(deck: Deck, t: Template): Promise<RenderOutcome> {
  let result: RenderResult;
  try {
    const base64 = deck.slides.length ? await buildPptxBase64(deck, t) : null;
    result = await replaceAddinSlides(base64);
  } catch (error) {
    console.error("renderDeck error:", error);
    throw new Error(FAIL_TEXT);
  }
  // 슬라이드는 이미 들어갔으므로 저장이 실패해도 오류로 막지 않고, 호출한 쪽이 안내할 수 있게 알려 줌
  const saveFailed = await trySave({ deck, templateId: t.id, renderId: result.renderId ?? "" });
  return { ...result, saveFailed };
}

// 추가 기능이 만든 슬라이드만 모두 지움 (사람이 만든 슬라이드는 그대로)
export async function removeAddinSlides(): Promise<RenderOutcome> {
  let result: RenderResult;
  try {
    result = await replaceAddinSlides(null);
  } catch (error) {
    console.error("removeAddinSlides error:", error);
    throw new Error("PowerPoint에서 슬라이드를 지우지 못했어요. 다시 시도해 주세요.");
  }
  const templateId = loadSaved()?.templateId ?? TEMPLATES[0].id;
  const saveFailed = await trySave({ deck: { slides: [] }, templateId, renderId: "" });
  return { ...result, saveFailed };
}

// 저장을 한 번 더 시도하고, 그래도 실패하면 true를 돌려줌
async function trySave(saved: Saved): Promise<boolean> {
  try {
    await retryOnce(() => saveSettings(saved));
    return false;
  } catch (error) {
    console.error("save settings error:", error);
    return true;
  }
}

export type SyncResult = {
  deck: Deck; // 화면(PowerPoint)에 지금 있는 추가 기능 슬라이드의 데이터, 화면 순서대로
  templateId: TemplateId | null; // 저장돼 있던 템플릿 (저장된 것이 없으면 null)
  positions: number[]; // deck의 각 슬라이드가 PowerPoint 화면 몇 번인지 (1부터)
  humanSlides: number[]; // 사람이 만든 슬라이드의 화면 번호
  changed: boolean; // 사용자가 PowerPoint에서 추가 기능 슬라이드를 지우거나 옮겨서 저장된 데이터와 달라졌는지
};

/**
 * 파일 안에 저장된 데이터와 PowerPoint 화면을 맞춰서, 지금 화면 기준의 데이터와 번호 대응표를 돌려줌 (읽기 전용 + 복제본 표식 정리).
 * 작업창을 열 때와 슬라이드를 반영하기 직전에 부름. 저장된 데이터 자체는 고치지 않으므로, 다음 반영 때 새로 저장됨.
 */
export async function syncFromPresentation(): Promise<SyncResult> {
  return PowerPoint.run(async (context) => {
    const slides = await stripDuplicateTags(context, await readSlides(context));
    const saved = loadSaved();

    // 저장된 표식("{renderId}-{순번}")이 지금 화면에 있는 슬라이드만 찾음. 순번이 저장된 deck의 몇 번째 슬라이드인지 알려 줌
    const found: { screenIndex: number; seq: number }[] = [];
    slides.forEach((s, screenIndex) => {
      const m = saved && s.tag ? /^(.+)-(\d+)$/.exec(s.tag) : null;
      if (!saved || !m || m[1] !== saved.renderId) return;
      const seq = Number(m[2]);
      if (seq >= 1 && seq <= saved.deck.slides.length) found.push({ screenIndex, seq });
    });

    const deck: Deck = { slides: saved ? found.map((f) => saved.deck.slides[f.seq - 1]) : [] };
    const inOrder = found.every((f, i) => i === 0 || f.seq > found[i - 1].seq);
    return {
      deck,
      templateId: saved?.templateId ?? null,
      positions: found.map((f) => f.screenIndex + 1),
      humanSlides: slides.flatMap((s, i) => (s.tag === null ? [i + 1] : [])),
      changed: !!saved && (found.length !== saved.deck.slides.length || !inOrder),
    };
  });
}

export type Overview = {
  slides: { number: number; tag: string | null }[];
  masters: number; // 슬라이드 마스터 수 (넣을 때마다 계속 늘어나는지 확인용)
};

// 지금 파일의 슬라이드 현황을 읽기만 함 (아무것도 바꾸지 않음)
export async function readOverview(): Promise<Overview> {
  return PowerPoint.run(async (context) => {
    const slides = await readSlides(context);
    const masters = context.presentation.slideMasters;
    masters.load("items/id");
    await context.sync();
    return { slides: slides.map((s, i) => ({ number: i + 1, tag: s.tag })), masters: masters.items.length };
  });
}
