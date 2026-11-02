import "server-only";
import OpenAI from "openai";

// 비밀 키는 서버에서만 .env의 OPENAI_API_KEY로 읽음 (브라우저로 절대 보내지 않음)
export function getOpenAI() {
  const apiKey = process.env.OPENAI_API_KEY;
  if (!apiKey) {
    throw new Error("OPENAI_API_KEY가 .env에 설정되어 있지 않습니다.");
  }
  // 30초 안에 답이 없으면 실패로 보고 화면에서 [다시 시도]를 안내함
  return new OpenAI({ apiKey, timeout: 30_000, maxRetries: 0 });
}

export const CHAT_MODEL = "gpt-4o-mini";
