import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "PPT 작성기",
  description: "템플릿을 고르고 채팅으로 PPT를 만들고 고치는 사내 도구",
};

export default function RootLayout({ children }: LayoutProps<"/">) {
  return (
    <html lang="ko" className="h-full antialiased">
      <body className="min-h-full flex flex-col">{children}</body>
    </html>
  );
}
