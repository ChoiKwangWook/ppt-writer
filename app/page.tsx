// PPT 작성기는 PowerPoint 안의 작업창에서 쓰므로, 일반 브라우저로 열면 여는 방법만 안내함
export default function Home() {
  return (
    <main className="mx-auto flex min-h-screen w-full max-w-xl flex-col gap-4 px-6 py-16 text-zinc-900">
      <h1 className="text-3xl font-extrabold">PPT 작성기</h1>
      <p className="text-zinc-700">PPT 작성기는 PowerPoint 안에서 쓰는 추가 기능이에요. PowerPoint에서 열어 주세요.</p>
      <ol className="list-decimal space-y-1 pl-5 text-sm text-zinc-700">
        <li>Windows용 PowerPoint 2024 또는 Microsoft 365를 열어요.</li>
        <li>홈 탭 오른쪽 끝의 [PPT 작성기] 버튼을 눌러요.</li>
        <li>오른쪽 작업창에서 템플릿을 고르고 채팅으로 슬라이드를 만들어요.</li>
      </ol>
      <p className="text-xs text-zinc-500">버튼이 보이지 않으면 추가 기능을 설치한 담당자에게 문의해 주세요.</p>
    </main>
  );
}
