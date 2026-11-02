// PowerPoint 작업창은 HTTPS 주소만 열 수 있어서, office-addin-dev-certs가 만든 인증서로 개발 서버를 HTTPS로 띄움
import { spawn } from "node:child_process";
import { existsSync } from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";

const dir = join(homedir(), ".office-addin-dev-certs");
const key = join(dir, "localhost.key");
const cert = join(dir, "localhost.crt");

if (!existsSync(key) || !existsSync(cert)) {
  console.error("개발용 HTTPS 인증서가 없어요. 먼저 아래 명령을 한 번 실행해 주세요.\n  npx office-addin-dev-certs install");
  process.exit(1);
}

// 셸을 거치지 않고 node로 next를 직접 실행해, 경로에 공백이 있어도 안전하게 넘김
const nextBin = join(process.cwd(), "node_modules", "next", "dist", "bin", "next");
const child = spawn(
  process.execPath,
  [nextBin, "dev", "--experimental-https", "--experimental-https-key", key, "--experimental-https-cert", cert],
  { stdio: "inherit" },
);
child.on("exit", (code) => process.exit(code ?? 0));
