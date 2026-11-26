import fs from "node:fs";
import { fileURLToPath } from "node:url";

const configPath = fileURLToPath(new URL("../../.workflow-state/sellersprite/browser-config.json", import.meta.url));

export function getDebuggerAddress() {
  // 配置由启动器在确认浏览器归属后保存，旧配置没有端口时兼容9223。
  if (!fs.existsSync(configPath)) return "127.0.0.1:9223";
  const config = JSON.parse(fs.readFileSync(configPath, "utf8").replace(/^\uFEFF/, ""));
  const address = config.cdp_url || config.debugger_address;
  const value = config.port ?? (address ? new URL(address.includes("://") ? address : `http://${address}`).port : 9223);
  const port = Number(value);
  if (typeof value === "boolean" || !/^\d+$/.test(String(value)) || !Number.isInteger(port) || port < 1 || port > 65535) {
    throw new Error(`项目浏览器端口记录无效，请重新运行浏览器启动器：${configPath}`);
  }
  return `127.0.0.1:${port}`;
}

export function getCdpUrl() {
  return `http://${getDebuggerAddress()}`;
}
