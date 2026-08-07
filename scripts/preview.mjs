// Builds screenshot harness pages from the REAL dist taskpane HTML:
// strips office.js + app bundle scripts, keeps every class/ID/data-*,
// injects a driver that populates chat states via shared/proposal-card.js.
import { readFileSync, writeFileSync, mkdirSync, copyFileSync } from "node:fs";
import { join } from "node:path";

const root = process.cwd();
const out = join(root, ".preview");
mkdirSync(out, { recursive: true });
copyFileSync(join(root, "shared/proposal-card.js"), join(out, "proposal-card.js"));
copyFileSync(join(root, "shared/design-system.css"), join(out, "design-system.css"));

const driver = `
<script type="module">
import { appendMessage, appendTypingIndicator, renderProposalCard, mountContextBar, setStatus } from "./proposal-card.js";
const q = new URLSearchParams(location.search);
const dark = q.get("dark") === "1";
if (dark) document.documentElement.classList.add("dark");
if (q.get("light") === "1") document.documentElement.classList.add("light");
const state = q.get("state") || "empty";
const host = q.get("pane"); // word | excel

const log = document.getElementById("log");
const preview = document.getElementById("preview");
const empty = document.getElementById("empty");
const apply = document.getElementById("apply");
const toggle = document.getElementById("markRedWrap") || document.getElementById("highlightWrap");
const statusEl = document.getElementById("status");

if (state === "chat") {
  empty.classList.add("is-hidden");
  appendMessage(log, "user", "Soát lỗi chính tả đoạn em vừa chọn giúp anh nhé.");
  appendMessage(log, "bot", "Mình đã đọc đoạn được ghim (214 từ). Có 3 chỗ cần sửa — xem thẻ đề xuất bên dưới.");
  appendMessage(log, "bot", "Lưu ý: một chỗ nằm trong bảng, Apply sẽ thay đúng ô đó.", { tone: "warn" });
  appendTypingIndicator(log);

  const actions = host === "excel" ? [
    { type: "setCell", cell: "B4", old: "1200000", new: "1.200.000" },
    { type: "format", range: "C2:C9", numberFormat: "#,##0", bold: true },
    { type: "setCells", range: "E2:F5", values: [["Q1", "Q2"], [120, 140], [98, 115], [130, 152]] },
  ] : [
    { type: "replace", find: "sử dụng", replace: "dùng", matchCount: 4 },
    { type: "replace", find: "một cách nhanh chóng", replace: "nhanh", matchCount: 2 },
  ];
  renderProposalCard(preview, { title: host === "excel" ? "Chuẩn hoá bảng doanh thu" : "Biên tập đoạn được ghim", actions, badgeAction: host === "excel" ? (a) => a.type === "setCells" ? "ghi đè" : null : null });

  apply.hidden = false;
  if (toggle) { toggle.hidden = false; toggle.dataset.show = "true"; }
  setStatus(statusEl, "Đang chờ duyệt đề xuất…");
  document.getElementById("statusRow").dataset.tone = "busy";
} else {
  setStatus(statusEl, "Sẵn sàng.");
}

if (host === "word") {
  mountContextBar(document.getElementById("contextRow"), [
    { label: "Ghim", value: "“…doanh thu quý II tăng 18% so với cùng kỳ…”", state: "pinned" },
    { label: "Snapshot", value: "214 từ" },
  ]);
}
</script>`;

for (const pane of ["word", "excel"]) {
  let html = readFileSync(join(root, pane, "dist/taskpane.html"), "utf8");
  // Drop the app bundle + office.js; keep the font + CSS links.
  html = html.replace(/<script[^>]*appsforoffice[^>]*><\/script>/, "");
  html = html.replace(/<script[^>]*src="[^"]*taskpane[^"]*"[^>]*><\/script>/, "");
  html = html.replace(/<script[^>]*src="[^"]*polyfill[^"]*"[^>]*><\/script>/, "");
  // Point CSS at the local copy so dark/light is driven by file, not dist.
  html = html.replace('href="assets/design-system.css"', 'href="design-system.css"');
  html = html.replace("</body>", driver + "\n</body>");
  writeFileSync(join(out, `${pane}.html`), html);
}
console.log("preview pages ready:", out);
