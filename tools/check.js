/* 정리기를 node 로 돌려 본다 — 브라우저와 같은 엔진(clean.js)과 같은 파일 읽기(SheetJS).
 *
 *   node tools/check.js          사람이 읽는 표
 *   node tools/check.js --json   숫자만 (실측결과.md 대조용)
 *
 * 정답 파일(tools/answer/)은 배포하지 않는다(.vercelignore). 내 컴퓨터에서 대조할 때만 쓴다 (교재 12-9).
 */
const fs = require("fs");
const path = require("path");
const XLSX = require("../lib/xlsx.full.min.js");
const Clean = require("../clean.js");

const here = path.join(__dirname, "..");
const readTable = file => {
  const wb = XLSX.read(fs.readFileSync(file), { type: "buffer", raw: false, cellDates: false });
  const ws = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(ws, { defval: "", raw: true });
};
const readCsv = file => {
  const wb = XLSX.read(fs.readFileSync(file, "utf8").replace(/^﻿/, ""), { type: "string", raw: true });
  return XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: "", raw: false });
};
const rules = name => JSON.parse(fs.readFileSync(path.join(here, "rules", name + ".json"), "utf8"));

const out = {};

// ── 상담이력 ─────────────────────────────────────────────
{
  const raw = readTable(path.join(here, "data", "상담이력.xlsx"));
  const d = Clean.diagnose(raw);
  const r1 = Clean.apply(raw, rules("상담이력"));
  const r2 = Clean.apply(r1.rows, rules("상담이력"));
  const ans = readCsv(path.join(here, "tools", "answer", "상담이력_정답.csv"));
  const R = rules("상담이력");
  const cmp = Clean.compare(r1.rows, ans, R.answer.key, R.answer.compare);
  const match = cmp.match, notInAnswer = cmp.notInAnswer, diff = cmp.diff;
  const scaleMiss = cmp.examples.filter(e => e.col === "상담시간").map(e => `${e.key.split("|")[0]}: ${e.got} (정답 ${e.want})`);
  const scaleMissCount = cmp.diff["상담시간"] || 0;
  const 채널 = d.columns["채널"];
  out.상담이력 = {
    rows: raw.length,
    channel_kinds: Object.keys(채널.values || {}).length,
    duration_max: d.columns["상담시간"].stats.max,
    dup_before_rules: d.duplicates,
    after: r1.summary.after, dropped: r1.summary.dropped, changes: r1.summary.changes,
    idempotent: Clean.same(r1.rows, r2.rows) && r2.summary.changes === 0,
    answer_rows: ans.length, match, not_in_answer: notInAnswer, diff,
    scale_rescaled: r1.log.filter(l => l.col === "상담시간").length,
  };
  out.상담이력.scale_missed = scaleMissCount;
  out._scaleMissExamples = scaleMiss.slice(0, 5);
  out._symptoms1 = d.symptoms.map(s => s.text);
}

// ── 교육이수로그 ─────────────────────────────────────────
{
  const raw = readCsv(path.join(here, "data", "교육이수로그.csv"));
  const d = Clean.diagnose(raw);
  const r1 = Clean.apply(raw, rules("교육이수로그"));
  const r2 = Clean.apply(r1.rows, rules("교육이수로그"));
  const flags = {};
  r1.rows.forEach(r => (r[Clean.FLAG] ? r[Clean.FLAG].split(" · ") : []).forEach(f => {
    const k = f.replace(/\(.*\)/, "");
    flags[k] = (flags[k] || 0) + 1;
  }));
  out.교육이수로그 = {
    rows: raw.length, after: r1.summary.after, flagged: r1.summary.flagged, flags,
    changes: r1.summary.changes,
    idempotent: Clean.same(r1.rows, r2.rows) && r2.summary.changes === 0,
  };
  out._symptoms2 = d.symptoms.map(s => s.text);
}

if (process.argv.includes("--json")) {
  const { _symptoms1, _symptoms2, _scaleMissExamples, ...nums } = out;
  console.log(JSON.stringify(nums));
} else {
  const a = out.상담이력, b = out.교육이수로그;
  console.log("상담이력.xlsx — 진단 (규칙 없이)");
  out._symptoms1.forEach(s => console.log("  · " + s));
  console.log(`규칙 적용: ${a.rows}행 → ${a.after}행 (완전 중복 ${a.dropped}행 지움) · 값 고친 곳 ${a.changes}군데 · 초→분 ${a.scale_rescaled}행`);
  console.log(`두 번 돌려도 같은가: ${a.idempotent ? "같다" : "다르다"}`);
  console.log(`정답 대조: 정답 ${a.answer_rows}행 중 일치 ${a.match}행 · 정답에 없는 행(재발송) ${a.not_in_answer}행 · 어긋난 열 ${JSON.stringify(a.diff)}`);
  console.log(`  180 기준이 놓친 초 단위 ${a.scale_missed}행 — 예: ${out._scaleMissExamples.join(" / ")}`);
  console.log("");
  console.log("교육이수로그.csv — 진단 (규칙 없이)");
  out._symptoms2.forEach(s => console.log("  · " + s));
  console.log(`규칙 적용: ${b.rows}행 → ${b.after}행 · 값 고친 곳 ${b.changes}군데 · 표시한 행 ${b.flagged}행 ${JSON.stringify(b.flags)}`);
  console.log(`두 번 돌려도 같은가: ${b.idempotent ? "같다" : "다르다"}`);
}
