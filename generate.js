#!/usr/bin/env node
// エントリポイント: JSONデータを読み込んでPPTX生成
// 使い方:
//   単品: node generate.js templates/lifestyle-tumbler.json
//   バンドル: node generate.js templates/bundle-naruse.json
const { generate, generateBundle } = require("./lib/slidekit");
const fs = require("fs");
const path = require("path");

const jsonPath = process.argv[2];
if (!jsonPath) {
  console.error("使い方: node generate.js <template.json>");
  console.error("例: node generate.js templates/lifestyle-tumbler.json");
  console.error("例: node generate.js templates/bundle-naruse.json");
  process.exit(1);
}

const data = JSON.parse(fs.readFileSync(jsonPath, "utf8"));
const outPath = data.output || `/tmp/${data.id || "proposal"}.pptx`;

if (data.type === "bundle") {
  // バンドルモード: baseDir未指定なら JSONファイルのディレクトリを使う
  const resolved = { ...data, baseDir: data.baseDir || path.dirname(path.resolve(jsonPath)) };
  generateBundle(resolved, outPath).catch(e => { console.error(e); process.exit(1); });
} else {
  const resolved = { ...data, baseDir: data.baseDir || path.dirname(path.resolve(jsonPath)) };
  generate(resolved, outPath).catch(e => { console.error(e); process.exit(1); });
}
