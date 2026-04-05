#!/usr/bin/env node
// エントリポイント: JSONデータを読み込んでPPTX生成
// 使い方: node generate.js templates/food-seasoning.json
const { generate } = require("./lib/slidekit");
const fs = require("fs");

const jsonPath = process.argv[2];
if (!jsonPath) {
  console.error("使い方: node generate.js <template.json>");
  console.error("例: node generate.js templates/food-seasoning.json");
  process.exit(1);
}

const data = JSON.parse(fs.readFileSync(jsonPath, "utf8"));
const outPath = data.output || `/tmp/${data.id || "proposal"}.pptx`;

generate(data, outPath).catch(e => { console.error(e); process.exit(1); });
