// Evaluate real generated formulas in an independent spreadsheet engine.
// Usage: set NODE_PATH to the bundled node_modules; node this-file.mjs <python>
// Synthetic fixtures only: this never reads or changes patient workbooks.
import assert from "node:assert/strict";
import { execFileSync } from "node:child_process";
import { createRequire } from "node:module";
import { fileURLToPath, pathToFileURL } from "node:url";

const require = createRequire(import.meta.url);
const { Workbook } = await import(pathToFileURL(require.resolve("@oai/artifact-tool")).href);
const formulas = JSON.parse(execFileSync(
  process.argv[2] || "python",
  ["-m", "qa.test_special_formula_bindings", "--fixtures"],
  {
    cwd: fileURLToPath(new URL("../", import.meta.url)),
    encoding: "utf8",
    env: { ...process.env, PYTHONIOENCODING: "utf-8" },
    windowsHide: true,
  },
));
const codes = Object.keys(formulas);
const workbook = Workbook.create();
const sheet = workbook.worksheets.add("FormulaRegression");
const output = sheet.getRangeByIndexes(4, 12, 1, codes.length);
output.formulas = [codes.map(code => formulas[code])];

const scenarios = [
  { k: "腎", expected: ["107", "110"] },
  { k: "胰", expected: ["125"] },
  { k: "肝", expected: ["101"] },
  { k: "腎,胰,肝", expected: ["107", "110", "125", "101"] },
  { k: "", expected: [] },
  { k: "A", expected: ["107", "110"] },
  { k: "ABCD.V7,炎", expected: ["107", "110", "125", "147", "101", "200"] },
  { k: "V14.", expected: ["199"] },
  { k: "V140.", expected: [] },
  { k: "V7.", expected: ["125", "147", "101", "200"] },
  { k: "V70.", expected: [] },
  { k: "腎", l: "V14", expected: ["107", "110", "199"] },
  { k: "", l: "V14", expected: ["199"] },
  { k: "", l: "V140", expected: [] },
  { k: "", l: "A,V14", expected: ["107", "110", "199"] },
  { k: "腎", h: "空號", expected: [] },
  { k: "腎", h: 0, expected: [] },
];

for (const { k, l = "", h = "測試受檢者", expected } of scenarios) {
  sheet.getRange("H5:L5").values = [[h, 1, "", k, l]];
  workbook.recalculate();
  const values = output.values[0];
  assert.ok(values.every(value => value === "@" || value === ""), JSON.stringify(values));
  const actual = codes.filter((code, index) => values[index] === "@").sort();
  assert.deepEqual(actual, [...expected].sort(), `K=${k}, L=${l}, H=${h}`);
  console.log(`PASS K=${JSON.stringify(k)}, L=${JSON.stringify(l)}, H=${h}: ${actual.join(",") || "(none)"}`);
}
console.log(`PASS: ${scenarios.length} formula-result scenarios`);
