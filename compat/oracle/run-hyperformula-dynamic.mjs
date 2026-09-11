import { HyperFormula } from "hyperformula";

const data = [
  [1, "one", 2, -1000, new Date(2024, 0, 1)],
  [2, "two", 4, 200, new Date(2024, 6, 1)],
  [3, "three", 6, 1000, new Date(2025, 0, 1)],
  [2, "two", 0, null, null],
];

const cases = [
  ["vstack", "=VSTACK(A1:A2,C1:C2)", [[1], [2], [2], [4]]],
  ["hstack", "=HSTACK(A1:B2,C1:C2)", [[1, "one", 2], [2, "two", 4]]],
  ["unique", "=UNIQUE(A1:A4)", [[1], [2], [3]]],
  ["sort", "=SORT(C1:C3,1,-1)", [[6], [4], [2]]],
  ["xirr", "=XIRR(D1:D3,E1:E3)", 0.22046382967],
];

const hf = HyperFormula.buildFromArray(data, { licenseKey: "gpl-v3" });
const records = cases.map(([id, formula, expected]) => {
  const actual = hf.calculateFormula(formula, 0);
  const match = JSON.stringify(actual) === JSON.stringify(expected)
    || (typeof actual === "number" && Math.abs(actual - expected) <= 1e-10);
  return { id, formula, expected, actual, match };
});
const result = {
  schema_version: 1,
  oracle: "hyperformula",
  oracle_version: "3.4.0",
  license_key_mode: "gpl-v3",
  comparable_cases: records.length,
  matches: records.filter((record) => record.match).length,
  records,
  boundary: "Fixed HyperFormula oracle evidence; not Microsoft Excel compatibility proof.",
};
console.log(JSON.stringify(result, null, 2));
if (result.matches !== result.comparable_cases) process.exitCode = 1;
