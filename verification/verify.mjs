import crypto from "node:crypto";
import fs from "node:fs";
import fsp from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import process from "node:process";
import { spawn } from "node:child_process";
import { inflateRawSync } from "node:zlib";

const repoRoot = path.resolve(import.meta.dirname, "..");
const artifactRoot = path.join(repoRoot, "artifacts");
const evidenceRoot = path.join(repoRoot, "verification", "evidence");
const npmCli = process.env.npm_execpath;
const npmCommand = process.platform === "win32" ? "npm.cmd" : "npm";
const attachments = ["输入数据包.zip", "reference.zip", "关键标准答案.xlsx", "任务规格转化.xlsx"];
const expectedReference = [
  "reports/asset_publish_matrix.csv",
  "reports/bandwidth_ladder.csv",
  "reports/segment_requeue.csv",
  "reports/variant_playlist_audit.csv",
  "src/audit_hls_publish.mjs"
].sort();
const reportKeys = {
  "reports/asset_publish_matrix.csv": ["asset_id"],
  "reports/bandwidth_ladder.csv": ["asset_id"],
  "reports/segment_requeue.csv": ["asset_id"],
  "reports/variant_playlist_audit.csv": ["asset_id", "profile"]
};

const sha256 = (bytes) => crypto.createHash("sha256").update(bytes).digest("hex");
const sha256File = (file) => sha256(fs.readFileSync(file));
const assert = (value, message) => { if (!value) throw new Error(message); };

function parseZipBytes(data) {
  const files = new Map();
  let offset = 0;
  while (offset + 30 <= data.length) {
    if (data.readUInt32LE(offset) !== 0x04034b50) break;
    const flags = data.readUInt16LE(offset + 6);
    const method = data.readUInt16LE(offset + 8);
    const compressedSize = data.readUInt32LE(offset + 18);
    const uncompressedSize = data.readUInt32LE(offset + 22);
    const nameLength = data.readUInt16LE(offset + 26);
    const extraLength = data.readUInt16LE(offset + 28);
    assert(!(flags & 0x08), "ZIP数据描述符不受支持");
    const name = data.subarray(offset + 30, offset + 30 + nameLength).toString("utf8").replaceAll("\\", "/");
    const start = offset + 30 + nameLength + extraLength;
    const compressed = data.subarray(start, start + compressedSize);
    if (!name.endsWith("/")) {
      const body = method === 0 ? compressed : method === 8 ? inflateRawSync(compressed) : null;
      assert(body && body.length === uncompressedSize, `无法解压${name}`);
      files.set(name, body);
    }
    offset = start + compressedSize;
  }
  return files;
}

const parseZip = (file) => parseZipBytes(fs.readFileSync(file));
async function extractZip(file, destination) {
  for (const [name, bytes] of parseZip(file)) {
    const target = path.resolve(destination, name);
    assert(target.startsWith(path.resolve(destination) + path.sep), `非法ZIP路径${name}`);
    await fsp.mkdir(path.dirname(target), { recursive: true });
    await fsp.writeFile(target, bytes);
  }
}

function workbookSheets(file) {
  const workbook = parseZipBytes(fs.readFileSync(file)).get("xl/workbook.xml")?.toString("utf8") ?? "";
  return [...workbook.matchAll(/<(?:[A-Za-z]+:)?sheet[^>]+name="([^"]+)"/gu)].map((match) => match[1]);
}

function workbookText(file) {
  const zip = parseZip(file);
  return [...zip]
    .filter(([name]) => name === "xl/sharedStrings.xml" || /^xl\/worksheets\/sheet\d+\.xml$/u.test(name))
    .map(([, bytes]) => bytes.toString("utf8"))
    .join("\n");
}

async function run(command, args, cwd) {
  const started = Date.now();
  return await new Promise((resolve) => {
    let child;
    try {
      child = spawn(command, args, { cwd, env: process.env, windowsHide: true });
    } catch (error) {
      resolve({ code: 1, stdout: "", stderr: error.stack ?? error.message, elapsed_ms: Date.now() - started });
      return;
    }
    let stdout = "";
    let stderr = "";
    let settled = false;
    child.stdout.on("data", (chunk) => { stdout += chunk; });
    child.stderr.on("data", (chunk) => { stderr += chunk; });
    child.on("error", (error) => {
      if (!settled) {
        settled = true;
        resolve({ code: 1, stdout, stderr: `${stderr}${error.stack ?? error.message}`, elapsed_ms: Date.now() - started });
      }
    });
    child.on("exit", (code) => {
      if (!settled) {
        settled = true;
        resolve({ code: code ?? 1, stdout, stderr, elapsed_ms: Date.now() - started });
      }
    });
  });
}

async function runNpm(args, cwd) {
  return npmCli ? await run(process.execPath, [npmCli, ...args], cwd) : await run(npmCommand, args, cwd);
}

function treeDigest(root, ignored = new Set()) {
  const lines = [];
  function visit(current, prefix = "") {
    for (const entry of fs.readdirSync(current, { withFileTypes: true }).toSorted((left, right) => left.name.localeCompare(right.name))) {
      const relative = prefix ? `${prefix}/${entry.name}` : entry.name;
      if (ignored.has(relative.split("/")[0])) continue;
      const full = path.join(current, entry.name);
      if (entry.isDirectory()) visit(full, relative);
      else lines.push(`${relative}\0${sha256File(full)}`);
    }
  }
  visit(root);
  return sha256(Buffer.from(lines.join("\n")));
}

function parseCsv(text) {
  const rows = [];
  let row = [];
  let cell = "";
  let quoted = false;
  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    if (quoted) {
      if (char === "\"" && text[index + 1] === "\"") {
        cell += "\"";
        index += 1;
      } else if (char === "\"") quoted = false;
      else cell += char;
    } else if (char === "\"") quoted = true;
    else if (char === ",") {
      row.push(cell);
      cell = "";
    } else if (char === "\n") {
      row.push(cell.replace(/\r$/u, ""));
      rows.push(row);
      row = [];
      cell = "";
    } else cell += char;
  }
  if (cell || row.length) {
    row.push(cell.replace(/\r$/u, ""));
    rows.push(row);
  }
  const headers = rows.shift() ?? [];
  return rows.filter((values) => values.some((value) => value !== ""))
    .map((values) => Object.fromEntries(headers.map((header, index) => [header, values[index] ?? ""])));
}

function normalizedRows(file, text) {
  const keys = reportKeys[file];
  return parseCsv(text).toSorted((left, right) => keys
    .map((key) => String(left[key]).localeCompare(String(right[key])))
    .find((value) => value !== 0) ?? 0);
}

function classifyExecutable(name, bytes) {
  const lower = name.toLowerCase();
  if (bytes.length >= 4 && bytes[0] === 0x7f && bytes.subarray(1, 4).toString("ascii") === "ELF") return "linux_elf";
  if (bytes.length >= 4 && [0xfeedface, 0xfeedfacf, 0xcafebabe].includes(bytes.readUInt32BE(0))) return "macos_macho";
  if (/\.(?:sh|bash|so)(?:\.|$)/u.test(lower)) return "posix_member";
  if (/^#!.*(?:ba|z|k)?sh/mu.test(bytes.subarray(0, 128).toString("utf8"))) return "posix_shebang";
  return null;
}

async function prepare(label, mutate) {
  const root = path.join(os.tmpdir(), label);
  await fsp.rm(root, { recursive: true, force: true });
  await fsp.mkdir(root, { recursive: true });
  await extractZip(path.join(artifactRoot, "输入数据包.zip"), root);
  const inputRoot = path.join(root, "input_data");
  const reference = parseZip(path.join(artifactRoot, "reference.zip"));
  await fsp.mkdir(path.join(inputRoot, "src"), { recursive: true });
  await fsp.writeFile(path.join(inputRoot, "src", "audit_hls_publish.mjs"), reference.get("src/audit_hls_publish.mjs"));
  if (mutate) await mutate(inputRoot);
  return { root, inputRoot, outputRoot: path.join(inputRoot, "output"), reference };
}

function outputPaths(root) {
  const paths = [];
  function walk(current, prefix = "") {
    for (const entry of fs.readdirSync(current, { withFileTypes: true })) {
      const relative = prefix ? `${prefix}/${entry.name}` : entry.name;
      if (entry.isDirectory()) walk(path.join(current, entry.name), relative);
      else paths.push(relative);
    }
  }
  walk(root);
  return paths.sort();
}

function compareReference(outputRoot, reference) {
  assert(JSON.stringify(outputPaths(outputRoot)) === JSON.stringify(expectedReference), "输出成员与Reference不一致");
  const semantic = crypto.createHash("sha256");
  for (const file of expectedReference) {
    const actual = fs.readFileSync(path.join(outputRoot, file));
    const expected = reference.get(file);
    if (file.endsWith(".csv")) {
      const actualRows = normalizedRows(file, actual.toString("utf8"));
      const expectedRows = normalizedRows(file, expected.toString("utf8"));
      assert(JSON.stringify(actualRows) === JSON.stringify(expectedRows), `${file}与Reference不一致`);
      semantic.update(JSON.stringify(actualRows));
    } else {
      assert(actual.toString("utf8").replaceAll("\r\n", "\n") === expected.toString("utf8").replaceAll("\r\n", "\n"), `${file}与Reference不一致`);
      semantic.update(actual.toString("utf8").replaceAll("\r\n", "\n"));
    }
  }
  return semantic.digest("hex");
}

await fsp.rm(evidenceRoot, { recursive: true, force: true });
await fsp.mkdir(evidenceRoot, { recursive: true });
assert(process.platform === "win32" && process.env.GITHUB_ACTIONS === "true", "该验证器只接受GitHub托管Windows运行");

const attachmentSha256 = Object.fromEntries(attachments.map((name) => [name, sha256File(path.join(artifactRoot, name))]));
const inputMembers = parseZip(path.join(artifactRoot, "输入数据包.zip"));
const executableScan = [...inputMembers].map(([name, bytes]) => ({ name, classification: classifyExecutable(name, bytes) })).filter((item) => item.classification);
assert(executableScan.length === 0, `输入包含平台专用可执行成员：${JSON.stringify(executableScan)}`);
const referenceMembers = [...parseZip(path.join(artifactRoot, "reference.zip")).keys()].sort();
assert(JSON.stringify(referenceMembers) === JSON.stringify(expectedReference), "Reference成员错误");
assert(JSON.stringify(workbookSheets(path.join(artifactRoot, "关键标准答案.xlsx"))) === JSON.stringify(["交付物答案清单", "固定字段答案", "固定集合答案", "固定数值答案", "允许变体答案"]), "关键标准答案Sheet错误");
assert(JSON.stringify(workbookSheets(path.join(artifactRoot, "任务规格转化.xlsx"))) === JSON.stringify(["任务规格转化"]), "任务规格Sheet错误");
const workbookControlTerms = /reference_members|reference\.zip成员|可验证点|不适合作为评分点的内容/iu;
assert(!workbookControlTerms.test(workbookText(path.join(artifactRoot, "关键标准答案.xlsx"))), "关键标准答案残留制题控制语");
assert(!workbookControlTerms.test(workbookText(path.join(artifactRoot, "任务规格转化.xlsx"))), "任务规格残留制题控制语");
const solutionText = parseZip(path.join(artifactRoot, "reference.zip")).get("src/audit_hls_publish.mjs").toString("utf8");
assert(!/\bvid(?:100|200|300|400)\b|https?:\/\/|node:net|node:http|fetch\s*\(/u.test(solutionText), "完成版模块含样本ID硬编码或外部网络调用");

const cleanRuns = [];
for (const label of ["Q10132 第一次 空目录", "Q10132 第二次 中文 空格目录"]) {
  const prepared = await prepare(label);
  const before = treeDigest(prepared.inputRoot, new Set(["output"]));
  const result = await runNpm(["run", "process:hls"], prepared.inputRoot);
  assert(result.code === 0, `${label}执行失败\n${result.stdout}\n${result.stderr}`);
  const after = treeDigest(prepared.inputRoot, new Set(["output"]));
  assert(before === after, `${label}修改了输入`);
  const semantic = compareReference(prepared.outputRoot, prepared.reference);
  cleanRuns.push({ directory_label: label, exit_code: result.code, input_digest_before: before, input_digest_after: after, semantic_digest: semantic, elapsed_ms: result.elapsed_ms });
}
assert(cleanRuns[0].semantic_digest === cleanRuns[1].semantic_digest, "两次结构化结果不一致");

const crlf = await prepare("Q10132 CRLF 输入", async (inputRoot) => {
  const file = path.join(inputRoot, "data", "encode_jobs.csv");
  const text = await fsp.readFile(file, "utf8");
  await fsp.writeFile(file, text.replace(/\r?\n/gu, "\r\n"));
});
let result = await runNpm(["run", "process:hls"], crlf.inputRoot);
assert(result.code === 0, `CRLF输入执行失败\n${result.stdout}\n${result.stderr}`);
const crlfDigest = compareReference(crlf.outputRoot, crlf.reference);
assert(crlfDigest === cleanRuns[0].semantic_digest, "CRLF输入改变业务结果");

const mutation = await prepare("Q10132 原因顺序变化", async (inputRoot) => {
  const file = path.join(inputRoot, "rules", "hls_publish_policy.json");
  const value = JSON.parse(await fsp.readFile(file, "utf8"));
  const left = value.requeue_reason_order.indexOf("bandwidth_out_of_range");
  const right = value.requeue_reason_order.indexOf("ladder_not_increasing");
  assert(left >= 0 && right >= 0, "原因顺序缺少目标项");
  value.requeue_reason_order.splice(left, 1);
  value.requeue_reason_order.splice(right, 0, "bandwidth_out_of_range");
  await fsp.writeFile(file, `${JSON.stringify(value, null, 2)}\n`);
});
result = await runNpm(["run", "process:hls"], mutation.inputRoot);
assert(result.code === 0, `原因顺序变化执行失败\n${result.stdout}\n${result.stderr}`);
const mutatedAssets = normalizedRows("reports/asset_publish_matrix.csv", fs.readFileSync(path.join(mutation.outputRoot, "reports", "asset_publish_matrix.csv"), "utf8"));
const mutatedRequeues = normalizedRows("reports/segment_requeue.csv", fs.readFileSync(path.join(mutation.outputRoot, "reports", "segment_requeue.csv"), "utf8"));
const changedAsset = mutatedAssets.find((row) => row.asset_id === "vid300");
const changedRequeue = mutatedRequeues.find((row) => row.asset_id === "vid300");
assert(changedAsset?.primary_reason === "bandwidth_out_of_range", "原因顺序变化没有联动vid300主原因");
assert(changedRequeue?.reason === "bandwidth_out_of_range" && changedRequeue?.action === "review_bandwidth_policy", "原因顺序变化没有联动vid300动作");
const baselineAssets = normalizedRows("reports/asset_publish_matrix.csv", mutation.reference.get("reports/asset_publish_matrix.csv").toString("utf8"));
assert(JSON.stringify(mutatedAssets.filter((row) => row.asset_id !== "vid300")) === JSON.stringify(baselineAssets.filter((row) => row.asset_id !== "vid300")), "原因顺序变化影响了无关资产");

const negative = await prepare("Q10132 无效输入", async (inputRoot) => {
  await fsp.rm(path.join(inputRoot, "rules", "hls_publish_policy.json"));
});
result = await runNpm(["run", "process:hls"], negative.inputRoot);
const deliverablesAbsent = !fs.existsSync(negative.outputRoot) || fs.readdirSync(negative.outputRoot).length === 0;
assert(result.code !== 0 && deliverablesAbsent, "无效输入没有失败关闭");

const evidence = {
  schema_version: 1,
  task_asset_id: "node_hls_publish_disposition",
  result: "PASS",
  generated_at_utc: new Date().toISOString(),
  git_commit_sha: process.env.GITHUB_SHA,
  workflow_run_id: process.env.GITHUB_RUN_ID,
  runner: {
    os: process.env.RUNNER_OS,
    arch: process.env.RUNNER_ARCH,
    image_os: process.env.ImageOS,
    image_version: process.env.ImageVersion,
    node: process.version,
    powershell_hosted_workflow: true
  },
  software: { node: process.version, executed: true },
  attachment_sha256: attachmentSha256,
  workbook_checks: {
    answer_sheet_names: workbookSheets(path.join(artifactRoot, "关键标准答案.xlsx")),
    specification_sheet_names: ["任务规格转化"]
  },
  platform_audit: {
    linux_executables: executableScan,
    linux_executables_executed: false,
    no_wsl_required: true,
    no_linux_container_required: true,
    no_posix_shell_required: true,
    no_unix_only_api_required: true,
    cross_platform_paths: true
  },
  clean_runs: cleanRuns,
  crlf_input: { file: "data/encode_jobs.csv", exit_code: 0, semantic_digest: crlfDigest, reference_match: true },
  positive_mutation: {
    changed_rule: "bandwidth_out_of_range移动到ladder_not_increasing之前",
    exit_code: 0,
    affected_asset: "vid300",
    primary_reason: changedAsset.primary_reason,
    action: changedRequeue.action,
    unrelated_assets_unchanged: true
  },
  invalid_input: {
    removed_input: "rules/hls_publish_policy.json",
    exit_code: result.code,
    deliverables_absent: deliverablesAbsent
  },
  network: {
    installation_network_access: "Node.js安装阶段",
    formal_run_network_access: "none, local files and local Node.js only"
  }
};
await fsp.writeFile(path.join(evidenceRoot, "windows-verification.json"), `${JSON.stringify(evidence, null, 2)}\n`);
console.log(JSON.stringify(evidence, null, 2));
