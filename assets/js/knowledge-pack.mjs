/** 启动口令与 Markdown 交付：单篇原文；合集／历史打包带用法与目录壳。 */
import { businessCatalog, packOrderedFiles, resourcePaths } from "./catalog.mjs";

/** 顶栏短口令：克隆仓库并加载；仓名／自称跟 manifest.packPrefix 与 name。 */
export function localPackPrompt(manifest = {}) {
  const prefix = typeof manifest.packPrefix === "string" && manifest.packPrefix.trim()
    ? manifest.packPrefix.trim()
    : "bro";
  const name = typeof manifest.name === "string" && manifest.name.trim()
    ? manifest.name.trim()
    : "小弟";
  return `克隆 https://github.com/luckymaomi/${prefix} ，并加载${name}`;
}

const marker = file => `===== ${file.title}（${file.path}） =====`;

function usageBlock() {
  return [
    "## 怎么用",
    "",
    "1. 克隆本仓库后加载小弟（或分身启动、加载分身，说法不挑；旧口令宝宝／宝贝等一律按小弟处理）。",
    "2. 总启时本轮实读 AGENTS、分身与 manifest，再按分身启动报法报名；点名某技能只读该篇正文。",
    "3. 已有具体任务直接办，不必先做完整总启。",
    "4. 配套模板、脚本见文末仓库相对路径；原始业务表由主人另附。",
    "",
  ];
}

async function buildMarkdown(manifest, files, loadRaw, stamp, title) {
  const paths = resourcePaths(files);
  const parts = [
    `# ${title}`,
    "",
    `生成日期：${stamp}`,
    "",
    ...usageBlock(),
    "## 技能目录",
    "",
    businessCatalog(manifest),
    "",
    "## 本文件收录",
    "",
    ...files.map((file, index) => `${index + 1}. 〔${file.group}〕${file.title} —— ${marker(file)}`),
    "",
    "## 配套资源",
    "",
  ];
  if (paths.length) {
    for (const path of paths) parts.push(`- ${path}`);
  } else {
    parts.push("本文件收录篇未另列配套资源。");
  }
  parts.push("");
  for (const file of files) {
    parts.push(marker(file), "", String(await loadRaw(file.path)).trimEnd(), "");
  }
  return parts.join("\n").trimEnd() + "\n";
}

/** 单篇下载：只下该篇原文。 */
export async function buildSkillMarkdown(file, loadRaw) {
  return String(await loadRaw(file.path)).trimEnd() + "\n";
}

export function buildPackMarkdown(manifest, loadRaw, stamp) {
  return buildMarkdown(manifest, packOrderedFiles(manifest), loadRaw, stamp, `${manifest.name}合集`);
}

/** 历史过程稿打包：manifest.history 列出的 Markdown 合并为一个文件。 */
export async function buildHistoryMarkdown(manifest, loadRaw, stamp) {
  const entries = Array.isArray(manifest.history) ? manifest.history : [];
  if (!entries.length) {
    throw new Error("清单未配置 history，无法打包历史。");
  }
  const parts = [
    `# ${manifest.name || "小弟"} · 历史打包`,
    "",
    `生成日期：${stamp}`,
    "",
    "## 怎么用",
    "",
    "1. 本文件是 history/ 过程底稿合集，**不是**现行技能正文。",
    "2. 现行口径以仓库 core/ 各技能与 AGENTS 为准；冲突时听技能页。",
    "3. 技术实力相关精华见「飞行部数据挖掘_清洗与分析」；挖数动手见「南货航飞行部潜规则」。",
    "",
    "## 本文件收录",
    "",
    ...entries.map((entry, index) => {
      const title = entry.title || entry.path;
      return `${index + 1}. ${title} —— ${entry.path}`;
    }),
    "",
  ];
  for (const entry of entries) {
    const title = entry.title || entry.path;
    const section = `===== ${title}（${entry.path}） =====`;
    parts.push(section, "", String(await loadRaw(entry.path)).trimEnd(), "");
  }
  return parts.join("\n").trimEnd() + "\n";
}
