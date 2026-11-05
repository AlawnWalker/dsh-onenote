/**
 * Live smoke test — talks to the real OneNote application.
 *
 * NOT part of `node --test` (it lives outside test/ on purpose). Run it
 * explicitly when OneNote is available:
 *
 *   node dev/live-smoke.mjs [快速笔记 .one 路径]
 *
 * It resolves the section by file path (OneNote IDs are unstable across
 * restarts), creates one page, exercises read / update / append / publish,
 * verifies the exported .one header, and deletes the page again. Nothing else
 * in your notebooks is touched.
 */
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import os from "node:os";
import path from "node:path";

import { OneNoteBridge } from "../lib/bridge.js";
import { blocksToMarkdown, markdownToBlocks } from "../lib/markdown.js";

// No absolute paths in the repo: pass the section .one path (argv[2] or env).
const SECTION_PATH = process.argv[2] ?? process.env.ONENOTE_SMOKE_SECTION ?? "";
const SECTION_NAME = process.env.ONENOTE_SMOKE_SECTION_NAME ?? "快速笔记";
if (SECTION_PATH === "") {
  console.error("usage: node dev/live-smoke.mjs <path-to-section.one>   (or set ONENOTE_SMOKE_SECTION)");
  process.exit(2);
}
const TITLE = `DSH-LIVE-SMOKE-${new Date().toISOString().replace(/[:.]/g, "-")}`;
const ONE_MAGIC = Buffer.from([0xe4, 0x52, 0x5c, 0x7b, 0x8c, 0xd8, 0xa7, 0x4d]);

const bridge = new OneNoteBridge({ extraSections: [{ name: SECTION_NAME, path: SECTION_PATH }] });
const work = path.join(os.tmpdir(), `dsh-onenote-smoke-${process.pid}`);
const checks = [];

function check(name, fn) {
  return Promise.resolve()
    .then(fn)
    .then(() => {
      checks.push({ name, ok: true });
      console.log(`PASS  ${name}`);
    })
    .catch((err) => {
      checks.push({ name, ok: false, error: err.message });
      console.log(`FAIL  ${name}: ${err.message}`);
      throw err;
    });
}

const body = markdownToBlocks(`冒烟测试正文。

## 小节标题

- 条目甲
- 条目乙
  - 子条目

1. 第一步
2. 第二步

| 列1 | 列2 |
| --- | --- |
| x | y |

\`\`\`
line one
line two
\`\`\`
`);

let pageId = null;
try {
  const listing = await bridge.invoke("list", { section: SECTION_NAME, includePages: false });
  await check("onenote_list finds the quick-notes section by path", () => {
    const local = listing.localSections.find((section) => section.name === SECTION_NAME);
    assert.ok(local, "configured local section is missing from onenote_list");
    assert.equal(local.error, null);
    assert.ok(local.id && local.id.startsWith("{"), "section id was not resolved");
  });

  const created = await bridge.invoke("create", { sectionName: SECTION_NAME, title: TITLE, blocks: body });
  pageId = created.pageId;
  await check("onenote_create returns a page id", () => assert.ok(pageId && pageId.startsWith("{")));

  const readBack = await bridge.invoke("read", { pageId });
  await check("onenote_read returns the title", () => assert.equal(readBack.title, TITLE));
  await check("onenote_read preserves the text content", () => {
    const markdown = blocksToMarkdown(readBack.blocks);
    assert.match(markdown, /冒烟测试正文/);
    assert.match(markdown, /## 小节标题/);
    assert.match(markdown, /- 条目甲/);
    assert.match(markdown, /  - 子条目/);
    assert.match(markdown, /1\. 第一步/);
    assert.match(markdown, /\| 列1 \| 列2 \|/);
    assert.match(markdown, /line one/);
  });

  await bridge.invoke("update", {
    pageId,
    title: `${TITLE} (updated)`,
    blocks: markdownToBlocks("更新后的正文\n\n- 只剩一项\n"),
  });
  const afterUpdate = await bridge.invoke("read", { pageId });
  await check("onenote_update replaces the body", () => {
    const markdown = blocksToMarkdown(afterUpdate.blocks);
    assert.match(markdown, /更新后的正文/);
    assert.doesNotMatch(markdown, /冒烟测试正文/);
    assert.doesNotMatch(markdown, /小节标题/);
    assert.equal(afterUpdate.title, `${TITLE} (updated)`);
  });

  await bridge.invoke("append", { pageId, blocks: markdownToBlocks("追加的一行\n") });
  const afterAppend = await bridge.invoke("read", { pageId });
  await check("onenote_append keeps the old content and adds the new one", () => {
    const markdown = blocksToMarkdown(afterAppend.blocks);
    assert.match(markdown, /更新后的正文/);
    assert.match(markdown, /追加的一行/);
  });

  const target = path.join(work, "smoke.one");
  const published = await bridge.invoke("publish", { pageId, path: target, format: "one" });
  await check("onenote_publish writes a valid .one file", () => {
    assert.ok(published.bytes > 0, "exported file is empty");
    const head = readFileSync(target).subarray(0, 8);
    assert.ok(head.equals(ONE_MAGIC), `unexpected .one header: ${head.toString("hex")}`);
  });

  const pdfTarget = path.join(work, "smoke.pdf");
  await bridge.invoke("publish", { pageId, path: pdfTarget, format: "pdf" });
  await check("onenote_publish writes a PDF", () => {
    const head = readFileSync(pdfTarget).subarray(0, 5).toString("latin1");
    assert.equal(head, "%PDF-");
  });
} finally {
  if (pageId) {
    try {
      await bridge.invoke("delete", { pageId, permanent: true });
      console.log("PASS  onenote_delete removed the test page");
    } catch (err) {
      console.log(`WARN  cleanup failed for ${pageId}: ${err.message}`);
    }
  }
  try {
    const { rm } = await import("node:fs/promises");
    await rm(work, { recursive: true, force: true });
  } catch {
    /* ignore */
  }
}

const failed = checks.filter((entry) => !entry.ok);
console.log(`\n${checks.length - failed.length}/${checks.length} checks passed`);
if (failed.length > 0) process.exit(1);
