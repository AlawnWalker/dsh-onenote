import assert from "node:assert/strict";
import test from "node:test";

import { blocksToMarkdown, markdownToBlocks } from "../lib/markdown.js";

/**
 * Emulate what the bridge does to a block before OneNote stores it:
 * numbered items become bullets whose text carries the number, which is the
 * workaround for OneNote refusing `one:Number` imports on Office 2024.
 */
function throughOneNote(blocks) {
  return blocks.map((block) => {
    if (block.type === "number") {
      return { ...block, type: "bullet", text: `${block.ordinal ?? 1}. ${block.text}` };
    }
    return block;
  });
}

test("headings become typed blocks and round-trip", () => {
  const markdown = "# 页面标题\n\n正文一段。\n\n## 二级标题\n\n### 三级标题\n";
  const blocks = markdownToBlocks(markdown);
  assert.deepEqual(
    blocks.map((block) => block.type),
    ["title", "para", "h2", "h3"],
  );
  assert.equal(blocks[2].text, "二级标题");
  assert.equal(blocks[3].text, "三级标题");
  assert.equal(blocksToMarkdown(blocks), markdown);
});

test("legacy symbol headings still read back as markdown", () => {
  const blocks = [
    { type: "para", text: "━━━ 旧二级 ━━━", level: 0 },
    { type: "para", text: "▌ 旧三级", level: 0 },
  ];
  assert.equal(blocksToMarkdown(blocks), "## 旧二级\n\n### 旧三级\n");
});

test("nested bullet lists keep their levels", () => {
  const markdown = "- 一\n- 二\n  - 二点一\n    - 二点一之甲\n- 三\n";
  const blocks = markdownToBlocks(markdown);
  assert.deepEqual(
    blocks.map((block) => [block.type, block.level]),
    [
      ["bullet", 0],
      ["bullet", 0],
      ["bullet", 1],
      ["bullet", 2],
      ["bullet", 0],
    ],
  );
  assert.equal(blocksToMarkdown(blocks), markdown);
});

test("numbered lists survive the bullet workaround", () => {
  const markdown = "1. 第一步\n2. 第二步\n";
  const blocks = markdownToBlocks(markdown);
  assert.deepEqual(
    blocks.map((block) => [block.type, block.ordinal]),
    [
      ["number", 1],
      ["number", 2],
    ],
  );
  assert.equal(blocksToMarkdown(throughOneNote(blocks)), markdown);
});

test("tables render as a markdown table with separators", () => {
  const markdown = "| 列A | 列B |\n| --- | --- |\n| a1 | b1 |\n| a2 | b2 |\n";
  const blocks = markdownToBlocks(markdown);
  assert.equal(blocks.length, 1);
  assert.deepEqual(blocks[0].rows, [
    ["列A", "列B"],
    ["a1", "b1"],
    ["a2", "b2"],
  ]);
  const back = blocksToMarkdown([
    { type: "table", rows: [["列A", "列B"], ["a1", "b1"], ["a2", "b2"]], level: 0 },
  ]);
  assert.equal(back, markdown);
});

test("fenced code becomes gutter lines and back", () => {
  const markdown = "```js\nconst x = 1;\nconsole.log(x);\n```\n";
  const blocks = markdownToBlocks(markdown);
  assert.deepEqual(
    blocks.map((block) => [block.type, block.text]),
    [
      ["code", "│ const x = 1;"],
      ["code", "│ console.log(x);"],
    ],
  );
  // OneNote has nowhere to store the fence language, so it comes back bare.
  assert.equal(blocksToMarkdown(blocks), "```\nconst x = 1;\nconsole.log(x);\n```\n");
});

test("blank lines separate lists from the blocks that follow", () => {
  const blocks = [
    { type: "bullet", text: "一", level: 0 },
    { type: "bullet", text: "二", level: 0 },
    { type: "para", text: "━━━ 标题 ━━━", level: 0 },
    { type: "table", rows: [["a", "b"], ["1", "2"]], level: 0 },
    { type: "para", text: "正文", level: 0 },
  ];
  const markdown = blocksToMarkdown(blocks);
  assert.equal(markdown, "- 一\n- 二\n\n## 标题\n\n| a | b |\n| --- | --- |\n| 1 | 2 |\n\n正文\n");
});

test("quotes and paragraph breaks are preserved", () => {
  const markdown = "> 一句话\n\n普通段落。\n";
  const blocks = markdownToBlocks(markdown);
  assert.equal(blocks[0].text, "▎ 一句话");
  assert.equal(blocksToMarkdown(blocks), markdown);
});

test("pipes inside table cells are escaped on the way out", () => {
  const markdown = blocksToMarkdown([{ type: "table", rows: [["a|b", "c"]], level: 0 }]);
  assert.match(markdown, /a\\\|b/);
});
