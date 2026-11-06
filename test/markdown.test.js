import assert from "node:assert/strict";
import test from "node:test";

import { alignCodeComments, blocksToMarkdown, findInlineComment, markdownToBlocks } from "../lib/markdown.js";

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

test("indented tree lines become separate paragraphs with EM SPACE indents", () => {
  const blocks = markdownToBlocks(
    "kernel_init_freeable()\n  ├─ do_pre_smp_initcalls()\n  └─ do_basic_setup()\n        │    note\n",
  );
  assert.deepEqual(blocks.map((b) => b.type), ["para", "para", "para", "para"]);
  assert.equal(blocks[0].text, "kernel_init_freeable()");
  assert.equal(blocks[1].text, "　├─ do_pre_smp_initcalls()");
  assert.equal(blocks[2].text, "　└─ do_basic_setup()");
  assert.equal(blocks[3].text, "　　　　│    note");
});
test("findInlineComment skips comments, URLs and string contents", () => {
  assert.equal(findInlineComment("// whole line"), -1);
  assert.equal(findInlineComment('url = "http://x";'), -1);
  assert.equal(findInlineComment('s = "a // b";'), -1);
  assert.equal(findInlineComment("int a = 1; // note"), 11);
});

test("alignCodeComments aligns trailing comments to longest code + 2 by default", () => {
  const lines = ["int a = 1; // one", "long_call(a, b); // two", "// whole", "x;"];
  const aligned = alignCodeComments(lines, 0);
  const columns = aligned
    .map((line) => ({ line, index: findInlineComment(line) }))
    .filter((entry) => entry.index > 0)
    .map((entry) => entry.index);
  assert.equal(new Set(columns).size, 1, "all comments share one column");
  assert.equal(aligned[2], "// whole");
  assert.equal(aligned[3], "x;");
});

test("alignCodeComments honours an explicit column", () => {
  const aligned = alignCodeComments(["a; // one", "longer_name; // two"], 30);
  for (const line of aligned) {
    assert.equal(findInlineComment(line), 30);
  }
});

test("markdownToBlocks: alignCommentColumn works inside a fence only", () => {
  const markdown = ["```c", "int a = 1; // one", "long_call(a, b); // two", "```", "", "int z = 9; // outside", ""].join("\n");
  const blocks = markdownToBlocks(markdown, { alignCommentColumn: -1 });
  const codeLines = blocks.filter((block) => block.type === "code").map((block) => block.text);
  assert.equal(codeLines.length, 2);
  const columns = codeLines.map((line) => line.indexOf("//"));
  assert.equal(columns[0], columns[1]);
  const outside = blocks.find((block) => block.type === "para");
  assert.equal(outside.text, "int z = 9; // outside", "paragraphs are not aligned");
});

test("markdownToBlocks: codePrefix can be empty", () => {
  const blocks = markdownToBlocks(["```", "int a = 1;", "```", ""].join("\n"), { codePrefix: "" });
  assert.equal(blocks[0].type, "code");
  assert.equal(blocks[0].text, "int a = 1;");
});

test("markdownToBlocks: alignCommentColumn 0 leaves code untouched", () => {
  const markdown = ["```c", "int a = 1; // one", "long_call(a, b); // two", "```", ""].join("\n");
  const blocks = markdownToBlocks(markdown, { alignCommentColumn: 0 });
  const codeLines = blocks.filter((block) => block.type === "code").map((block) => block.text);
  assert.deepEqual(codeLines, ["│ int a = 1; // one", "│ long_call(a, b); // two"]);
});

test("markdownToBlocks: default output is unchanged", () => {
  const markdown = ["```c", "int a = 1; // one", "```", ""].join("\n");
  const plain = markdownToBlocks(markdown);
  const explicit = markdownToBlocks(markdown, { codePrefix: "│ ", alignCommentColumn: 0 });
  assert.deepEqual(explicit, plain);
});