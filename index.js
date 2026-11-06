/**
 * dsh-onenote 0.1.1 — DSH/Cordis host plugin registering OneNote desktop tools.
 *
 * The plugin talks to the local Microsoft OneNote application through its COM
 * automation interface (via Windows PowerShell 5.1), so it works against the
 * real notebooks on this machine instead of a cloud Graph copy.
 */
import { OneNoteBridge } from "./lib/bridge.js";
import { makeToolHandlers, renderResult } from "./lib/tools.js";

export const name = "dsh-onenote";
export const version = "0.5.2";
export const inject = ["tools"];

export const DEFAULT_CONFIG = Object.freeze({
  powershellPath: "powershell.exe",
  timeoutMs: 180000,
  extraSections: [],
  pageColor: "",
  pageSize: null,
  // code blocks: each source line gets this prefix (the box-drawing bar keeps
  // read-back working). "" removes the bar, at the cost of losing the code type
  // when the page is read back.
  codePrefix: "│ ",
  // 0 = leave comments where they are; >0 aligns trailing // comments of one
  // fenced block to that column (bytes) or, when 0, to longest code + 2.
  alignCommentColumn: 0,
  // Upper bound for the automatic column (keeps very long lines sane).
  alignCommentMax: 80,
});

export const TOOL_NAMES = Object.freeze([
  "onenote_list",
  "onenote_read",
  "onenote_search",
  "onenote_create",
  "onenote_update",
  "onenote_append",
  "onenote_edit",
  "onenote_publish",
  "onenote_delete",
  "onenote_create_section",
  "onenote_rename_section",
]);

function str(description, extra = {}) {
  return { type: "string", description, ...extra };
}
function int(description, extra = {}) {
  return { type: "integer", description, ...extra };
}
function bool(description) {
  return { type: "boolean", description };
}

const PAGE_LOCATOR = {
  pageId: str("OneNote page ID (from onenote_list / onenote_search). Highest priority."),
  pageTitle: str("Page title to resolve when pageId is not given (exact match first, then substring)."),
  sectionId: str("Section ID used together with pageTitle."),
  sectionName: str("Section name used together with pageTitle (e.g. 快速笔记)."),
};

export const TOOL_PARAMETERS = Object.freeze({
  onenote_list: {
    notebook: str("Filter by notebook name or ID (substring match)."),
    section: str("Filter by section name or ID (substring match)."),
    includePages: bool("Include page lists (default true). Set false for a compact notebook/section map."),
    includeRecycleBin: bool("Include recycle-bin sections and deleted pages (default false)."),
    maxSections: int("Cap the number of sections returned (0 = all, default)."),
    maxPages: int("Cap the pages returned per section (0 = all, default)."),
  },
  onenote_read: { ...PAGE_LOCATOR, maxBlocks: int("Cap the number of content blocks returned (0 = all, default).") },
  onenote_search: {
    query: str("Text to look for. Page titles are matched first (case-insensitive substring)."),
    scopeId: str("Optional notebook/section/group ID limiting the search scope."),
    limit: int("Max results (default 20)."),
    includeContent: bool("Also scan page text, not just titles (slower; default false)."),
    maxScan: int("Max pages whose text is read when includeContent=true (default 200)."),
    includeRecycleBin: bool("Include pages that live in a recycle-bin section (default false)."),
  },
  onenote_create: {
    sectionId: str("Target section ID (preferred)."),
    sectionName: str("Target section name, used when sectionId is absent."),
    notebook: str("Notebook name used to disambiguate sectionName."),
    title: str("Page title. Optional when markdown starts with '# Title'."),
    markdown: str("Page body in the supported markdown subset."),
  },
  onenote_update: {
    ...PAGE_LOCATOR,
    title: str("New page title. Omit to keep the current title."),
    markdown: str("New page body; replaces all existing text outlines on the page."),
  },
  onenote_append: {
    ...PAGE_LOCATOR,
    title: str("Optional new page title."),
    markdown: str("Content appended as a new outline at the end of the page."),
  },
  onenote_edit: {
    ...PAGE_LOCATOR,
    mode: str("replaceText | replaceBlock | insertAfter | deleteBlock."),
    objectId: str("Target block objectId from onenote_read(withIds=true)."),
    find: str("Target by unique text instead of objectId (fails when ambiguous)."),
    search: str("replaceText: exact substring to find in that block."),
    replacement: str("replaceText: replacement text (default empty string)."),
    all: bool("replaceText: replace every occurrence instead of the first."),
    text: str("replaceBlock/insertAfter: the new paragraph text (may be empty to clear)."),
    kind: str("insertAfter: para | h2 | h3 | bullet | number (default para)."),
  },
  onenote_publish: {
    ...PAGE_LOCATOR,
    path: str("Absolute target file path. Existing files are replaced unless overwrite=false."),
    format: str("one | onepkg | mht | pdf | xps | docx | emf | html (default one)."),
    overwrite: bool("Replace an existing target file (default true)."),
  },
  onenote_delete: {
    ...PAGE_LOCATOR,
    permanent: bool("Skip the recycle bin and delete permanently (default false)."),
  },
  onenote_create_section: {
    path: str("Absolute path of the .one file to create or open, e.g. \"<notebooks>\\\\topic.one\"."),
    keepBlankPage: bool("Keep the empty page OneNote adds to a new section (default false: it is removed)."),
  },
  onenote_rename_section: {
    sectionId: str("Section ID to rename (from onenote_list)."),
    path: str("Standalone .one path, used when sectionId is absent."),
    name: str("New section name."),
  },
});

export const TOOL_REQUIRED = Object.freeze({
  onenote_list: [],
  onenote_read: [],
  onenote_search: ["query"],
  onenote_create: ["markdown"],
  onenote_update: ["markdown"],
  onenote_append: ["markdown"],
  onenote_edit: ["mode"],
  onenote_publish: ["path"],
  onenote_delete: [],
  onenote_create_section: ["path"],
  onenote_rename_section: ["name"],
});

export const TOOL_DESCRIPTIONS = Object.freeze({
  onenote_list:
    "List OneNote notebooks, sections and pages (local desktop OneNote via COM). Includes configured local sections that are not part of an open notebook.",
  onenote_read:
    "Read a OneNote page and return it as markdown (headings, lists, tables, code lines). Resolve the page by pageId, or by section + pageTitle.",
  onenote_search:
    "Search OneNote pages by title (case-insensitive substring). Set includeContent=true to also scan page text; that reads up to maxScan pages and is slower. No Windows Search index required.",
  onenote_create:
    "Create a new OneNote page from markdown in an existing section. Supports a markdown subset: # title, ##/### headings (written at body size +4 and +2 with paragraph spacing), bullets, numbered lists, tables, fenced code, quotes. Arbitrary fonts and colours are not applied (the OneNote XML style import rejects them).",
  onenote_update:
    "Replace the text content of an existing OneNote page (keeps the page ID). Embedded images/ink inside outlines are removed with the outlines; page-level images are kept.",
  onenote_append:
    "Append a new outline with markdown content to the end of an existing OneNote page, leaving existing content untouched.",
  onenote_edit:
    "Edit ONE block of a page in place and leave everything else untouched. Modes: replaceText (fix text inside a block), replaceBlock (rewrite a block), insertAfter (add a paragraph after a block), deleteBlock. Get the objectId from onenote_read first. Only the outline that owns the block is re-sent, so other blocks keep their content and formatting.",
  onenote_publish:
    "Export a OneNote page to a file: .one, .onepkg, .pdf, .docx, .xps, .mht, .html or .emf. The target file must be absolute.",
  onenote_delete:
    "Delete a OneNote page (moved to the recycle bin unless permanent=true).",
  onenote_create_section:
    "Create a standalone OneNote section (.one file) at an absolute path, or open it when it already exists. If the new section comes with an empty placeholder page it is removed unless keepBlankPage=true (OneNote does not always add one).",
  onenote_rename_section:
    "Rename a section. Notebook sections can be renamed; standalone .one sections are read-only in OneNote and fail with SECTION_READ_ONLY.",
});

export function buildParameters(toolName) {
  const required = TOOL_REQUIRED[toolName] ?? [];
  const params = {};
  for (const [paramName, schema] of Object.entries(TOOL_PARAMETERS[toolName] ?? {})) {
    params[paramName] = required.includes(paramName) ? { ...schema, required: true } : { ...schema };
  }
  return params;
}

export const OUTPUT = Object.freeze({
  schema: { type: "object", additionalProperties: true },
  render: renderResult,
});

function pickConfig(config = {}) {
  const merged = { ...DEFAULT_CONFIG, ...(config && typeof config === "object" ? config : {}) };
  if (typeof merged.powershellPath !== "string" || merged.powershellPath.trim() === "") {
    merged.powershellPath = DEFAULT_CONFIG.powershellPath;
  }
  if (!Number.isFinite(merged.timeoutMs) || merged.timeoutMs <= 0) merged.timeoutMs = DEFAULT_CONFIG.timeoutMs;
  if (typeof merged.codePrefix !== "string") merged.codePrefix = DEFAULT_CONFIG.codePrefix;
  // "auto" = longest code + 2 (capped by alignCommentMax); 0 = off; N > 0 = fixed.
  const column = merged.alignCommentColumn;
  merged.alignCommentColumn =
    column === "auto" || column === true
      ? "auto"
      : Number.isFinite(column) && column > 0
        ? column
        : 0;
  if (!Number.isFinite(merged.alignCommentMax) || merged.alignCommentMax <= 0) {
    merged.alignCommentMax = DEFAULT_CONFIG.alignCommentMax;
  }
  if (!Array.isArray(merged.extraSections)) merged.extraSections = [];
  merged.extraSections = merged.extraSections
    .filter((entry) => entry && (typeof entry.path === "string" || typeof entry.id === "string"))
    .map((entry) => {
      const path = typeof entry.path === "string" ? entry.path.trim() : "";
      const id = typeof entry.id === "string" ? entry.id.trim() : "";
      return {
        name: typeof entry.name === "string" && entry.name.trim() !== "" ? entry.name : path || id,
        path,
        id,
      };
    })
    .filter((entry) => entry.path !== "" || entry.id !== "");
  // Page background: "" inherits OneNote's default, otherwise #RRGGBB /
  // "automatic" / "none". Anything else is ignored rather than breaking writes.
  const pageColor = typeof merged.pageColor === "string" ? merged.pageColor.trim() : "";
  merged.pageColor = /^#[0-9a-f]{6}$/i.test(pageColor) || pageColor === "automatic" || pageColor === "none" ? pageColor : "";
  merged.pageSize = normalizePageSize(merged.pageSize);
  return merged;
}

/**
 * Page size in OneNote points (1/72 in). OneNote only accepts either
 * <Automatic/> or Orientation+Dimensions+Margins, so width/height are required.
 *   { width: 1190, height: 842, landscape: true, margins: { top, bottom, left, right } }
 * The string "automatic" forces OneNote's Automatic size; null leaves the page
 * size untouched.
 */
export function normalizePageSize(value) {
  if (typeof value === "string") return value.trim().toLowerCase() === "automatic" ? "automatic" : null;
  if (!value || typeof value !== "object") return null;
  const width = Number(value.width);
  const height = Number(value.height);
  if (!Number.isFinite(width) || !Number.isFinite(height) || width <= 0 || height <= 0) return null;
  const out = { width, height, landscape: Boolean(value.landscape) };
  if (value.margins && typeof value.margins === "object") {
    const margins = {};
    for (const key of ["top", "bottom", "left", "right"]) {
      const n = Number(value.margins[key]);
      if (Number.isFinite(n) && n >= 0) margins[key] = n;
    }
    if (Object.keys(margins).length > 0) out.margins = margins;
  }
  return out;
}

/** Wire the plugin against an injected `defineTool` and a Cordis-like `ctx`. */
export function applyWith({ defineTool, ctx, config = {}, bridge } = {}) {
  if (typeof defineTool !== "function") throw new Error("applyWith: defineTool is required");
  if (!ctx || !ctx.tools || typeof ctx.tools.register !== "function") {
    throw new Error("applyWith: ctx.tools.register is required");
  }
  const cfg = pickConfig(config);
  const instance = bridge ?? new OneNoteBridge(cfg);
  const handlers = makeToolHandlers({
    bridge: instance,
    renderOptions: {
      codePrefix: cfg.codePrefix,
      alignCommentColumn: cfg.alignCommentColumn,
      alignCommentMax: cfg.alignCommentMax,
    },
  });
  const registered = [];
  for (const toolName of TOOL_NAMES) {
    const def = defineTool({
      name: toolName,
      description: TOOL_DESCRIPTIONS[toolName],
      parameters: buildParameters(toolName),
      output: OUTPUT,
      execute: (args = {}) => handlers[toolName](args),
    });
    ctx.tools.register(def);
    registered.push(toolName);
  }
  return { config: cfg, tools: registered };
}

/**
 * Real Cordis entry point. Cordis treats a plugin's return value as an effect,
 * so this must stay `undefined`.
 */
export async function apply(ctx, config = {}) {
  const dshTools = await import("@deepseek-ai/dsh-tools");
  applyWith({ defineTool: dshTools.defineTool, ctx, config });
}

export { OneNoteBridge };
