import assert from "node:assert/strict";
import test from "node:test";

import { OneNoteBridge } from "../lib/bridge.js";

test("bridge normalises per-level bullet glyphs", () => {
  const bridge = new OneNoteBridge({ bulletValues: [20, "2", -1, "x"] });
  assert.deepEqual(bridge.bulletValues, [20, 2]);
});

test("bridge falls back to a single bullet when no list is given", () => {
  const bridge = new OneNoteBridge({ bulletValue: 6 });
  assert.deepEqual(bridge.bulletValues, []);
  assert.equal(bridge.bulletValue, 6);
});

test("bridge accepts bullet indexes above 9", () => {
  const bridge = new OneNoteBridge({ bulletValue: 20 });
  assert.equal(bridge.bulletValue, 20);
});
