import assert from "node:assert/strict";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { CONFIG_DEFAULTS, parseWorkspaceConfig, resolveFromConfig } from "./workspace-config.js";

const CONFIG = "/ws/pptx-gen.config.yml";

test("an empty config is a valid, fully-default workspace", () => {
  const config = parseWorkspaceConfig("", CONFIG);
  assert.equal(config.version, 1);
  assert.equal(config.templates, CONFIG_DEFAULTS.templates);
  assert.equal(config.design, CONFIG_DEFAULTS.design);
});

test("version alone is enough", () => {
  const config = parseWorkspaceConfig("version: 1\n", CONFIG);
  assert.equal(config.projects, "projects");
});

test("given values override the defaults, others stay default", () => {
  const config = parseWorkspaceConfig("templates: ../shared/templates\n", CONFIG);
  assert.equal(config.templates, "../shared/templates");
  assert.equal(config.assets, "assets");
});

test("an unknown key is rejected with a suggestion and the valid list", () => {
  assert.throws(
    () => parseWorkspaceConfig("template: templates\n", CONFIG),
    (error: Error) => {
      assert.match(error.message, /^\/ws\/pptx-gen\.config\.yml: /);
      assert.match(error.message, /unknown key "template"/);
      assert.match(error.message, /did you mean "templates"/);
      assert.match(error.message, /Valid keys: version, templates, projects, assets, design, designDoc, customize/);
      return true;
    }
  );
});

test("a non-string path value names the offending key", () => {
  assert.throws(
    () => parseWorkspaceConfig("templates:\n  - a\n  - b\n", CONFIG),
    /"templates" must be a non-empty string path, got a list\./
  );
});

test("an unsupported version says which version this install understands", () => {
  assert.throws(
    () => parseWorkspaceConfig("version: 2\n", CONFIG),
    /unsupported version 2 \(this install understands version 1\)\. Update pptx-gen\./
  );
});

test("unparseable YAML is reported with the file path", () => {
  assert.throws(
    () => parseWorkspaceConfig("templates: [unclosed\n", CONFIG),
    /WorkspaceConfigError: \/ws\/pptx-gen\.config\.yml: could not parse YAML/
  );
});

test("a top-level list is rejected", () => {
  assert.throws(() => parseWorkspaceConfig("- a\n", CONFIG), /expected a mapping of settings, got a list\./);
});

test("relative paths resolve against the config file, not the cwd", () => {
  assert.equal(resolveFromConfig("/ws/pptx-gen.config.yml", "templates"), path.resolve("/ws/templates"));
  assert.equal(resolveFromConfig("/ws/nested/pptx-gen.config.yml", "../shared"), path.resolve("/ws/shared"));
});

test("absolute and ~-prefixed paths are honoured as given", () => {
  assert.equal(resolveFromConfig(CONFIG, "/srv/templates"), path.resolve("/srv/templates"));
  assert.equal(resolveFromConfig(CONFIG, "~/shared/templates"), path.join(os.homedir(), "shared", "templates"));
});
