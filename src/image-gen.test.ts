import { test } from "node:test";
import assert from "node:assert/strict";
import util from "node:util";
import path from "node:path";
import os from "node:os";
import { existsSync } from "node:fs";
import { mkdir, mkdtemp, readFile, realpath, rm, writeFile } from "node:fs/promises";
import {
  DEFAULT_IMAGE_MODEL,
  generateImages,
  type ImageGenFn,
  type ImageGenRequest,
  LOCK_FILE,
  openAiImageGen,
  resolveImageGeneration
} from "./image-gen.js";
import { loadWorkspace } from "./workspace.js";
import { makePng } from "./test-fixtures.js";

// Every test here runs offline: generation goes through an injected ImageGenFn
// or an injected fetch, and `env: {}` keeps a developer's own key out of it.

const NO_ENV = {};

type FakeGen = ImageGenFn & { calls: ImageGenRequest[] };

function fakeGen(options: { failFor?: string[]; returns?: number } = {}): FakeGen {
  const gen = (async (request: ImageGenRequest) => {
    gen.calls.push(request);
    if (options.failFor?.some((word) => request.prompt.includes(word))) {
      throw new Error("OpenAI returned 400 (moderation_blocked): Your request was rejected.");
    }
    return Array.from({ length: options.returns ?? request.n }, (_unused, index) => makePng(30 + index, 20));
  }) as FakeGen;
  gen.calls = [];
  return gen;
}

function brief(images: string): string {
  return `# Deck\n\n## Narrative arc\n\nTODO\n\n## Images\n\n\`\`\`yaml\n${images}\n\`\`\`\n\n## Claims and numbers\n\nNone.\n`;
}

const TWO_IMAGES = `- id: hero-city
  description: Aerial view of a European city at dawn
  variants: 3
  pick: 2
- id: texture
  description: Soft paper texture`;

async function project(t: { after: (fn: () => unknown) => void }, images = TWO_IMAGES): Promise<string> {
  const dir = await realpath(await mkdtemp(path.join(os.tmpdir(), "pptx-image-gen-")));
  t.after(() => rm(dir, { recursive: true, force: true }));
  await writeFile(path.join(dir, "brief.md"), brief(images), "utf8");
  return dir;
}

test("generating writes every variant into inputs/ and records it in the lock", async (t) => {
  const dir = await project(t);
  const gen = fakeGen();

  const summary = await generateImages({ projectDir: dir, env: NO_ENV, gen });

  assert.deepEqual(
    summary.images.map((image) => [image.id, image.status, image.files]),
    [
      ["hero-city", "generated", ["inputs/hero-city-1.jpg", "inputs/hero-city-2.jpg", "inputs/hero-city-3.jpg"]],
      ["texture", "generated", ["inputs/texture-1.jpg"]]
    ]
  );
  assert.equal(gen.calls.length, 2, "one request per image, however many variants");
  assert.equal(gen.calls[0].n, 3);
  assert.equal(gen.calls[0].size, "1536x1024");
  assert.equal(gen.calls[0].format, "jpeg");
  assert.equal(gen.calls[0].model, DEFAULT_IMAGE_MODEL);
  assert.match(gen.calls[0].prompt, /^Aerial view of a European city at dawn/);
  for (const file of summary.images.flatMap((image) => image.files)) {
    assert.ok(existsSync(path.join(dir, file)), `${file} should exist`);
  }

  const lock = JSON.parse(await readFile(path.join(dir, "inputs", LOCK_FILE), "utf8"));
  assert.deepEqual(Object.keys(lock.images), ["hero-city", "texture"]);
  assert.equal(lock.images["hero-city"].prompt, gen.calls[0].prompt);
  assert.equal(summary.notesUpdated, false, "a clean run leaves the brief alone");
});

test("a second run is cached; changing pick never regenerates; --force always does", async (t) => {
  const dir = await project(t);
  const gen = fakeGen();
  await generateImages({ projectDir: dir, env: NO_ENV, gen });

  const again = await generateImages({ projectDir: dir, env: NO_ENV, gen });
  assert.deepEqual(
    again.images.map((image) => image.status),
    ["cached", "cached"]
  );
  assert.equal(gen.calls.length, 2);

  const briefPath = path.join(dir, "brief.md");
  await writeFile(briefPath, (await readFile(briefPath, "utf8")).replace("pick: 2", "pick: 3"), "utf8");
  await generateImages({ projectDir: dir, env: NO_ENV, gen });
  assert.equal(gen.calls.length, 2, "pick only chooses between files that already exist");

  await writeFile(briefPath, (await readFile(briefPath, "utf8")).replace("at dawn", "at dusk"), "utf8");
  const edited = await generateImages({ projectDir: dir, env: NO_ENV, gen });
  assert.deepEqual(
    edited.images.map((image) => image.status),
    ["generated", "cached"],
    "an edited description regenerates that image only"
  );

  await generateImages({ projectDir: dir, env: NO_ENV, gen, force: true, only: ["texture"] });
  assert.equal(gen.calls.length, 4);
});

test("a file someone put in inputs/ by hand is never overwritten without --force", async (t) => {
  const dir = await project(t);
  await mkdir(path.join(dir, "inputs"));
  const mine = path.join(dir, "inputs", "texture-1.png");
  await writeFile(mine, "my own picture");
  const gen = fakeGen();

  const summary = await generateImages({ projectDir: dir, env: NO_ENV, gen });
  const texture = summary.images.find((image) => image.id === "texture");
  assert.equal(texture?.status, "supplied");
  assert.equal(await readFile(mine, "utf8"), "my own picture");
  assert.equal(gen.calls.length, 1, "only hero-city was generated");

  await generateImages({ projectDir: dir, env: NO_ENV, gen, force: true, only: ["texture"] });
  assert.ok(existsSync(path.join(dir, "inputs", "texture-1.jpg")));
  assert.ok(!existsSync(mine), "--force replaces it, so the old file no longer shadows the new one");
});

test("variants the brief no longer asks for are pruned", async (t) => {
  const dir = await project(t);
  const gen = fakeGen();
  await generateImages({ projectDir: dir, env: NO_ENV, gen });

  const briefPath = path.join(dir, "brief.md");
  await writeFile(
    briefPath,
    (await readFile(briefPath, "utf8")).replace("variants: 3\n  pick: 2", "variants: 1"),
    "utf8"
  );
  await generateImages({ projectDir: dir, env: NO_ENV, gen });

  assert.ok(existsSync(path.join(dir, "inputs", "hero-city-1.jpg")));
  assert.ok(!existsSync(path.join(dir, "inputs", "hero-city-2.jpg")));
  assert.ok(!existsSync(path.join(dir, "inputs", "hero-city-3.jpg")));
});

test("a failure is written into the brief, and cleared once it succeeds", async (t) => {
  const dir = await project(t);
  const briefPath = path.join(dir, "brief.md");
  const original = await readFile(briefPath, "utf8");

  const failing = await generateImages({ projectDir: dir, env: NO_ENV, gen: fakeGen({ failFor: ["paper"] }) });
  const texture = failing.images.find((image) => image.id === "texture");
  assert.equal(texture?.status, "failed");
  assert.match(texture?.reason ?? "", /moderation_blocked/);
  assert.equal(failing.notesUpdated, true);

  const noted = await readFile(briefPath, "utf8");
  assert.match(noted, /> - `texture`: generation failed — OpenAI returned 400 \(moderation_blocked\)/);
  assert.doesNotMatch(noted, /`hero-city`:/, "the image that worked gets no note");

  await generateImages({ projectDir: dir, env: NO_ENV, gen: fakeGen() });
  assert.equal(await readFile(briefPath, "utf8"), original, "the brief is back to exactly what the author wrote");
});

test("with no API key nothing is called, and one note says why", async (t) => {
  const dir = await project(t);

  const summary = await generateImages({ projectDir: dir, env: NO_ENV });
  assert.equal(summary.available, false);
  assert.deepEqual(
    summary.images.map((image) => image.status),
    ["failed", "failed"]
  );
  assert.ok(!existsSync(path.join(dir, "inputs")), "no files and no lock");

  const noted = await readFile(path.join(dir, "brief.md"), "utf8");
  assert.match(noted, /> - OPENAI_API_KEY is not set, so `hero-city`, `texture` were not generated\./);
});

test("a dry run composes prompts without calling anything or writing any file", async (t) => {
  const dir = await project(t);
  const before = await readFile(path.join(dir, "brief.md"), "utf8");
  const gen = fakeGen();

  const summary = await generateImages({ projectDir: dir, env: NO_ENV, gen, dryRun: true });
  assert.deepEqual(
    summary.images.map((image) => image.status),
    ["planned", "planned"]
  );
  assert.match(summary.images[0].prompt, /Aerial view/);
  assert.equal(gen.calls.length, 0);
  assert.ok(!existsSync(path.join(dir, "inputs")));
  assert.equal(await readFile(path.join(dir, "brief.md"), "utf8"), before);
});

test("--only names an image that exists, or explains which one was meant", async (t) => {
  const dir = await project(t);
  await assert.rejects(
    () => generateImages({ projectDir: dir, env: NO_ENV, gen: fakeGen(), only: ["texure"] }),
    /No image 'texure'.*did you mean 'texture'\?/
  );
  await assert.rejects(() => generateImages({ projectDir: path.join(dir, "nowhere"), env: NO_ENV }), /No brief\.md/);
});

test("customize.md's style and examples reach the prompt; a missing example is noted, not fatal", async (t) => {
  const root = await realpath(await mkdtemp(path.join(os.tmpdir(), "pptx-image-ws-")));
  t.after(() => rm(root, { recursive: true, force: true }));
  await writeFile(path.join(root, "pptx-gen.config.yml"), "version: 1\n", "utf8");
  await mkdir(path.join(root, "assets", "style"), { recursive: true });
  await writeFile(path.join(root, "assets", "style", "ref.png"), makePng(4, 4));
  await writeFile(
    path.join(root, "customize.md"),
    "# Customizations\n\n## Image generation\n\nEditorial photography.\n\nExamples:\n- assets/style/ref.png\n- assets/style/gone.png\n",
    "utf8"
  );
  const dir = path.join(root, "projects", "deck");
  await mkdir(dir, { recursive: true });
  await writeFile(path.join(dir, "brief.md"), brief("- id: hero\n  description: A calm harbour"), "utf8");

  const gen = fakeGen();
  const summary = await generateImages({ projectDir: dir, workspace: loadWorkspace(root, "flag"), env: NO_ENV, gen });

  assert.equal(gen.calls[0].references.length, 1);
  assert.equal(gen.calls[0].references[0].mime, "image/png");
  assert.match(gen.calls[0].prompt, /Style: Editorial photography\./);
  assert.match(gen.calls[0].prompt, /style references only/);
  assert.equal(summary.images[0].status, "generated", "the missing example did not stop the image");
  assert.match(summary.images[0].notes.join(" "), /assets\/style\/gone\.png was not found/);
  assert.match(await readFile(path.join(dir, "brief.md"), "utf8"), /`hero`: Reference assets\/style\/gone\.png/);
});

test("the key comes from the environment first, then the workspace .env, never printed", async (t) => {
  const root = await mkdtemp(path.join(os.tmpdir(), "pptx-image-env-"));
  t.after(() => rm(root, { recursive: true, force: true }));

  assert.deepEqual(resolveImageGeneration(root, {}), {
    available: false,
    keySource: null,
    model: DEFAULT_IMAGE_MODEL,
    quality: "medium",
    baseUrl: "https://api.openai.com/v1",
    apiKey: undefined
  });

  if (typeof util.parseEnv !== "function") return t.skip("this Node has no util.parseEnv, so .env is not read");
  await writeFile(path.join(root, ".env"), 'OPENAI_API_KEY="sk-from-file"\nPPTX_GEN_IMAGE_MODEL=gpt-image-2\n', "utf8");

  const fromFile = resolveImageGeneration(root, {});
  assert.equal(fromFile.keySource, ".env");
  assert.equal(fromFile.apiKey, "sk-from-file");
  assert.equal(fromFile.model, "gpt-image-2");

  const fromEnv = resolveImageGeneration(root, { OPENAI_API_KEY: "sk-from-env", PPTX_GEN_IMAGE_MODEL: "" });
  assert.equal(fromEnv.keySource, "env");
  assert.equal(fromEnv.apiKey, "sk-from-env");
  assert.equal(fromEnv.model, "gpt-image-2", "an empty variable counts as unset");
});

type Call = { url: string; init: RequestInit };

function fakeFetch(responses: Response[]): typeof fetch & { calls: Call[] } {
  const fn = (async (url: string | URL | Request, init?: RequestInit) => {
    fn.calls.push({ url: String(url), init: init ?? {} });
    const next = responses.shift();
    if (!next) throw new Error("no more fake responses");
    return next;
  }) as typeof fetch & { calls: Call[] };
  fn.calls = [];
  return fn;
}

const IMAGE = makePng(8, 8);

function imagesResponse(count = 1): Response {
  return Response.json({ data: Array.from({ length: count }, () => ({ b64_json: IMAGE.toString("base64") })) });
}

const REQUEST: ImageGenRequest = {
  model: "gpt-image-2.5-flare",
  prompt: "A harbour",
  size: "1536x1024",
  quality: "medium",
  format: "jpeg",
  n: 2,
  references: []
};

test("a request without references is JSON to /images/generations, decoded from base64", async () => {
  const fetch = fakeFetch([imagesResponse(2)]);
  const images = await openAiImageGen({ apiKey: "sk-test", fetch })(REQUEST);

  assert.equal(images.length, 2);
  assert.ok(images[0].equals(IMAGE));
  assert.equal(fetch.calls[0].url, "https://api.openai.com/v1/images/generations");
  const headers = fetch.calls[0].init.headers as Record<string, string>;
  assert.equal(headers.Authorization, "Bearer sk-test");
  assert.deepEqual(JSON.parse(String(fetch.calls[0].init.body)), {
    model: "gpt-image-2.5-flare",
    prompt: "A harbour",
    n: 2,
    size: "1536x1024",
    quality: "medium",
    output_format: "jpeg"
  });
});

test("a request with references is multipart to /images/edits, one image[] part each", async () => {
  const fetch = fakeFetch([imagesResponse()]);
  await openAiImageGen({ apiKey: "sk-test", baseUrl: "https://proxy.example/v1/", fetch })({
    ...REQUEST,
    references: [
      { name: "a.png", bytes: IMAGE, mime: "image/png" },
      { name: "b.png", bytes: IMAGE, mime: "image/png" }
    ]
  });

  assert.equal(fetch.calls[0].url, "https://proxy.example/v1/images/edits");
  const form = fetch.calls[0].init.body as FormData;
  assert.ok(form instanceof FormData);
  assert.equal(form.getAll("image[]").length, 2);
  assert.equal(form.get("n"), "2");
  assert.equal(form.get("output_format"), "jpeg");
});

test("a rate limit is retried after retry-after; a rejected prompt is not retried", async () => {
  const waits: number[] = [];
  const sleep = async (ms: number) => {
    waits.push(ms);
  };

  const limited = fakeFetch([
    new Response(JSON.stringify({ error: { message: "Rate limit reached", code: "rate_limit_exceeded" } }), {
      status: 429,
      headers: { "retry-after": "3" }
    }),
    imagesResponse()
  ]);
  await openAiImageGen({ apiKey: "sk-test", fetch: limited, sleep })(REQUEST);
  assert.equal(limited.calls.length, 2);
  assert.deepEqual(waits, [3000]);

  const rejected = fakeFetch([
    Response.json(
      { error: { message: "Your request was rejected by the safety system.", code: "moderation_blocked" } },
      { status: 400 }
    )
  ]);
  await assert.rejects(
    () => openAiImageGen({ apiKey: "sk-test", fetch: rejected, sleep })(REQUEST),
    /OpenAI returned 400 \(moderation_blocked\): Your request was rejected by the safety system\./
  );
  assert.equal(rejected.calls.length, 1);

  const broke = fakeFetch([
    Response.json(
      { error: { message: "You exceeded your current quota.", code: "insufficient_quota" } },
      { status: 429 }
    )
  ]);
  await assert.rejects(
    () => openAiImageGen({ apiKey: "sk-test", fetch: broke, sleep })(REQUEST),
    /insufficient_quota/,
    "an empty account is not a rate limit, and waiting will not fix it"
  );
  assert.equal(broke.calls.length, 1);
});

test("server errors are retried with backoff until the attempts run out", async () => {
  const waits: number[] = [];
  const fetch = fakeFetch([
    new Response("upstream down", { status: 502 }),
    new Response("upstream down", { status: 502 }),
    new Response("upstream down", { status: 502 })
  ]);
  await assert.rejects(
    () =>
      openAiImageGen({
        apiKey: "sk-test",
        fetch,
        maxAttempts: 3,
        sleep: async (ms) => {
          waits.push(ms);
        }
      })(REQUEST),
    /OpenAI returned 502: upstream down/
  );
  assert.deepEqual(waits, [2000, 4000]);
});

test("a rejected key stops the run after one call, and no piece of the key reaches the brief", async (t) => {
  const dir = await project(t);
  const fetch = fakeFetch([
    Response.json(
      {
        error: {
          message: "Incorrect API key provided: sk-proj-abc*****wxyz. You can find your API key at ...",
          code: "invalid_api_key"
        }
      },
      { status: 401 }
    )
  ]);

  const summary = await generateImages({
    projectDir: dir,
    env: NO_ENV,
    gen: openAiImageGen({ apiKey: "sk-proj-abcdefwxyz", fetch })
  });
  assert.equal(fetch.calls.length, 1, "the second image is not even tried");
  assert.deepEqual(
    summary.images.map((image) => image.status),
    ["failed", "failed"]
  );
  const noted = await readFile(path.join(dir, "brief.md"), "utf8");
  assert.match(noted, /OpenAI returned 401 \(invalid_api_key\): Incorrect API key provided: sk-…\. /);
  assert.doesNotMatch(noted, /abc|wxyz/);
});
