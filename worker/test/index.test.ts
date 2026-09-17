import { beforeEach, describe, expect, it, vi } from "vitest";
import { handleRequest, type Dependencies, type Env } from "../src/index";

const FOLDER_ID = "folder-12345";
const endpoint = `https://worker.example/v1/drive/folders/${FOLDER_ID}/images`;

function setup() {
  const pending: Promise<unknown>[] = [];
  const cache = {
    match: vi.fn<Dependencies["cache"]["match"]>(async () => undefined),
    put: vi.fn(async () => undefined),
  };
  const dependencies: Dependencies = {
    fetch: vi.fn(),
    cache,
    sleep: vi.fn(async () => undefined),
  };
  const env: Env = {
    GOOGLE_DRIVE_API_KEY: "super-secret-key",
    RATE_LIMITER: { limit: vi.fn(async () => ({ success: true })) },
  };
  const context = { waitUntil: (promise: Promise<unknown>) => pending.push(promise) };
  return { cache, context, dependencies, env, pending };
}

describe("Drive folder listing Worker", () => {
  beforeEach(() => vi.restoreAllMocks());

  it("rejects methods, unknown routes, query strings, and malformed IDs", async () => {
    const { context, dependencies, env } = setup();
    expect((await handleRequest(new Request(endpoint, { method: "POST" }), env, context, dependencies)).status).toBe(405);
    expect((await handleRequest(new Request("https://worker.example/health"), env, context, dependencies)).status).toBe(404);
    expect((await handleRequest(new Request(`${endpoint}?key=bad`), env, context, dependencies)).status).toBe(404);
    expect((await handleRequest(new Request("https://worker.example/v1/drive/folders/bad!/images"), env, context, dependencies)).status).toBe(400);
    expect(dependencies.fetch).not.toHaveBeenCalled();
  });

  it("paginates, filters images, and naturally sorts names", async () => {
    const state = setup();
    vi.mocked(state.dependencies.fetch)
      .mockResolvedValueOnce(Response.json({
        nextPageToken: "page-two",
        files: [
          { id: "image-00010", name: "photo10.jpg", mimeType: "image/jpeg" },
          { id: "document-01", name: "notes.pdf", mimeType: "application/pdf" },
        ],
      }))
      .mockResolvedValueOnce(Response.json({
        files: [{ id: "image-00002", name: "photo2.png", mimeType: "image/png" }],
      }));

    const response = await handleRequest(new Request(endpoint, {
      headers: { "CF-Connecting-IP": "192.0.2.1" },
    }), state.env, state.context, state.dependencies);
    await Promise.all(state.pending);

    expect(response.status).toBe(200);
    expect((await response.json() as { files: Array<{ id: string }> }).files.map((file) => file.id))
      .toEqual(["image-00002", "image-00010"]);
    expect(state.dependencies.fetch).toHaveBeenCalledTimes(2);
    const secondUrl = new URL(vi.mocked(state.dependencies.fetch).mock.calls[1][0] as string);
    expect(secondUrl.searchParams.get("pageToken")).toBe("page-two");
    expect(secondUrl.searchParams.has("key")).toBe(false);
    expect(vi.mocked(state.dependencies.fetch).mock.calls[0][1]?.headers)
      .toMatchObject({ "X-Goog-Api-Key": "super-secret-key" });
    expect(state.cache.put).toHaveBeenCalledOnce();
  });

  it("serves a cached response without contacting Google", async () => {
    const state = setup();
    state.cache.match.mockResolvedValue(Response.json({ files: [] }));
    const response = await handleRequest(new Request(endpoint), state.env, state.context, state.dependencies);
    expect(response.status).toBe(200);
    expect(state.dependencies.fetch).not.toHaveBeenCalled();
  });

  it("rejects folders above the bounded image result limit", async () => {
    const state = setup();
    vi.mocked(state.dependencies.fetch).mockResolvedValue(Response.json({
      files: Array.from({ length: 2_001 }, (_, index) => ({
        id: `image-${String(index).padStart(5, "0")}`,
        name: `image${index}.jpg`,
        mimeType: "image/jpeg",
      })),
    }));
    const response = await handleRequest(
      new Request(endpoint), state.env, state.context, state.dependencies,
    );
    expect(response.status).toBe(413);
    expect(await response.json()).toMatchObject({
      error: { code: "RESULT_LIMIT_EXCEEDED" },
    });
  });

  it("rate limits by Cloudflare client IP before cache lookup", async () => {
    const state = setup();
    vi.mocked(state.env.RATE_LIMITER.limit).mockResolvedValue({ success: false });
    const response = await handleRequest(new Request(endpoint, {
      headers: { "CF-Connecting-IP": "192.0.2.9" },
    }), state.env, state.context, state.dependencies);
    expect(response.status).toBe(429);
    expect(state.env.RATE_LIMITER.limit).toHaveBeenCalledWith({ key: "192.0.2.9" });
    expect(state.cache.match).not.toHaveBeenCalled();
  });

  it("retries bounded transient failures and sanitizes the final error", async () => {
    const state = setup();
    vi.mocked(state.dependencies.fetch).mockResolvedValue(new Response(
      JSON.stringify({ error: { message: "secret upstream detail" } }), { status: 503 },
    ));
    const response = await handleRequest(new Request(endpoint), state.env, state.context, state.dependencies);
    expect(response.status).toBe(502);
    expect(state.dependencies.fetch).toHaveBeenCalledTimes(3);
    expect(state.dependencies.sleep).toHaveBeenCalledTimes(2);
    expect(await response.text()).not.toContain("secret upstream detail");
    expect(await (await handleRequest(new Request(endpoint), { ...state.env, GOOGLE_DRIVE_API_KEY: "" }, state.context, state.dependencies)).json())
      .toMatchObject({ error: { code: "SERVICE_MISCONFIGURED" } });
  });
});