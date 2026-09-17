const API_PREFIX = "/v1/drive/folders/";
const ROUTE = /^\/v1\/drive\/folders\/([A-Za-z0-9_-]{10,200})\/images$/;
const DRIVE_FILES_URL = "https://www.googleapis.com/drive/v3/files";
const MAX_RESULTS = 2_000;
const MAX_PAGES = 10;
const CACHE_SECONDS = 60;
const RETRYABLE_STATUS = new Set([429, 500, 502, 503, 504]);

export interface Env {
  GOOGLE_DRIVE_API_KEY: string;
  RATE_LIMITER: {
    limit(options: { key: string }): Promise<{ success: boolean }>;
  };
}

interface DriveFile {
  id: string;
  name: string;
  mimeType: string;
}

interface DrivePayload {
  nextPageToken?: string;
  files?: unknown;
}

interface ExecutionContextLike {
  waitUntil(promise: Promise<unknown>): void;
}

interface CacheLike {
  match(request: Request): Promise<Response | undefined>;
  put(request: Request, response: Response): Promise<void>;
}

export interface Dependencies {
  fetch: typeof fetch;
  cache: CacheLike;
  sleep(milliseconds: number): Promise<void>;
}

const defaultDependencies = (): Dependencies => ({
  fetch: globalThis.fetch.bind(globalThis),
  cache: (caches as CacheStorage & { default: Cache }).default,
  sleep: (milliseconds) => new Promise((resolve) => setTimeout(resolve, milliseconds)),
});

const securityHeaders = {
  "Content-Type": "application/json; charset=utf-8",
  "X-Content-Type-Options": "nosniff",
  "Referrer-Policy": "no-referrer",
};

function json(payload: unknown, status = 200, extraHeaders: HeadersInit = {}): Response {
  return new Response(JSON.stringify(payload), {
    status,
    headers: { ...securityHeaders, ...extraHeaders },
  });
}

function error(code: string, message: string, status: number): Response {
  return json({ error: { code, message } }, status, { "Cache-Control": "no-store" });
}

function validFile(value: unknown): value is DriveFile {
  if (!value || typeof value !== "object") return false;
  const file = value as Record<string, unknown>;
  return typeof file.id === "string"
    && /^[A-Za-z0-9_-]{10,200}$/.test(file.id)
    && typeof file.name === "string"
    && file.name.length <= 1_000
    && typeof file.mimeType === "string"
    && file.mimeType.startsWith("image/");
}

async function fetchWithRetry(
  url: string,
  apiKey: string,
  dependencies: Dependencies,
): Promise<Response> {
  let response: Response | undefined;
  for (let attempt = 0; attempt < 3; attempt += 1) {
    try {
      response = await dependencies.fetch(url, {
        headers: { Accept: "application/json", "X-Goog-Api-Key": apiKey },
        signal: AbortSignal.timeout(10_000),
      });
    } catch (caught) {
      if (attempt === 2) throw caught;
      await dependencies.sleep(100 * (2 ** attempt));
      continue;
    }
    if (!RETRYABLE_STATUS.has(response.status) || attempt === 2) return response;
    await response.body?.cancel();
    await dependencies.sleep(100 * (2 ** attempt));
  }
  throw new Error("Upstream request failed");
}

async function listFolder(
  folderId: string,
  apiKey: string,
  dependencies: Dependencies,
): Promise<DriveFile[]> {
  const files: DriveFile[] = [];
  let pageToken: string | undefined;

  for (let page = 0; page < MAX_PAGES; page += 1) {
    const url = new URL(DRIVE_FILES_URL);
    url.searchParams.set("q", `'${folderId}' in parents and trashed = false`);
    url.searchParams.set("fields", "nextPageToken,files(id,name,mimeType)");
    url.searchParams.set("pageSize", "1000");
    url.searchParams.set("supportsAllDrives", "true");
    url.searchParams.set("includeItemsFromAllDrives", "true");
    if (pageToken) url.searchParams.set("pageToken", pageToken);

    const response = await fetchWithRetry(url.toString(), apiKey, dependencies);
    if (!response.ok) {
      await response.body?.cancel();
      if (response.status === 404) throw new UpstreamError("FOLDER_NOT_FOUND", 404);
      if (response.status === 400) throw new UpstreamError("INVALID_FOLDER", 400);
      throw new UpstreamError("UPSTREAM_UNAVAILABLE", 502);
    }

    let payload: DrivePayload;
    try {
      payload = await response.json() as DrivePayload;
    } catch {
      throw new UpstreamError("UPSTREAM_UNAVAILABLE", 502);
    }
    if (!payload || typeof payload !== "object" || !Array.isArray(payload.files)) {
      throw new UpstreamError("UPSTREAM_UNAVAILABLE", 502);
    }
    for (const candidate of payload.files) {
      if (validFile(candidate)) files.push(candidate);
      if (files.length > MAX_RESULTS) throw new UpstreamError("RESULT_LIMIT_EXCEEDED", 413);
    }

    pageToken = typeof payload.nextPageToken === "string" && payload.nextPageToken
      ? payload.nextPageToken
      : undefined;
    if (!pageToken) break;
    if (page === MAX_PAGES - 1) throw new UpstreamError("RESULT_LIMIT_EXCEEDED", 413);
  }

  const collator = new Intl.Collator("en", { numeric: true, sensitivity: "base" });
  files.sort((left, right) => collator.compare(left.name, right.name)
    || left.id.localeCompare(right.id));
  return files;
}

class UpstreamError extends Error {
  constructor(readonly code: string, readonly status: number) {
    super(code);
  }
}

export async function handleRequest(
  request: Request,
  env: Env,
  context: ExecutionContextLike,
  dependencies: Dependencies = defaultDependencies(),
): Promise<Response> {
  const url = new URL(request.url);
  if (request.method !== "GET") {
    return error("METHOD_NOT_ALLOWED", "Only GET is supported.", 405);
  }
  if (url.search || !url.pathname.startsWith(API_PREFIX)) {
    return error("NOT_FOUND", "Route not found.", 404);
  }
  const match = ROUTE.exec(url.pathname);
  if (!match) return error("INVALID_FOLDER_ID", "Invalid Google Drive folder ID.", 400);
  if (!env.GOOGLE_DRIVE_API_KEY) {
    console.error("GOOGLE_DRIVE_API_KEY secret is not configured");
    return error("SERVICE_MISCONFIGURED", "Folder listing service is unavailable.", 503);
  }

  const clientKey = request.headers.get("CF-Connecting-IP") || "unknown";
  const rate = await env.RATE_LIMITER.limit({ key: clientKey });
  if (!rate.success) {
    return error("RATE_LIMITED", "Too many requests. Try again later.", 429);
  }

  const cacheKey = new Request(url.toString(), { method: "GET" });
  const cached = await dependencies.cache.match(cacheKey);
  if (cached) return cached;

  try {
    const files = await listFolder(match[1], env.GOOGLE_DRIVE_API_KEY, dependencies);
    const response = json({ files }, 200, {
      "Cache-Control": `public, max-age=${CACHE_SECONDS}`,
    });
    context.waitUntil(dependencies.cache.put(cacheKey, response.clone()));
    return response;
  } catch (caught) {
    if (caught instanceof UpstreamError) {
      const messages: Record<string, string> = {
        FOLDER_NOT_FOUND: "Public Google Drive folder not found.",
        INVALID_FOLDER: "Google Drive rejected the folder request.",
        RESULT_LIMIT_EXCEEDED: `Folder exceeds the ${MAX_RESULTS}-image limit.`,
        UPSTREAM_UNAVAILABLE: "Google Drive is temporarily unavailable.",
      };
      return error(caught.code, messages[caught.code], caught.status);
    }
    console.error("Drive request failed", caught instanceof Error ? caught.name : "unknown");
    return error("UPSTREAM_UNAVAILABLE", "Google Drive is temporarily unavailable.", 502);
  }
}

export default {
  fetch(request: Request, env: Env, context: ExecutionContext): Promise<Response> {
    return handleRequest(request, env, context);
  },
};