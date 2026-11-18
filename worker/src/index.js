/**
 * tierstats-publish Worker (replacement, built 2026-11-09).
 *
 * Chain: owned Google Sheet (authoritative, auto-recalc) -> this Worker
 * -> data.js committed to GitHub -> Pages redeploy.
 *
 * The Worker never writes to the Sheet. It only READS the Sheet's public
 * gviz export (the Sheet must be shared as "Anyone with the link: Viewer";
 * everything read here is already published on the website, so nothing
 * private is exposed) and commits the regenerated data.js.
 *
 * Endpoints (must stay compatible with Kode.gs + the site footer/admin):
 *   GET  /                diagnostics text
 *   GET  /sync-status     { matches, players, ageMinutes, stale, lastError, generated }
 *   POST /sync            { password } -> { ok, changed, matches, players } | { error }
 *   cron every 5 minutes      same as POST /sync (trusted, no password)
 *
 * Env (set in Cloudflare dashboard, never in code):
 *   SHEET_ID      spreadsheet id (var)
 *   GITHUB_REPO   e.g. Stats4Tanks/stats4tanks.github.io (var)
 *   GITHUB_BRANCH e.g. main (var)
 *   GITHUB_TOKEN  fine-grained PAT, Contents read+write on the repo (secret)
 *   WORKER_SECRET admin/sync password shared with the Sheet (secret)
 */

const SERVERS = ["London", "Newark", "Dallas", "Atlanta", "Fremont", "Singapore", "Sydney"];
const DATA_PATH = "data.js";
const STALE_MINUTES = 30;

const DEFAULTS = {
  GITHUB_REPO: "Stats4Tanks/stats4tanks.github.io",
  GITHUB_BRANCH: "main",
};

function json(data, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: { "Content-Type": "application/json", "Cache-Control": "no-store" },
  });
}

function secretsEqual(a, b) {
  a = String(a || "");
  b = String(b || "");
  if (a.length === 0 || a.length !== b.length) return false;
  let d = 0;
  for (let i = 0; i < a.length; i++) d |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return d === 0;
}

function b64encode(str) {
  const bytes = new TextEncoder().encode(str);
  let bin = "";
  const CHUNK = 8192;
  for (let i = 0; i < bytes.length; i += CHUNK) {
    bin += String.fromCharCode.apply(null, bytes.subarray(i, i + CHUNK));
  }
  return btoa(bin);
}

function b64decode(b64) {
  const bin = atob(b64.replace(/\s+/g, ""));
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return new TextDecoder().decode(bytes);
}

function round1(x) {
  const n = Number(x);
  if (!Number.isFinite(n)) return 0;
  return Math.round(n * 10) / 10;
}

function asInt(x, fallback = 0) {
  const n = Number(x);
  return Number.isFinite(n) ? Math.trunc(n) : fallback;
}

/** Excel serial (e.g. 46302) -> "YYYY-MM-DD". Returns "" for junk. */
function serialToISO(v) {
  const n = Number(v);
  if (!Number.isFinite(n) || n < 20000 || n > 80000) return "";
  const ms = Math.round((n - 25569) * 86400) * 1000;
  const d = new Date(ms);
  return Number.isNaN(d.getTime()) ? "" : d.toISOString().slice(0, 10);
}

/** gviz "Date(2026,9,18)" (month is 0-based) -> "2026-10-18". */
function gvizDateToISO(s) {
  const m = /^Date\((\d+),(\d+),(\d+)/.exec(String(s || ""));
  if (!m) return "";
  const d = new Date(Date.UTC(+m[1], +m[2], +m[3]));
  return Number.isNaN(d.getTime()) ? "" : d.toISOString().slice(0, 10);
}

function todayISO() {
  return new Date().toISOString().slice(0, 10);
}

/* ------------------------------------------------------------------ *
 * Google Sheet reading (gviz JSON, no auth needed beyond link sharing)
 * ------------------------------------------------------------------ */

// Deterministic per-tab ranges: each range starts exactly at that tab's real
// header row (title/instruction rows above are excluded). Requested with
// headers=1 so gviz puts the header text into table.cols[].label (labels are
// never type-coerced, unlike header cells inside numeric columns), and the
// Worker prepends those labels as a synthetic header row for the parsers.
const SHEET_RANGES = {
  Matches: "A3:L",
  Aliases: "A3:C",
  Seeds: "A3:D",
  Settings: "A3:C",
  Leaderboard: "A5:L",
  "All Players": "A5:L",
  Provisional: "A5:H",
  Inactive: "A5:I",
};

async function fetchSheetGrid(sheetId, sheetName) {
  const range = SHEET_RANGES[sheetName];
  if (!range) throw new Error("No range configured for Sheet tab '" + sheetName + "'.");
  const url =
    "https://docs.google.com/spreadsheets/d/" +
    encodeURIComponent(sheetId) +
    "/gviz/tq?tqx=out:json&headers=1&sheet=" +
    encodeURIComponent(sheetName) +
    "&range=" +
    encodeURIComponent(range);
  const resp = await fetch(url, { headers: { Accept: "application/json" } });
  const text = await resp.text();
  if (!resp.ok) throw new Error("Sheet tab '" + sheetName + "' unreachable (HTTP " + resp.status + ").");
  const start = text.indexOf("{");
  const end = text.lastIndexOf("}");
  if (start < 0 || end < 0) {
    throw new Error(
      "Sheet tab '" + sheetName + "' did not return data. " +
      "Check link sharing: Anyone with the link can view."
    );
  }
  let payload;
  try {
    payload = JSON.parse(text.slice(start, end + 1));
  } catch (e) {
    throw new Error("Sheet tab '" + sheetName + "' returned unreadable data.");
  }
  const table = (payload && payload.table) || {};
  const cols = table.cols || [];
  const headerRow = cols.map((c) => {
    const label = c && c.label != null ? String(c.label).trim() : "";
    return { v: label === "" ? null : label, f: "" };
  });
  const rows = table.rows || [];
  const grid = rows.map((r) => (r.c || []).map((c) => ({ v: c ? c.v : null, f: c && c.f != null ? String(c.f) : "" })));
  return [headerRow].concat(grid); // synthetic header row at index 0; data follows
}

function cellNum(cell) {
  if (!cell || cell.v === null || cell.v === undefined) return NaN;
  if (typeof cell.v === "number") return cell.v;
  const s = String(cell.v).trim();
  if (s === "") return NaN;
  if (/^Date\(/.test(s)) return NaN;
  const n = Number(s);
  return Number.isFinite(n) ? n : NaN;
}

/** Percent-aware number: formatted "84.1%" -> 84.1, fraction 0.841 -> 84.1. */
function cellPct(cell) {
  if (!cell) return NaN;
  if (cell.f && /%\s*$/.test(cell.f)) {
    const n = Number(cell.f.replace(/%/g, "").trim());
    return Number.isFinite(n) ? n : NaN;
  }
  const n = cellNum(cell);
  if (!Number.isFinite(n)) return NaN;
  return n > 0 && n <= 1 ? n * 100 : n;
}

/** ISO date from a gviz cell: serial number, Date(...), or "YYYY-MM-DD". */
function cellISODate(cell) {
  if (!cell || cell.v === null || cell.v === undefined) return "";
  if (typeof cell.v === "number") return serialToISO(cell.v);
  const s = String(cell.v).trim();
  if (s === "") return "";
  if (/^Date\(/.test(s)) return gvizDateToISO(s);
  return /^\d{4}-\d{2}-\d{2}/.test(s) ? s.slice(0, 10) : "";
}

function findHeaderRow(grid, pred) {
  for (let i = 0; i < grid.length; i++) {
    if (pred(grid[i], i)) return i;
  }
  return -1;
}

function headerIndex(headerRow, names) {
  const want = names.map((n) => n.toLowerCase());
  for (let i = 0; i < headerRow.length; i++) {
    if (want.includes(cellStrRaw(headerRow[i]).toLowerCase())) return i;
  }
  return -1;
}

function cellStrRaw(cell) {
  if (!cell || cell.v === null || cell.v === undefined) return "";
  if (typeof cell.v === "number") return "";
  const s = String(cell.v);
  if (/^Date\(/.test(s)) return "";
  return s.trim();
}

function normServer(s) {
  const t = String(s || "").trim();
  const hit = SERVERS.find((x) => x.toLowerCase() === t.toLowerCase());
  return hit || "";
}

/* ------------------------------------------------------------------ *
 * Sheet -> data.js sections
 * ------------------------------------------------------------------ */

function buildAliases(grid) {
  const exact = {};
  const lower = {};
  for (const row of grid) {
    const a = cellStrRaw(row[0]);
    const c = cellStrRaw(row[1]);
    if (!a || !c) continue;
    const al = a.toLowerCase();
    const cl = c.toLowerCase();
    if (al === "wrongname" || al === "alias" || cl === "correctname" || cl === "canonical name") continue; // header row
    exact[a] = c;
    lower[al] = c;
  }
  return { exact, lower };
}

function canonical(name, aliases) {
  let n = String(name || "").trim();
  const seen = new Set();
  for (let i = 0; i < 10; i++) {
    const hit = aliases.lower[n.toLowerCase()];
    if (!hit || hit === n || seen.has(n)) return n;
    seen.add(n);
    n = hit;
  }
  return n;
}

function buildMatches(grid, aliases) {
  const hi = findHeaderRow(grid, (row) =>
    row.some((c) => cellStrRaw(c).toLowerCase() === "player 1")
  );
  if (hi < 0) throw new Error("Matches tab: header row (Player 1) not found.");
  const head = grid[hi];
  const ai = headerIndex(head, ["player 1"]);
  const bi = headerIndex(head, ["player 2"]);
  const sai = headerIndex(head, ["score 1"]);
  const sbi = headerIndex(head, ["score 2"]);
  const di = headerIndex(head, ["date"]);
  if (ai < 0 || bi < 0 || sai < 0 || sbi < 0) {
    throw new Error("Matches tab: expected columns Date/Player 1/Player 2/Score 1/Score 2.");
  }
  // Server block sits right after the score columns (G..L when A..E are 0..4).
  const base = sbi + 2;
  const out = [];
  for (let r = hi + 1; r < grid.length; r++) {
    const row = grid[r];
    const a = canonical(cellStrRaw(row[ai]), aliases);
    const b = canonical(cellStrRaw(row[bi]), aliases);
    const sa = cellNum(row[sai]);
    const sb = cellNum(row[sbi]);
    if (!a || !b || !Number.isFinite(sa) || !Number.isFinite(sb)) continue;
    const m = { a, b, sa, sb, date: di >= 0 ? cellISODate(row[di]) : "" };
    const s1 = normServer(cellStrRaw(row[base]));
    const s2 = normServer(cellStrRaw(row[base + 3]));
    const servers = [];
    if (s1) servers.push(s1);
    if (s2 && s2 !== s1) servers.push(s2);
    if (servers.length) m.sv = servers;
    if (servers.length === 2) {
      const h = cellNum(row[base + 1]);
      const ic = cellNum(row[base + 2]);
      const k = cellNum(row[base + 4]);
      const l = cellNum(row[base + 5]);
      if ([h, ic, k, l].every(Number.isFinite)) m.svc = [[h, ic], [k, l]];
    }
    out.push(m);
  }
  return out;
}

function buildSeeds(grid, aliases) {
  const seeds = {};
  for (const row of grid) {
    const name = cellStrRaw(row[0]);
    const old = cellNum(row[1]);
    if (!name || !Number.isFinite(old)) continue;
    if (name.toLowerCase() === "player") continue; // header row
    seeds[canonical(name, aliases)] = old;
  }
  return seeds;
}

function buildSettings(grid) {
  const out = [];
  for (const row of grid) {
    const name = cellStrRaw(row[0]);
    if (!name || name.toLowerCase() === "setting") continue; // title/header
    const cell = row[1];
    if (!cell || cell.v === null || cell.v === undefined || String(cell.v).trim() === "") continue;
    let value = typeof cell.v === "number" ? cell.v : String(cell.v).trim();
    if (typeof value === "number" && /date/i.test(name)) {
      const iso = serialToISO(value);
      if (iso) value = iso;
    }
    out.push({ name, value, desc: cellStrRaw(row[2]) });
  }
  return out;
}

function buildInactive(grid, aliases) {
  const list = [];
  for (const row of grid) {
    const name = cellStrRaw(row[0]);
    if (!name || name.toLowerCase() === "player") continue;
    list.push(canonical(name, aliases));
  }
  return list;
}

function mapPlayerRow(row, head) {
  const get = (names) => {
    const i = headerIndex(head, names);
    return i >= 0 ? row[i] : undefined;
  };
  const name = cellStrRaw(get(["player"]));
  const rating = round1(cellNum(get(["rating"])));
  if (!name || !Number.isFinite(rating)) return null;
  // Derive the win rate from the W/Matches columns instead of the sheet's
  // "Win %" column: that column is written as a 0-1 fraction and has been
  // syncing as a flat 0 for every player, which made the board read 0.0%.
  const winMatches = asInt(cellNum(get(["matches"])));
  const winWins = asInt(cellNum(get(["w"])));
  const winPct = winMatches > 0 ? (winWins / winMatches) * 100 : 0;
  return {
    rank: asInt(cellNum(get(["rank"])), 0),
    name,
    rating,
    glicko: round1(cellNum(get(["glicko"]))),
    rd: round1(cellNum(get(["rd"]))),
    matches: asInt(cellNum(get(["matches"]))),
    w: asInt(cellNum(get(["w"]))),
    l: asInt(cellNum(get(["l"]))),
    d: asInt(cellNum(get(["d"]))),
    winPct: round1(Number.isFinite(winPct) ? winPct : 0),
    opponents: asInt(cellNum(get(["opponents"]))),
    avgOpp: round1(cellNum(get(["avg opp rating"]))),
    lastMatch: cellISODate(get(["last match"])),
  };
}

function buildBoardTabs(lbGrid, apGrid, provGrid) {
  const lbHi = findHeaderRow(lbGrid, (row) =>
    row.some((c) => cellStrRaw(c).toLowerCase() === "rank")
  );
  if (lbHi < 0) throw new Error("Leaderboard tab: header row (Rank) not found.");
  const lbHead = lbGrid[lbHi];
  const qualified = [];
  for (let r = lbHi + 1; r < lbGrid.length; r++) {
    const p = mapPlayerRow(lbGrid[r], lbHead);
    if (!p) continue;
    const { lastMatch, ...rest } = p; // qualified entries carry no lastMatch
    void lastMatch;
    qualified.push(rest);
  }

  const apHi = findHeaderRow(apGrid, (row) => {
    const cells = row.map((c) => cellStrRaw(c).toLowerCase());
    return cells.includes("player") && cells.includes("rating");
  });
  if (apHi < 0) throw new Error("All Players tab: header row not found.");
  const apHead = apGrid[apHi];
  const players = [];
  for (let r = apHi + 1; r < apGrid.length; r++) {
    const p = mapPlayerRow(apGrid[r], apHead);
    if (!p) continue;
    const { rank, ...rest } = p; // players entries carry no rank
    void rank;
    players.push(rest);
  }

  const provNames = new Set();
  for (const row of provGrid) {
    const name = cellStrRaw(row[0]);
    if (!name || name.toLowerCase() === "player") continue;
    provNames.add(name);
  }
  return { qualified, players, provNames };
}

function buildHistories(matches) {
  const hist = {};
  const push = (who, opp, f, a, date, sv) => {
    if (!hist[who]) hist[who] = [];
    const e = { opp, for_: f, against: a, res: f > a ? "W" : f < a ? "L" : "D", date };
    if (sv) e.sv = sv.slice();
    hist[who].push(e);
  };
  for (const m of matches) {
    push(m.a, m.b, m.sa, m.sb, m.date, m.sv);
    push(m.b, m.a, m.sb, m.sa, m.date, m.sv);
  }
  return hist;
}

/* ------------------------------------------------------------------ *
 * GitHub Contents API
 * ------------------------------------------------------------------ */

function ghHeaders(env) {
  return {
    Authorization: "Bearer " + env.GITHUB_TOKEN,
    Accept: "application/vnd.github+json",
    "User-Agent": "tierstats-publish-worker",
    "Content-Type": "application/json",
  };
}

function repoVars(env) {
  return {
    repo: env.GITHUB_REPO || DEFAULTS.GITHUB_REPO,
    branch: env.GITHUB_BRANCH || DEFAULTS.GITHUB_BRANCH,
  };
}

async function ghGetFile(env) {
  const { repo, branch } = repoVars(env);
  const url =
    "https://api.github.com/repos/" + repo + "/contents/" + DATA_PATH + "?ref=" + encodeURIComponent(branch);
  const resp = await fetch(url, { headers: ghHeaders(env) });
  if (resp.status === 404) return null; // first publish
  if (!resp.ok) throw new Error("GitHub read failed (HTTP " + resp.status + ").");
  const data = await resp.json();
  return { sha: data.sha, body: b64decode(data.content || "") };
}

async function ghPutFile(env, sha, body, message) {
  const { repo, branch } = repoVars(env);
  const url = "https://api.github.com/repos/" + repo + "/contents/" + DATA_PATH;
  const payload = { message, content: b64encode(body), branch };
  if (sha) payload.sha = sha;
  const resp = await fetch(url, {
    method: "PUT",
    headers: ghHeaders(env),
    body: JSON.stringify(payload),
  });
  if (!resp.ok) {
    const t = await resp.text();
    throw new Error("GitHub write failed (HTTP " + resp.status + "): " + t.slice(0, 160));
  }
  return resp.json();
}

async function ghLastCommitDate(env) {
  const { repo, branch } = repoVars(env);
  const url =
    "https://api.github.com/repos/" + repo + "/commits?path=" + encodeURIComponent(DATA_PATH) +
    "&sha=" + encodeURIComponent(branch) + "&per_page=1";
  const resp = await fetch(url, { headers: ghHeaders(env) });
  if (!resp.ok) throw new Error("GitHub commits read failed (HTTP " + resp.status + ").");
  const list = await resp.json();
  const iso = list && list[0] && list[0].commit && (list[0].commit.committer || {}).date;
  return iso || null;
}

/* ------------------------------------------------------------------ *
 * Pages deployment recovery (legacy branch builds).
 *
 * Failure mode seen in production: a Sheet sync commit lands on GitHub
 * but the corresponding Pages deployment fails with an internal server
 * error. The next cron then sees unchanged data (changed=false) and does
 * nothing, so the site stays broken until the next manual commit.
 *
 * Recovery: when the latest Pages build for the current branch HEAD is
 * "errored", request a rebuild via POST /pages/builds. That endpoint
 * rebuilds the latest revision WITHOUT creating a git commit, so no
 * unchanged-data commits and no extra data builds are produced.
 *
 * Bounds (GitHub documents a soft limit of ~10 Pages builds/hour):
 * - Only acts when the latest build errored AND targets the current HEAD
 *   (a newer push means its own build is pending/in flight — leave it).
 * - Just-pushed ticks skip recovery (the push-triggered build is in flight).
 * - Exponential backoff on consecutive same-commit failures: 10/20/40/60m.
 * - Gives up after PAGES_MAX_REBUILDS consecutive failures and surfaces it
 *   (sync-status lastError), so a human looks instead of burning builds.
 * - Never throws: any API hiccup degrades to "unknown" and the normal
 *   sync result stands. NOTE: requesting a build needs Pages:write on the
 *   PAT (Contents read+write alone yields HTTP 403 here); until the owner
 *   grants it, rebuilds report "forbidden" and recovery stays dormant.
 * ------------------------------------------------------------------ */

const PAGES_MAX_REBUILDS = 5;
const PAGES_BASE_BACKOFF_MIN = 10;
const PAGES_MAX_BACKOFF_MIN = 60;

async function ghHeadSha(env) {
  const { repo, branch } = repoVars(env);
  const url = "https://api.github.com/repos/" + repo + "/commits/" + encodeURIComponent(branch) + "?per_page=1";
  const resp = await fetch(url, { headers: ghHeaders(env) });
  if (!resp.ok) throw new Error("GitHub HEAD read failed (HTTP " + resp.status + ").");
  const data = await resp.json();
  return data.sha || null;
}

async function ghPagesLatest(env) {
  const { repo } = repoVars(env);
  const resp = await fetch("https://api.github.com/repos/" + repo + "/pages/builds/latest", { headers: ghHeaders(env) });
  if (resp.status === 404) return null; // Pages not enabled / no builds yet
  if (!resp.ok) throw new Error("Pages builds read failed (HTTP " + resp.status + ").");
  return resp.json();
}

async function ghPagesRecent(env, n) {
  const { repo } = repoVars(env);
  const resp = await fetch("https://api.github.com/repos/" + repo + "/pages/builds?per_page=" + n, { headers: ghHeaders(env) });
  if (!resp.ok) throw new Error("Pages builds list failed (HTTP " + resp.status + ").");
  const list = await resp.json();
  return Array.isArray(list) ? list : [];
}

async function ghRequestPagesBuild(env) {
  const { repo } = repoVars(env);
  const resp = await fetch("https://api.github.com/repos/" + repo + "/pages/builds", {
    method: "POST",
    headers: ghHeaders(env),
    body: JSON.stringify({}),
  });
  if (resp.status === 403 || resp.status === 404) {
    return { ok: false, reason: "forbidden (PAT needs Pages: read+write)" };
  }
  if (!resp.ok) {
    const t = await resp.text();
    return { ok: false, reason: "HTTP " + resp.status + " " + t.slice(0, 80) };
  }
  return { ok: true };
}

function pagesBackoffMin(failures) {
  return Math.min(PAGES_BASE_BACKOFF_MIN * Math.pow(2, Math.max(0, failures - 1)), PAGES_MAX_BACKOFF_MIN);
}

async function ensurePagesDeploy(env, justCommitted) {
  try {
    if (justCommitted) return { action: "skipped", detail: "fresh commit; push-triggered build in flight" };
    const latest = await ghPagesLatest(env);
    if (!latest) return { action: "unknown", detail: "no Pages build info" };
    const status = latest.status;
    if (status === "built") return { action: "none", status, updatedAt: latest.updated_at || null };
    if (status === "queued" || status === "building") {
      return { action: "none", status, updatedAt: latest.updated_at || null, detail: "build in flight" };
    }
    if (status !== "errored") return { action: "unknown", status, updatedAt: latest.updated_at || null };
    // Latest build errored: only recover if it targeted the current HEAD.
    let headSha = null;
    try { headSha = await ghHeadSha(env); } catch (e) { /* fall through to commit-agnostic count */ }
    if (headSha && latest.commit && latest.commit !== headSha) {
      return { action: "none", status, updatedAt: latest.updated_at || null, detail: "newer commit pending" };
    }
    // Count consecutive errored builds for this commit (newest first).
    let failures = 0;
    try {
      const recent = await ghPagesRecent(env, 10);
      for (const b of recent) {
        if (!b || b.status !== "errored") break;
        if (headSha && b.commit && b.commit !== headSha) break;
        failures++;
      }
    } catch (e) { failures = 1; } // list failed: fall back to single-failure backoff
    if (failures < 1) failures = 1;
    if (failures > PAGES_MAX_REBUILDS) {
      return { action: "exhausted", status, updatedAt: latest.updated_at || null, failures };
    }
    const ts = latest.updated_at || latest.created_at;
    const ageMin = ts ? (Date.now() - Date.parse(ts)) / 60000 : Infinity;
    const waitMin = pagesBackoffMin(failures);
    if (ageMin < waitMin) {
      return { action: "waiting", status, updatedAt: latest.updated_at || null, failures, retryInMin: Math.ceil(waitMin - ageMin) };
    }
    const req = await ghRequestPagesBuild(env);
    if (req.ok) return { action: "rebuilt", status, updatedAt: latest.updated_at || null, failures };
    if (req.reason && req.reason.indexOf("forbidden") === 0) {
      return { action: "forbidden", status, updatedAt: latest.updated_at || null, detail: req.reason };
    }
    return { action: "request-failed", status, updatedAt: latest.updated_at || null, detail: req.reason };
  } catch (err) {
    return { action: "unknown", detail: String((err && err.message) || err).slice(0, 120) };
  }
}

/**
 * Split an existing data.js into { data, suffix }.
 * data.js holds two top-level assignments: window.LB_DATA = {...}; ...
 * window.LB_TANKS = {...};. Only the LB_DATA object is parsed (brace-matched
 * from its opening brace with string/escape awareness, so the later LB_TANKS
 * block is never swallowed). suffix is everything after the LB_DATA
 * terminating ";" and is re-appended byte-for-byte, so the LB_TANKS block
 * and its comments survive every sync unchanged.
 */
function splitLbFile(body) {
  const key = "window.LB_DATA";
  const i = body.indexOf(key);
  if (i < 0) throw new Error("Existing data.js is not a window.LB_DATA file.");
  const eq = body.indexOf("=", i + key.length);
  const start = eq < 0 ? -1 : body.indexOf("{", eq);
  if (start < 0) throw new Error("Existing data.js window.LB_DATA assignment has no JSON object.");
  let depth = 0;
  let inStr = false;
  let esc = false;
  let end = -1;
  for (let j = start; j < body.length; j++) {
    const ch = body[j];
    if (inStr) {
      if (esc) esc = false;
      else if (ch === "\\") esc = true;
      else if (ch === '"') inStr = false;
    } else if (ch === '"') {
      inStr = true;
    } else if (ch === "{") {
      depth++;
    } else if (ch === "}") {
      depth--;
      if (depth === 0) { end = j; break; }
    }
  }
  if (end < 0) throw new Error("Existing data.js window.LB_DATA assignment is unterminated.");
  let k = end + 1;
  while (k < body.length && (body[k] === " " || body[k] === "\t" || body[k] === "\r" || body[k] === "\n")) k++;
  if (body[k] !== ";") throw new Error("Existing data.js window.LB_DATA assignment is not terminated with ';'.");
  let data;
  try {
    data = JSON.parse(body.slice(start, end + 1));
  } catch (e) {
    throw new Error("Existing data.js window.LB_DATA is not valid JSON.");
  }
  return { data: data, suffix: body.slice(k + 1) };
}

/* ------------------------------------------------------------------ *
 * Sync
 * ------------------------------------------------------------------ */

function settingsMap(settings) {
  const out = {};
  for (const s of settings) out[s.name] = s.value;
  return out;
}

function numOr(obj, key, fallback) {
  const n = Number(obj[key]);
  return Number.isFinite(n) ? n : fallback;
}

function daysBetween(aISO, bISO) {
  const a = Date.parse(aISO + "T00:00:00Z");
  const b = Date.parse(bISO + "T00:00:00Z");
  if (!Number.isFinite(a) || !Number.isFinite(b)) return NaN;
  return Math.floor((a - b) / 86400000);
}

async function runSync(env) {
  const sheetId = env.SHEET_ID;
  if (!sheetId) throw new Error("SHEET_ID is not configured on the Worker.");
  if (!env.GITHUB_TOKEN) throw new Error("GITHUB_TOKEN is not configured on the Worker.");
  const [matchesGrid, aliasesGrid, seedsGrid, settingsGrid, lbGrid, apGrid, provGrid, inGrid] =
    await Promise.all([
      fetchSheetGrid(sheetId, "Matches"),
      fetchSheetGrid(sheetId, "Aliases"),
      fetchSheetGrid(sheetId, "Seeds"),
      fetchSheetGrid(sheetId, "Settings"),
      fetchSheetGrid(sheetId, "Leaderboard"),
      fetchSheetGrid(sheetId, "All Players"),
      fetchSheetGrid(sheetId, "Provisional"),
      fetchSheetGrid(sheetId, "Inactive"),
    ]);

  const aliases = buildAliases(aliasesGrid);
  const matches = buildMatches(matchesGrid, aliases);
  const seeds = buildSeeds(seedsGrid, aliases);
  const settings = buildSettings(settingsGrid);
  const inactiveList = buildInactive(inGrid, aliases);
  const { qualified, players, provNames } = buildBoardTabs(lbGrid, apGrid, provGrid);

  const cfg = settingsMap(settings);
  const minM = numOr(cfg, "Minimum matches for leaderboard", 5);
  const minO = numOr(cfg, "Minimum different opponents", 3);
  const inactiveDays = numOr(cfg, "Inactive after days", 365);
  const graceDays = numOr(cfg, "Legacy grace days", 365);
  const graceStart = typeof cfg["Legacy grace start date"] === "string" && /^\d{4}-\d{2}-\d{2}/.test(cfg["Legacy grace start date"])
    ? cfg["Legacy grace start date"].slice(0, 10)
    : "2026-10-04";
  const today = todayISO();

  // Provisional: the sheet's own Provisional tab is the verdict, with the
  // thresholds as fallback for names missing from the tab.
  const oppCount = {};
  const matchCount = {};
  for (const m of matches) {
    matchCount[m.a] = (matchCount[m.a] || 0) + 1;
    matchCount[m.b] = (matchCount[m.b] || 0) + 1;
    oppCount[m.a] = oppCount[m.a] || {};
    oppCount[m.a][m.b] = 1;
    oppCount[m.b] = oppCount[m.b] || {};
    oppCount[m.b][m.a] = 1;
  }
  const oppN = (n) => Object.keys(oppCount[n] || {}).length;

  const inactiveSet = new Set(inactiveList.map((n) => n.toLowerCase()));
  const pastThreshold = (p) => {
    if (p.lastMatch) {
      const d = daysBetween(today, p.lastMatch);
      return Number.isFinite(d) && d > inactiveDays;
    }
    const g = daysBetween(today, graceStart);
    return Number.isFinite(g) && g > graceDays;
  };

  for (const p of players) {
    const inProvTab = provNames.has(p.name);
    const failsMin = (matchCount[p.name] || 0) < minM || oppN(p.name) < minO;
    if (inProvTab || failsMin) Object.assign(p, { provisional: true });
    const inTab = inactiveSet.has(p.name.toLowerCase());
    if (inTab) Object.assign(p, { inactive: pastThreshold(p) });
    else if (pastThreshold(p)) Object.assign(p, { inactive: true });
  }

  const histories = buildHistories(matches);

  // Previous-publish snapshots: prev = old current, cur = new ratings.
  const oldFile = await ghGetFile(env);
  const oldSplit = oldFile ? splitLbFile(oldFile.body) : null;
  const old = oldSplit ? oldSplit.data : null;
  const oldCur = (old && old.curRatings) || {};
  if (!Object.keys(oldCur).length && old && Array.isArray(old.players)) {
    for (const p of old.players) {
      if (p && p.name && Number.isFinite(Number(p.rating))) oldCur[p.name] = Number(p.rating);
    }
  }
  const prevRanks = {};
  if (old && Array.isArray(old.qualified)) {
    old.qualified.forEach((q, i) => {
      if (q && q.name) prevRanks[q.name] = Number.isFinite(Number(q.rank)) ? Number(q.rank) : i + 1;
    });
  }
  const curRatings = {};
  for (const p of players) curRatings[p.name] = p.rating;

  // Rebuild data.js preserving key order + every non-sheet section (cards…).
  const next = {};
  const put = (k, v) => { next[k] = v; };
  const oldKeys = old ? Object.keys(old) : [];
  const have = new Set();
  const sheetSections = {
    generated: today,
    qualified,
    players,
    matches,
    aliases: aliases.exact,
    seeds,
    settings,
    inactiveList,
    histories,
    prevRatings: oldCur,
    prevRanks,
    curRatings,
    newSince: (old && old.newSince) || {},
  };
  if (!old) {
    for (const [k, v] of Object.entries(sheetSections)) put(k, v);
  } else {
    for (const k of oldKeys) {
      if (Object.prototype.hasOwnProperty.call(sheetSections, k)) put(k, sheetSections[k]);
      else put(k, old[k]);
      have.add(k);
    }
    for (const [k, v] of Object.entries(sheetSections)) {
      if (!have.has(k)) put(k, v);
    }
  }

  const body =
    "// 1v1 Leaderboard data - generated from Google Sheet (" + today + ")\n" +
    "window.LB_DATA = " + JSON.stringify(next) + ";" +
    (oldSplit ? oldSplit.suffix : "\n");

  if (oldFile && oldFile.body === body) {
    return { changed: false, matches: matches.length, players: players.length };
  }
  const d = new Date();
  const stamp =
    String(d.getUTCFullYear()) + "-" + String(d.getUTCMonth() + 1).padStart(2, "0") + "-" +
    String(d.getUTCDate()).padStart(2, "0") + " " +
    String(d.getUTCHours()).padStart(2, "0") + ":" + String(d.getUTCMinutes()).padStart(2, "0");
  await ghPutFile(env, oldFile ? oldFile.sha : null, body,
    "Sheet sync " + stamp + " (" + matches.length + " matches)");
  return { changed: true, matches: matches.length, players: players.length };
}

// --------------------------------------------------------------------
// Verify (availability identity) - KV-backed challenges, live oracle.
//
// The game exposes player tank state over plain HTTPS JSON-RPC:
//   POST https://tanktrouble.com/ajax/  {jsonrpc, method:
//   "tanktrouble.getPlayerDetailsByUsername", params:[name]}
// No auth needed. Unknown names yield result:false ("No such player").
// A challenge asks the player to change one accessory on their exact
// account; /check polls the oracle until the change appears. The client
// is never trusted: only the oracle decides pass/fail.
//
// Frontend contract (app.js, unchanged):
//   GET  /availability/tank?name=       -> {ok, tank} | {ok:false}
//   POST /availability/verify/start {name}
//     -> {ok, id, player, main, tank, pick, pair, acc, mode, slot,
//             expires, verifiedAlready}
//   POST /availability/verify/check {id}
//     -> {ok, verified, token?, player?, met?, tank?} | {ok:false,
//             error:"expired"}
//   POST /availability/verify/claim {name, token}
//     -> {ok, player} | {ok:false, error:"that code is not right"}
// Error strings match the site's ns() mapper: "expired", "that code is
// not right", "too many verification attempts for this name",
// "no in-game tank found for this name", "could not reach the game
// profile". Requires a KV binding named VERIFY_KV (challenge state,
// claim codes, throttle counters). Without it every verify route answers
// 503 and the normal publish pipeline is unaffected.
// --------------------------------------------------------------------

const VERIFY_CHALLENGE_MS = 15 * 60 * 1000;
const VERIFY_CHAL_KV_S = 20 * 60;
const VERIFY_CODE_S = 90 * 24 * 3600;
const VERIFY_NAME_MAX = 10;
const VERIFY_IP_MAX = 120;
const VERIFY_ACC_SLOTS = [
  "turretAccessory", "barrelAccessory", "frontAccessory",
  "backAccessory", "treadAccessory", "backgroundAccessory",
];
const VERIFY_CODE_ALPHABET = "ABCDEFGHJKMNPQRSTUVWXYZ23456789";

function verifyCors(data, status) {
  const r = json(data, status || 200);
  r.headers.set("Access-Control-Allow-Origin", "*");
  r.headers.set("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  r.headers.set("Access-Control-Allow-Headers", "Content-Type");
  return r;
}

function verifyKv(env) {
  return (env && env.VERIFY_KV) || null;
}

// Bodyless 204 for CORS preflights: a 204 MUST NOT carry a body, and the
// Workers runtime rejects constructing one with a JSON payload (HTTP 500),
// which fails the browser preflight and blocks every verify POST.
function verifyPreflight() {
  return new Response(null, {
    status: 204,
    headers: {
      "Access-Control-Allow-Origin": "*",
      "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
      "Access-Control-Allow-Headers": "Content-Type",
    },
  });
}

function randHex(n) {
  const b = new Uint8Array(n);
  crypto.getRandomValues(b);
  let s = "";
  for (let i = 0; i < b.length; i++) s += b[i].toString(16).padStart(2, "0");
  return s;
}

function randCode() {
  const b = new Uint8Array(8);
  crypto.getRandomValues(b);
  let s = "";
  for (let i = 0; i < b.length; i++) s += VERIFY_CODE_ALPHABET[b[i] % VERIFY_CODE_ALPHABET.length];
  return s;
}

function cleanName(v) {
  const s = String(v == null ? "" : v).trim().replace(/[\u0000-\u001F\u007F]+/g, "");
  if (!s || s.length > 40) return "";
  return s;
}

async function kvBump(kv, key, ttlSeconds, max) {
  const cur = parseInt((await kv.get(key)) || "0", 10) || 0;
  if (cur >= max) return false;
  // Fixed window: first hit sets the TTL, later hits reuse the same window.
  if (cur === 0) await kv.put(key, "1", { expirationTtl: ttlSeconds });
  else await kv.put(key, String(cur + 1));
  return true;
}

async function ttLookup(name) {
  let resp;
  try {
    resp = await fetch("https://tanktrouble.com/ajax/", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ jsonrpc: "2.0", method: "tanktrouble.getPlayerDetailsByUsername", id: 1, params: [name] }),
    });
  } catch (e) {
    return { ok: false, reason: "oracle" };
  }
  let body = null;
  try { body = await resp.json(); } catch (e) { /* non-JSON */ }
  const res = body && body.result;
  if (!resp.ok || !body || body.jsonrpc !== "2.0" || !res) return { ok: false, reason: "oracle" };
  if (!res.result || !res.data) return { ok: false, reason: "unknown" };
  return { ok: true, username: String(res.data.username || name), data: res.data };
}

function tankSubset(d) {
  const pick = (k) => (d && d[k] !== undefined ? d[k] : (k.slice(-6) === "Colour" ? null : "0"));
  const acc = (k) => {
    const v = d ? d[k] : undefined;
    return v === undefined || v === null || v === "" ? "0" : String(v);
  };
  return {
    turretColour: pick("turretColour"),
    treadColour: pick("treadColour"),
    baseColour: pick("baseColour"),
    turretAccessory: acc("turretAccessory"),
    barrelAccessory: acc("barrelAccessory"),
    frontAccessory: acc("frontAccessory"),
    backAccessory: acc("backAccessory"),
  };
}

function tankFingerprint(t) {
  return JSON.stringify([
    t.turretColour, t.treadColour, t.baseColour, t.turretAccessory,
    t.barrelAccessory, t.frontAccessory, t.backAccessory,
  ]);
}

function accVal(t, slot) {
  const v = t ? t[slot] : undefined;
  return v === undefined || v === null || v === "" ? "0" : String(v);
}

async function verifyHasLiveCode(kv, canonical) {
  try {
    const list = await kv.list({ prefix: "vcode:", limit: 50 });
    for (const k of list.keys || []) {
      const v = await kv.get(k.name, "json");
      if (v && String(v.player || "").toLowerCase() === canonical.toLowerCase()) return true;
    }
  } catch (e) { /* treat as none */ }
  return false;
}

async function handleTank(env, url) {
  const kv = verifyKv(env);
  if (!kv) return verifyCors({ ok: false, error: "verification backend not configured" }, 503);
  const name = cleanName(url.searchParams.get("name"));
  if (!name) return verifyCors({ ok: false, error: "no in-game tank found for this name" });
  const r = await ttLookup(name);
  if (!r.ok) {
    return verifyCors({ ok: false, error: r.reason === "unknown" ? "no in-game tank found for this name" : "could not reach the game profile" });
  }
  const tank = tankSubset(r.data);
  // lastLogin is the one owner-approved extra field: the cards render a
  // "last seen" line from it, and the baked data.js block never refreshes.
  if (r.data && r.data.lastLogin != null) tank.lastLogin = String(r.data.lastLogin);
  return verifyCors({ ok: true, player: r.username, tank });
}

async function handleVerifyStart(env, request) {
  const kv = verifyKv(env);
  if (!kv) return verifyCors({ ok: false, error: "verification backend not configured" }, 503);
  let body = {};
  try { body = await request.json(); } catch (e) { /* empty/invalid body */ }
  const name = cleanName(body && body.name);
  if (!name) return verifyCors({ ok: false, error: "no in-game tank found for this name" });
  const ip = request.headers.get("CF-Connecting-IP") || "unknown";
  if (!(await kvBump(kv, "vrl:ip:" + ip, 60, VERIFY_IP_MAX))) {
    return verifyCors({ ok: false, error: "too many attempts - slow down" }, 429);
  }
  const r = await ttLookup(name);
  if (!r.ok) {
    return verifyCors({ ok: false, error: r.reason === "unknown" ? "no in-game tank found for this name" : "could not reach the game profile" });
  }
  const canonical = r.username;
  if (!(await kvBump(kv, "vrl:name:" + canonical.toLowerCase(), 3600, VERIFY_NAME_MAX))) {
    return verifyCors({ ok: false, error: "too many verification attempts for this name" }, 429);
  }
  const tank = tankSubset(r.data);
  let kind = "any";
  let slot = "";
  let original = "";
  let snapshot = "";
  const equipped = VERIFY_ACC_SLOTS.filter((s) => accVal(tank, s) !== "0");
  if (equipped.length) {
    // Take-off-or-swap challenge on a slot the account actually uses.
    kind = "slot";
    slot = equipped[Math.floor(Math.random() * equipped.length)];
    original = accVal(tank, slot);
  } else {
    snapshot = tankFingerprint(tank);
  }
  const id = randHex(16);
  const expires = Date.now() + VERIFY_CHALLENGE_MS;
  await kv.put("vchal:" + id, JSON.stringify({ player: canonical, kind, slot, original, snapshot, expires }), { expirationTtl: VERIFY_CHAL_KV_S });
  const verifiedAlready = await verifyHasLiveCode(kv, canonical);
  if (kind === "slot") {
    return verifyCors({
      ok: true, id, player: canonical, main: "", tank,
      pick: null, pair: [], acc: null, mode: "slot", slot,
      expires, verifiedAlready,
    });
  }
  return verifyCors({
    ok: true, id, player: canonical, main: "", tank,
    pick: null, pair: [], acc: null, mode: "any", slot: "",
    expires, verifiedAlready,
  });
}

async function handleVerifyCheck(env, request) {
  const kv = verifyKv(env);
  if (!kv) return verifyCors({ ok: false, error: "verification backend not configured" }, 503);
  let body = {};
  try { body = await request.json(); } catch (e) { /* empty/invalid body */ }
  const id = String((body && body.id) || "");
  if (!/^[0-9a-f]{32}$/.test(id)) return verifyCors({ ok: false, error: "expired" });
  const raw = await kv.get("vchal:" + id);
  if (!raw) return verifyCors({ ok: false, error: "expired" });
  let chal = null;
  try { chal = JSON.parse(raw); } catch (e) { /* corrupt */ }
  if (!chal || !chal.player || !chal.expires) return verifyCors({ ok: false, error: "expired" });
  if (Date.now() > chal.expires) {
    await kv.delete("vchal:" + id);
    return verifyCors({ ok: false, error: "expired" });
  }
  const ip = request.headers.get("CF-Connecting-IP") || "unknown";
  if (!(await kvBump(kv, "vrl:ip:" + ip, 60, VERIFY_IP_MAX))) {
    return verifyCors({ ok: false, error: "too many attempts - slow down" }, 429);
  }
  const r = await ttLookup(chal.player);
  if (!r.ok) {
    return verifyCors({ ok: false, error: "could not reach the game profile" });
  }
  const tank = tankSubset(r.data);
  let met = false;
  if (chal.kind === "slot" && chal.slot) {
    met = accVal(tank, chal.slot) !== String(chal.original);
  } else {
    met = tankFingerprint(tank) !== String(chal.snapshot);
  }
  if (!met) {
    return verifyCors({ ok: true, verified: false, met: { acc: false }, tank });
  }
  const token = randCode();
  await kv.put("vcode:" + token, JSON.stringify({ player: r.username }), { expirationTtl: VERIFY_CODE_S });
  await kv.delete("vchal:" + id);
  return verifyCors({ ok: true, verified: true, token, player: r.username });
}

async function handleVerifyClaim(env, request) {
  const kv = verifyKv(env);
  if (!kv) return verifyCors({ ok: false, error: "verification backend not configured" }, 503);
  let body = {};
  try { body = await request.json(); } catch (e) { /* empty/invalid body */ }
  const name = cleanName(body && body.name);
  const token = String((body && body.token) || "").toUpperCase().replace(/[^A-Z0-9]/g, "");
  if (!name || !token) return verifyCors({ ok: false, error: "that code is not right" });
  const v = await kv.get("vcode:" + token, "json");
  if (!v || !v.player || String(v.player).toLowerCase() !== name.toLowerCase()) {
    return verifyCors({ ok: false, error: "that code is not right" });
  }
  try { await kv.put("vcode:" + token, JSON.stringify(v), { expirationTtl: CLAIM_TTL_S }); } catch (e) { /* ttl refresh is best effort */ }
  try { await kv.put("av:vowner:" + name.toLowerCase(), JSON.stringify({ at: Date.now() })); } catch (e) { /* optional registry */ }
  return verifyCors({ ok: true, player: v.player });
}

/* ------------------------------------------------------------------ *
 * Router
 * ------------------------------------------------------------------ */

async function handleSyncStatus(env) {
  try {
    const file = await ghGetFile(env);
    if (!file) return json({ error: "data.js has not been published yet." }, 404);
    const lb = splitLbFile(file.body).data;
    const iso = await ghLastCommitDate(env);
    const ageMinutes = iso ? Math.max(0, Math.round((Date.now() - Date.parse(iso)) / 60000)) : null;
    const pages = await ensurePagesDeploy(env, false);
    let lastError = null;
    if (pages.action === "exhausted") {
      lastError = "Pages deployment failed " + (pages.failures || PAGES_MAX_REBUILDS) +
        "x for the current commit; manual rebuild needed (Cloudflare/Worker logs, then POST /repos/.../pages/builds).";
    } else if (pages.action === "forbidden") {
      lastError = "Pages rebuild not permitted: PAT needs Pages read+write (" + (pages.detail || "forbidden") + ").";
    }
    return json({
      matches: Array.isArray(lb.matches) ? lb.matches.length : 0,
      players: Array.isArray(lb.players) ? lb.players.length : 0,
      ageMinutes,
      stale: ageMinutes === null ? true : ageMinutes > STALE_MINUTES,
      lastError,
      generated: lb.generated || null,
      pagesDeploy: {
        action: pages.action,
        status: pages.status || null,
        updatedAt: pages.updatedAt || null,
      },
    });
  } catch (err) {
    return json({ error: String((err && err.message) || err).slice(0, 200) }, 500);
  }
}


// ---------------------------------------------------------------------------
// TankTrouble admission token for the ping test.
// Game servers only accept sockets whose ?t= carries a token minted by
// tanktrouble.com and injected into their own HTML, which sends no CORS
// headers, so the browser cannot read it directly. This proxies that string.
// ---------------------------------------------------------------------------
const TT_TOKEN_RE = /g_admissionToken\s*=\s*"([^"]+)"/;
const TT_TOKEN_EXP_RE = /g_admissionTokenExpiresAt\s*=\s*(\d+)/;

async function ttAdmissionToken(request) {
  let html = "";
  try {
    const r = await fetch("https://tanktrouble.com/", {
      headers: { "Accept": "text/html", "User-Agent": "Mozilla/5.0 (compatible; TierStats/1.0)" },
      cf: { cacheTtl: 45, cacheEverything: true },
    });
    if (!r.ok) return verifyCors({ error: "TankTrouble responded " + r.status }, 502);
    html = await r.text();
  } catch (err) {
    return verifyCors({ error: "Could not reach TankTrouble" }, 502);
  }
  const m = html.match(TT_TOKEN_RE);
  if (!m) return verifyCors({ error: "No admission token in TankTrouble page" }, 502);
  const exp = html.match(TT_TOKEN_EXP_RE);
  return verifyCors({
    token: m[1],
    expiresAt: exp ? asInt(exp[1]) : 0,
    country: (request.cf && request.cf.country) || "",
    network: null,
    floors: null,
    nonce: "",
  });
}

// The VPN/network verdict is not implemented; say so instead of guessing.
function ttAdmissionCheck(request) {
  return verifyCors({
    ok: true,
    verdict: "unchecked",
    allowEntry: false,
    review: false,
    reason: "Network check unavailable",
    token: newMeasureToken(),
    country: (request.cf && request.cf.country) || "",
    network: null,
  });
}


// ---------------------------------------------------------------------------
// Profile ping store (/profiles/*).
// Deliberately independent of tournament network validation: a verified
// player can measure and publish their server latency without the unfinished
// VPN / impossible-RTT check. Auth is the verification code they claimed
// their name with (stored as vcode:<CODE> in VERIFY_KV).
// ---------------------------------------------------------------------------
const PING_KEY_PREFIX = "ping:";
const PING_TOKEN_RE = /^m\.[a-f0-9]{32}\.\d{10,}$/;
const PING_TOKEN_TTL_MS = 30 * 60 * 1000;
const PING_VALUE_RE = /^\d{1,4}(?:-\d{1,4})?$/;
// A claimed code doubles as the long-lived profile credential in the browser,
// so every successful use refreshes its KV lifetime.
const CLAIM_TTL_S = 400 * 24 * 3600;

function newMeasureToken() {
  const b = new Uint8Array(16);
  crypto.getRandomValues(b);
  let h = "";
  for (let i = 0; i < b.length; i++) h += b[i].toString(16).padStart(2, "0");
  return "m." + h + "." + Date.now();
}

function measureTokenOk(tok) {
  const t = String(tok || "");
  if (!PING_TOKEN_RE.test(t)) return false;
  const ts = Number(t.split(".")[2]);
  if (!Number.isFinite(ts)) return false;
  const age = Date.now() - ts;
  return age >= 0 && age <= PING_TOKEN_TTL_MS;
}

async function profileAuth(env, name, claim) {
  const kv = verifyKv(env);
  if (!kv) return { error: verifyCors({ ok: false, error: "profile backend not configured" }, 503) };
  const player = cleanName(name);
  const code = String(claim || "").toUpperCase().replace(/[^A-Z0-9]/g, "");
  if (!player) return { error: verifyCors({ ok: false, error: "name required" }, 400) };
  if (!code) return { error: verifyCors({ ok: false, error: "claim required" }, 403) };
  const v = await kv.get("vcode:" + code, "json");
  if (!v || !v.player || String(v.player).toLowerCase() !== player.toLowerCase()) {
    return { error: verifyCors({ ok: false, error: "claim expired" }, 403) };
  }
  try { await kv.put("vcode:" + code, JSON.stringify(v), { expirationTtl: CLAIM_TTL_S }); } catch (e) { /* ttl refresh is best effort */ }
  return {
    kv,
    player,
    record: await kv.get(PING_KEY_PREFIX + player.toLowerCase(), "json"),
  };
}

function cleanPingMap(raw) {
  const out = {};
  if (!raw || typeof raw !== "object") return out;
  for (const server of SERVERS) {
    const value = String(raw[server] === undefined || raw[server] === null ? "" : raw[server]).trim();
    if (PING_VALUE_RE.test(value)) out[server] = value;
  }
  return out;
}

async function handleProfilePing(env, request) {
  let body = {};
  try { body = await request.json(); } catch (e) { /* empty/invalid body */ }
  const auth = await profileAuth(env, body && body.name, body && body.claim);
  if (auth.error) return auth.error;
  if (!measureTokenOk(body && body.token)) {
    return verifyCors({ ok: false, error: "measurement expired - measure again" }, 403);
  }
  const pings = cleanPingMap(body && body.pings);
  const main = SERVERS.indexOf(String((body && body.main) || "")) >= 0 ? String(body.main) : "";
  if (!main || !Object.keys(pings).length) return verifyCors({ ok: false, error: "no ping data" }, 400);
  const record = {
    name: auth.player,
    main,
    pings,
    samples: body && body.samples && typeof body.samples === "object" ? body.samples : {},
    at: Date.now(),
    flagHidden: !!(body && body.flagHidden !== undefined ? body.flagHidden : auth.record && auth.record.flagHidden),
  };
  await auth.kv.put(PING_KEY_PREFIX + auth.player.toLowerCase(), JSON.stringify(record));
  try { await auth.kv.put("av:vowner:" + auth.player.toLowerCase(), JSON.stringify({ at: Date.now() })); } catch (e) { /* optional registry */ }
  return verifyCors({ ok: true, record });
}

async function handleProfilePings(env) {
  const kv = verifyKv(env);
  if (!kv) return verifyCors({ ok: false, error: "profile backend not configured" }, 503);
  const pings = {};
  let cursor;
  do {
    const page = await kv.list({ prefix: PING_KEY_PREFIX, cursor });
    for (const key of page.keys) {
      const rec = await kv.get(key.name, "json");
      if (rec && rec.name && rec.main && rec.pings) pings[rec.name] = rec;
    }
    cursor = page.list_complete ? undefined : page.cursor;
  } while (cursor);
  return verifyCors({ ok: true, pings });
}

async function handleProfileFlag(env, request) {
  let body = {};
  try { body = await request.json(); } catch (e) { /* empty/invalid body */ }
  const auth = await profileAuth(env, body && body.name, body && body.claim);
  if (auth.error) return auth.error;
  if (!auth.record) return verifyCors({ ok: false, error: "no ping saved yet" }, 404);
  auth.record.flagHidden = !!(body && body.hidden);
  await auth.kv.put(PING_KEY_PREFIX + auth.player.toLowerCase(), JSON.stringify(auth.record));
  return verifyCors({ ok: true, flagHidden: auth.record.flagHidden });
}

async function handleProfileReset(env, request) {
  let body = {};
  try { body = await request.json(); } catch (e) { /* empty/invalid body */ }
  const auth = await profileAuth(env, body && body.name, body && body.claim);
  if (auth.error) return auth.error;
  await auth.kv.delete(PING_KEY_PREFIX + auth.player.toLowerCase());
  return verifyCors({ ok: true });
}


// ---------------------------------------------------------------------------
// Availability board ("post a 1v1"). KV-backed slots, one live post per
// player, with a per-post secret for edit/delete and claim-based protection
// of verified names.
//   GET  /availability            -> {ok, now, slots, owners}
//   POST /availability            -> create: {ok, slot, secret}
//   POST /availability/update     -> {id, secret, start, end}
//   POST /availability/delete     -> {id, secret, claim, start}
//   POST /availability/interest   -> {id, name, claim, on}
// ---------------------------------------------------------------------------
const AV_SLOT_PREFIX = "av:id:";
const AV_OWNER_PREFIX = "av:owner:";
const AV_VERIFIED_PREFIX = "av:vowner:";
const AV_ROSTER_KEY = "av:roster";
const AV_NOW_MS = 30 * 60 * 1000;
const AV_MAX_SPAN_MS = 6 * 3600 * 1000;
const AV_MAX_AHEAD_MS = 7 * 24 * 3600 * 1000;
const AV_GRACE_MS = 5 * 60 * 1000;
const AV_ROSTER_TTL_MS = 5 * 60 * 1000;
const AV_SLOT_TTL_S = 7 * 24 * 3600;
const AV_MAX_POSTS_PER_HOUR = 12;

function avHex(bytes) {
  const b = new Uint8Array(bytes);
  crypto.getRandomValues(b);
  let h = "";
  for (let i = 0; i < b.length; i++) h += b[i].toString(16).padStart(2, "0");
  return h;
}

function avPublic(slot) {
  return {
    id: slot.id,
    player: slot.player,
    main: slot.main,
    type: slot.type,
    servers: slot.servers || [],
    zone: slot.zone || "",
    start: slot.start,
    end: slot.end,
    verified: !!slot.verified,
    interest: Array.isArray(slot.interest) ? slot.interest : [],
  };
}

async function avClaimOk(kv, name, claim) {
  if (!kv) return false;
  const code = String(claim || "").toUpperCase().replace(/[^A-Z0-9]/g, "");
  if (!code) return false;
  const v = await kv.get("vcode:" + code, "json");
  if (!v || !v.player) return false;
  return String(v.player).toLowerCase() === String(name).toLowerCase();
}

async function avRoster(env) {
  const kv = verifyKv(env);
  if (kv) {
    try {
      const cached = await kv.get(AV_ROSTER_KEY, "json");
      if (cached && Array.isArray(cached.names) && cached.names.length && cached.at > Date.now() - AV_ROSTER_TTL_MS) {
        return cached.names;
      }
    } catch (e) { /* fall through to the published leaderboard */ }
  }
  let file = null;
  try { file = await ghGetFile(env); } catch (e) { file = null; }
  if (!file) return null;
  const lb = splitLbFile(file.body).data || {};
  const seen = new Map();
  for (const p of lb.players || []) {
    if (p && p.name) seen.set(String(p.name).toLowerCase(), String(p.name));
  }
  const alts = (lb.cards && lb.cards.alts) || {};
  for (const [base, list] of Object.entries(alts)) {
    const canon = seen.get(String(base).toLowerCase());
    if (!canon) continue;
    for (const alt of Array.isArray(list) ? list : []) {
      if (alt) seen.set(String(alt).toLowerCase(), canon);
    }
  }
  const names = [...seen.values()];
  if (kv && names.length) {
    try { await kv.put(AV_ROSTER_KEY, JSON.stringify({ at: Date.now(), names }), { expirationTtl: 3600 }); } catch (e) { /* cache only */ }
  }
  return names;
}

async function avRateOk(kv, request) {
  const ip = request.headers.get("CF-Connecting-IP") || request.headers.get("X-Forwarded-For") || "local";
  const key = "av:rl:" + String(ip).split(",")[0].trim();
  const now = Date.now();
  const cur = (await kv.get(key, "json")) || { n: 0, at: now };
  if (cur.at < now - 3600 * 1000) { cur.n = 0; cur.at = now; }
  if (cur.n >= AV_MAX_POSTS_PER_HOUR) return false;
  cur.n += 1;
  try { await kv.put(key, JSON.stringify(cur), { expirationTtl: 3600 }); } catch (e) { /* best effort */ }
  return true;
}

async function avSlot(kv, id) {
  if (!id) return null;
  try { return await kv.get(AV_SLOT_PREFIX + String(id), "json"); } catch (e) { return null; }
}

async function handleAvBoard(env) {
  const kv = verifyKv(env);
  if (!kv) return verifyCors({ ok: false, error: "availability backend not configured" }, 503);
  const now = Date.now();
  const slots = [];
  let cursor;
  do {
    const page = await kv.list({ prefix: AV_SLOT_PREFIX, cursor });
    for (const key of page.keys) {
      const rec = await kv.get(key.name, "json");
      if (rec && rec.id && rec.player && Number(rec.end) > now - AV_GRACE_MS) slots.push(avPublic(rec));
    }
    cursor = page.list_complete ? undefined : page.cursor;
  } while (cursor);
  slots.sort((a, b) => a.start - b.start || String(a.player).localeCompare(String(b.player)));
  const owners = [...new Set(slots.filter((s) => s.verified).map((s) => s.player))];
  return verifyCors({ ok: true, now, slots, owners });
}

async function handleAvCreate(env, request) {
  const kv = verifyKv(env);
  if (!kv) return verifyCors({ ok: false, error: "availability backend not configured" }, 503);
  let b = {};
  try { b = await request.json(); } catch (e) { /* empty/invalid body */ }
  const type = b && b.type === "later" ? "later" : "now";
  const asked = cleanName(b && b.name);
  if (!asked) return verifyCors({ ok: false, error: "name required" }, 400);
  const roster = await avRoster(env);
  if (!roster) return verifyCors({ ok: false, error: "leaderboard is not reachable" }, 503);
  const canon = roster.find((n) => String(n).toLowerCase() === asked.toLowerCase());
  if (!canon) return verifyCors({ ok: false, error: "not on the leaderboard" }, 400);
  const key = canon.toLowerCase();
  const now = Date.now();
  const existing = await kv.get(AV_OWNER_PREFIX + key, "json");
  if (existing && Number(existing.end) > now) {
    return verifyCors({ ok: false, error: existing.type === "later" ? "scheduled availability" : "right-now availability" }, 409);
  }
  const claimed = await avClaimOk(kv, canon, b && b.claim);
  const registered = await kv.get(AV_VERIFIED_PREFIX + key, "json");
  if (registered && !claimed) return verifyCors({ ok: false, error: "verified-owner" }, 403);
  const servers = (Array.isArray(b && b.servers) ? b.servers : [])
    .filter((v) => SERVERS.indexOf(String(v)) >= 0)
    .filter((v, i, all) => all.indexOf(v) === i)
    .slice(0, SERVERS.length);
  if (!servers.length) return verifyCors({ ok: false, error: "pick at least one server" }, 400);
  let start;
  let end;
  if (type === "now") {
    start = now;
    end = now + AV_NOW_MS;
  } else {
    start = Number(b && b.start);
    end = Number(b && b.end);
    if (!Number.isFinite(start) || !Number.isFinite(end)) return verifyCors({ ok: false, error: "pick a start and end time" }, 400);
    if (start < now - AV_GRACE_MS) return verifyCors({ ok: false, error: "start time is already in the past" }, 400);
    if (start > now + AV_MAX_AHEAD_MS) return verifyCors({ ok: false, error: "start time is too far ahead" }, 400);
    if (end <= start) return verifyCors({ ok: false, error: "end time must be after the start" }, 400);
    if (end - start > AV_MAX_SPAN_MS) return verifyCors({ ok: false, error: "that window is too long" }, 400);
  }
  if (!(await avRateOk(kv, request))) return verifyCors({ ok: false, error: "too many posts" }, 429);
  const slot = {
    id: "av_" + avHex(6),
    player: canon,
    main: servers[0],
    type,
    servers,
    zone: String((b && b.zone) || "").slice(0, 64),
    start,
    end,
    verified: !!claimed,
    interest: [],
    createdAt: now,
    secret: avHex(16),
  };
  try {
    await kv.put(AV_SLOT_PREFIX + slot.id, JSON.stringify(slot), { expirationTtl: AV_SLOT_TTL_S });
    await kv.put(AV_OWNER_PREFIX + key, JSON.stringify({ id: slot.id, end, type }), { expirationTtl: AV_SLOT_TTL_S });
    if (claimed) await kv.put(AV_VERIFIED_PREFIX + key, JSON.stringify({ at: now }));
  } catch (e) {
    return verifyCors({ ok: false, error: "could not save that post" }, 500);
  }
  return verifyCors({ ok: true, slot: avPublic(slot), secret: slot.secret });
}

async function handleAvUpdate(env, request) {
  const kv = verifyKv(env);
  if (!kv) return verifyCors({ ok: false, error: "availability backend not configured" }, 503);
  let b = {};
  try { b = await request.json(); } catch (e) { /* empty/invalid body */ }
  const rec = await avSlot(kv, b && b.id);
  if (!rec) return verifyCors({ ok: false, error: "that post is gone" }, 404);
  if (!b.secret || b.secret !== rec.secret) return verifyCors({ ok: false, error: "that post is not yours" }, 403);
  const start = Number(b.start) || rec.start;
  const end = Number(b.end);
  if (!Number.isFinite(end) || end <= start) return verifyCors({ ok: false, error: "end time must be after the start" }, 400);
  if (end - start > AV_MAX_SPAN_MS) return verifyCors({ ok: false, error: "that window is too long" }, 400);
  if (end < Date.now()) return verifyCors({ ok: false, error: "that window has already ended" }, 400);
  rec.start = start;
  rec.end = end;
  await kv.put(AV_SLOT_PREFIX + rec.id, JSON.stringify(rec), { expirationTtl: AV_SLOT_TTL_S });
  const key = String(rec.player).toLowerCase();
  await kv.put(AV_OWNER_PREFIX + key, JSON.stringify({ id: rec.id, end, type: rec.type }), { expirationTtl: AV_SLOT_TTL_S });
  return verifyCors({ ok: true, slot: avPublic(rec) });
}

async function handleAvDelete(env, request) {
  const kv = verifyKv(env);
  if (!kv) return verifyCors({ ok: false, error: "availability backend not configured" }, 503);
  let b = {};
  try { b = await request.json(); } catch (e) { /* empty/invalid body */ }
  const rec = await avSlot(kv, b && b.id);
  if (!rec) return verifyCors({ ok: true });
  const secretOk = !!b.secret && b.secret === rec.secret;
  const claimOk = await avClaimOk(kv, rec.player, b && b.claim);
  if (!secretOk && !claimOk) return verifyCors({ ok: false, error: "that post is not yours" }, 403);
  await kv.delete(AV_SLOT_PREFIX + rec.id);
  const key = String(rec.player).toLowerCase();
  const owner = await kv.get(AV_OWNER_PREFIX + key, "json");
  if (owner && owner.id === rec.id) await kv.delete(AV_OWNER_PREFIX + key);
  return verifyCors({ ok: true });
}

async function handleAvInterest(env, request) {
  const kv = verifyKv(env);
  if (!kv) return verifyCors({ ok: false, error: "availability backend not configured" }, 503);
  let b = {};
  try { b = await request.json(); } catch (e) { /* empty/invalid body */ }
  const rec = await avSlot(kv, b && b.id);
  if (!rec) return verifyCors({ ok: false, error: "that post is gone" }, 404);
  const who = String((b && b.name) || "").trim();
  if (!who) return verifyCors({ ok: false, error: "name required" }, 400);
  const roster = await avRoster(env);
  if (!roster) return verifyCors({ ok: false, error: "leaderboard is not reachable" }, 503);
  const canon = roster.find((n) => String(n).toLowerCase() === who.toLowerCase());
  if (!canon) return verifyCors({ ok: false, error: "not on the leaderboard" }, 400);
  const claimed = await avClaimOk(kv, canon, b && b.claim);
  if (b && b.claim && !claimed) return verifyCors({ ok: false, error: "claim expired" }, 403);
  const registered = await kv.get(AV_VERIFIED_PREFIX + String(canon).toLowerCase(), "json");
  if (registered && !claimed) return verifyCors({ ok: false, error: "verified-owner" }, 403);
  const key = String(canon).toLowerCase();
  if (Number(rec.end) <= Date.now()) return verifyCors({ ok: false, error: "that post is gone" }, 404);
  const list = (Array.isArray(rec.interest) ? rec.interest : []).filter((n) => String(n).toLowerCase() !== key);
  if (b.on) list.push(canon);
  rec.interest = list.slice(0, 200);
  await kv.put(AV_SLOT_PREFIX + rec.id, JSON.stringify(rec), { expirationTtl: AV_SLOT_TTL_S });
  return verifyCors({ ok: true, interest: rec.interest, slot: avPublic(rec) });
}

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    const path = url.pathname.replace(/\/+$/, "") || "/";
    if (request.method === "GET" && (path === "/" || path === "")) {
      return new Response("tierstats sync worker ok\n", {
        headers: { "Content-Type": "text/plain", "Cache-Control": "no-store" },
      });
    }
    if (path === "/sync-status" && request.method === "GET") {
      return handleSyncStatus(env);
    }
    if (url.pathname === "/availability" && request.method === "OPTIONS") {
      return verifyPreflight();
    }
    if (url.pathname === "/availability" && request.method === "GET") {
      return handleAvBoard(env);
    }
    if (url.pathname === "/availability" && request.method === "POST") {
      return handleAvCreate(env, request);
    }
    if (request.method === "OPTIONS" && ["/availability/update", "/availability/delete", "/availability/interest"].indexOf(url.pathname) >= 0) {
      return verifyPreflight();
    }
    if (url.pathname === "/availability/update" && request.method === "POST") {
      return handleAvUpdate(env, request);
    }
    if (url.pathname === "/availability/delete" && request.method === "POST") {
      return handleAvDelete(env, request);
    }
    if (url.pathname === "/availability/interest" && request.method === "POST") {
      return handleAvInterest(env, request);
    }
    if (url.pathname === "/availability/tank" && request.method === "OPTIONS") {
      return verifyPreflight();
    }
    if (request.method === "OPTIONS" && url.pathname.indexOf("/availability/verify/") === 0) {
      return verifyPreflight();
    }
    if (url.pathname === "/availability/tank" && request.method === "GET") {
      return handleTank(env, url);
    }
    if (url.pathname === "/availability/verify/start" && request.method === "POST") {
      return handleVerifyStart(env, request);
    }
    if (url.pathname === "/availability/verify/check" && request.method === "POST") {
      return handleVerifyCheck(env, request);
    }
    if (url.pathname === "/availability/verify/claim" && request.method === "POST") {
      return handleVerifyClaim(env, request);
    }
    if (url.pathname.indexOf("/tt-admission") === 0 && request.method === "OPTIONS") {
      return verifyPreflight();
    }
    if (path === "/tt-admission" && request.method === "GET") {
      return ttAdmissionToken(request);
    }
    if (path === "/tt-admission/check" && request.method === "POST") {
      return ttAdmissionCheck(request);
    }
    if (path === "/tt-admission/verify" && request.method === "POST") {
      let body = {};
      try { body = await request.json(); } catch (e) { /* empty/invalid body */ }
      const tok = String((body && body.token) || "");
      if (!measureTokenOk(tok)) {
        return verifyCors({ ok: false, allow: false, error: "measurement expired - measure again" }, 403);
      }
      return verifyCors({ ok: true, allow: true, verdict: "unchecked", review: false, reason: "" });
    }
    if (url.pathname.indexOf("/profiles/") === 0 && request.method === "OPTIONS") {
      return verifyPreflight();
    }
    if (path === "/profiles/pings" && request.method === "GET") {
      return handleProfilePings(env);
    }
    if ((path === "/profiles/ping" || path === "/profiles/recover") && request.method === "POST") {
      return handleProfilePing(env, request);
    }
    if (path === "/profiles/flag" && request.method === "POST") {
      return handleProfileFlag(env, request);
    }
    if (path === "/profiles/reset" && request.method === "POST") {
      return handleProfileReset(env, request);
    }
    if (path === "/sync" && request.method === "POST") {
      let body = {};
      try { body = await request.json(); } catch (e) { /* empty/invalid body */ }
      if (!secretsEqual(body && body.password, env.WORKER_SECRET)) {
        return json({ error: "Unauthorized." }, 401);
      }
      try {
        const r = await runSync(env);
        const pages = await ensurePagesDeploy(env, r.changed);
        let note = r.changed
          ? r.matches + " matches / " + r.players + " players published."
          : "Site already matched the sheet.";
        if (pages.action === "rebuilt") note += " Pages redeploy requested.";
        else if (pages.action === "exhausted") note += " WARNING: Pages deployment keeps failing; manual rebuild needed.";
        else if (pages.action === "forbidden") note += " NOTE: Pages rebuild not permitted (PAT needs Pages read+write).";
        return json({
          ok: true,
          changed: r.changed,
          matches: r.matches,
          players: r.players,
          note,
          pagesDeploy: pages.action,
        });
      } catch (err) {
        return json({ error: String((err && err.message) || err).slice(0, 200) }, 500);
      }
    }
    return json({ error: "Not found." }, 404);
  },

  async scheduled(event, env, ctx) {
    try {
      const r = await runSync(env);
      const pages = await ensurePagesDeploy(env, r.changed);
      console.log("cron sync: changed=" + r.changed + " matches=" + r.matches + " players=" + r.players +
        " pages=" + pages.action + (pages.status ? "(" + pages.status + ")" : ""));
    } catch (err) {
      console.log("cron sync failed: " + String((err && err.message) || err).slice(0, 200));
    }
  },
};
