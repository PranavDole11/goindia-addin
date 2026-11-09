/* global CustomFunctions, console */

// ── Custom Functions ──
// PING/WHOAMI/RECALCPROBE/GETTABLE below are leftover Phase 0 spike probes (identity
// propagation, recalculation behavior, dynamic arrays) and are harmless to leave in place.
// GET() is the real thing: =GIA.GET(company, metric, period) — a user should be able to type
// this by hand with just those 3 values, without knowing which real table/column a metric
// lives on. So resolution (metric name -> table/column, annual vs. quarterly, the period's
// literal stored value) happens HERE, inside the function itself, the one time it's actually
// needed — not pre-baked by a picker UI. `table`/`column`/`periodLiteral` exist only as
// optional power-user overrides (e.g. for a picker UI that already knows the answer and wants
// to skip the lookup, or to resolve a genuinely ambiguous metric by hand) — see resolveGetArgs.
//
// Imports METRIC_REGISTRY directly (the same JSON src/taskpane/taskpane.js's own Table Builder
// reads) but does NOT import any of taskpane.js's own JS — different webpack entry/bundle, and
// its resolution machinery (buildMetricRegistry/resolveStandardMetric/loadPeriodJoinMetadata)
// lives nested inside applyRibbonView's own view-specific closures anyway, not in a scope this
// file could reach even if it tried. This resolver is a smaller, self-contained parallel: it
// trusts METRIC_REGISTRY's own "prefer" table directly rather than cross-checking ties across
// the ENTIRE live catalog the way the regular Table Builder does, but still fetches and parses
// that table's real schema doc before ever building SQL, so a stale/wrong registry candidate
// surfaces as a clear GET_ERROR instead of a silently wrong query.
import METRIC_REGISTRY from "../taskpane/metric-registry.json";

/**
 * SPIKE ONLY. Proves the Custom Functions manifest/webpack wiring works at all — no network,
 * no async, nothing that could fail for any reason other than "the plumbing is broken."
 * @customfunction
 * @returns {string} A static string.
 */
function ping() {
    return "pong";
}

/**
 * SPIKE ONLY. Reads the SAME localStorage key the task pane's own getWalletIdentity() reads
 * (see taskpane.js:1079) to prove — or disprove — that this runtime shares the task pane's
 * identity. Log into the add-in via the task pane FIRST, then type this formula in the same
 * Excel session.
 * @customfunction
 * @returns {string} The logged-in user's id, or "NO_LOCAL_STORAGE_USER" if nothing is there.
 */
function whoami() {
    try {
        const raw = localStorage.getItem("user");
        if (!raw) return "NO_LOCAL_STORAGE_USER";
        const user = JSON.parse(raw);
        return user && user.UserId ? String(user.UserId) : "NO_LOCAL_STORAGE_USER";
    } catch (e) {
        return `WHOAMI_ERROR: ${e.message}`;
    }
}

/**
 * SPIKE ONLY. No network call — a free probe for how often Excel actually re-invokes a plain
 * async custom function. Put it in a cell, then: save+reopen the workbook, edit an unrelated
 * cell elsewhere on the sheet, press F9/Ctrl+Alt+F9, and edit `nonce` itself — note which of
 * these actually change the returned timestamp.
 * @customfunction
 * @param {number} nonce Any value — changing it is the "edited input" trigger specifically.
 * @returns {Promise<string>} The ISO timestamp of this invocation.
 */
async function recalcProbe(nonce) {
    return new Date().toISOString();
}

// Same bearer-token convention as gapiHeaders() in taskpane.js (taskpane.js:937-941) — a static,
// build-time-baked value via the SAME webpack.DefinePlugin injection, identical regardless of
// which bundle/runtime reads it.
const GOINDIASTOCKS_API_TOKEN = process.env.GOINDIASTOCKS_API_TOKEN || "";

const BASIC_INFO_TABLE = "GIAStock_Data.dbo.basic_info";
const SCHEMA_BASE_URL = "https://ir-reports.s3.ap-south-1.amazonaws.com/dbcatalog/";
const PERIOD_FORMAT_PLAIN_YEAR = "PLAIN_YEAR";
const PERIOD_FORMAT_MONTH_YEAR_ANNUAL = "MONTH_YEAR_ANNUAL";
const PERIOD_FORMAT_MONTH_YEAR_QUARTERLY = "MONTH_YEAR_QUARTERLY";
// The forward fiscal-quarter mapping sc_quater_data's own doc documents: Jun/Sep/Dec belong to
// the NEXT fiscal year, Mar to the SAME one — same convention taskpane.js's own Table Builder
// uses, duplicated here since this file is deliberately self-contained (see the header comment).
const QUARTER_MONTH_BY_NO = { 1: "Jun", 2: "Sep", 3: "Dec", 4: "Mar" };

// Identifier-only charset — for table/column/periodColumn/modeColumn names when used as raw SQL
// identifiers.
const safeIdentifier = (s) => String(s || "").replace(/[^A-Za-z0-9_.]/g, "");
// Real column names carry spaces/parens/%/&/-/. (e.g. "EBITDA Margin (%)", "Net Debt to Equity").
const safeColumnName = (s) => String(s || "").replace(/[^a-zA-Z0-9 %()&/._-]/g, "");
const shortTableName = (t) => String(t || "").split(".").pop();

// Shared by GET()'s own value fetch and the templated-table column resolver below — same
// auth/body shape taskpane.js's own runGeneratedSql uses.
async function runWalletQueryRaw(sql, userQuery) {
    const res = await fetch("https://api.goindiastocks.com/wallet/excel_run_query", {
        method: "POST",
        headers: { "Authorization": `Bearer ${GOINDIASTOCKS_API_TOKEN}`, "Content-Type": "application/json" },
        body: JSON.stringify({
            user_id: null, enterprise_id: null,
            sql, user_query: userQuery, display_name: "GIA.GET",
            tables_used: [], title: null, description: null,
        }),
    });
    if (!res.ok) {
        let detail = "";
        try { detail = (await res.json()).detail || ""; } catch (e) { /* body wasn't JSON */ }
        throw new Error(detail || String(res.status));
    }
    return res.json();
}

// EVERY /wallet/excel_run_query call from this runtime goes through one chain, one at a time.
// This is not a politeness throttle — it's a correctness requirement. taskpane.js hit this twice
// and documented it both times (its loadCatalogEntries templated-column loop and its
// gatherAllMetricNames call, both deliberately sequential rather than Promise.all): firing several
// of these simultaneously makes SOME come back truncated/incomplete, silently losing rows. Excel
// invokes custom functions in PARALLEL, so a sheet full of GIA.GET was doing exactly that — which
// is why batching below matters twice over: serializing thousands of calls would be unusable,
// serializing the handful that batching leaves is not.
let walletQueryChain = Promise.resolve();
function runWalletQuery(sql, userQuery) {
    const result = walletQueryChain.then(() => runWalletQueryRaw(sql, userQuery));
    // The chain itself must never stay rejected, or one failure would poison every later call.
    walletQueryChain = result.then(() => {}, () => {});
    return result;
}

// A column name containing "<...>" is a TEMPLATE placeholder in the schema doc (e.g.
// "<filter_name where Type='Y'> (e.g. Revenue, ROAE)" for sc_year_data) rather than a real
// column — these tables register their real per-metric columns dynamically, so the doc lists
// one placeholder row instead of all ~168 of them. Confirmed live: sc_year_data's own doc text
// says exactly this ("Column names in this table correspond exactly to 'filter_name' values in
// 'GIAStock_Data.dbo.global_filter'...").
const isTemplatedColumnName = (name) => /<[^>]+>/.test(name);
// Resolves a templated table's REAL column names straight from INFORMATION_SCHEMA — same
// mechanism taskpane.js's own resolveTemplatedTableColumns uses for the regular Table Builder,
// since different templated tables relate to global_filter differently (1:1 by filter_name for
// some, 1:many for others per taskpane.js's own comment) and INFORMATION_SCHEMA is the one
// source that's correct regardless of which convention a given table follows. Paginated on
// ORDINAL_POSITION since /wallet/excel_run_query row-caps every call (sc_year_data alone has
// ~168 real columns). Cached for this runtime's lifetime — every GIA.GET formula referencing the
// same templated table reuses one resolution.
// Caches the in-flight PROMISE, not the resolved value: the value was only ever stored after its
// own await, so a burst of cells (Excel invokes custom functions in parallel) could all miss the
// cache and fire the same lookup several times over before the first one came back. Storing the
// promise means every later caller joins the one already running.
const realColumnsCache = new Map();
function fetchRealTableColumns(qualifiedTable) {
    if (!realColumnsCache.has(qualifiedTable)) realColumnsCache.set(qualifiedTable, fetchRealTableColumnsUncached(qualifiedTable));
    return realColumnsCache.get(qualifiedTable);
}
async function fetchRealTableColumnsUncached(qualifiedTable) {
    const parts = qualifiedTable.split(".");
    let names = [];
    if (parts.length === 3 && parts.every((p) => /^[A-Za-z0-9_]+$/.test(p))) {
        const [dbName, schemaName, tblName] = parts;
        let cursor = 0;
        try {
            for (let guard = 0; guard < 50; guard++) {
                const sql = `SELECT COLUMN_NAME, ORDINAL_POSITION FROM ${dbName}.INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = '${schemaName}' AND TABLE_NAME = '${tblName}' AND ORDINAL_POSITION > ${cursor} ORDER BY ORDINAL_POSITION`;
                const data = await runWalletQuery(sql, "[GIA.GET] column lookup");
                const batch = data.rows || [];
                if (!batch.length) break;
                names.push(...batch.map((r) => r[0]));
                cursor = batch[batch.length - 1][1];
                if (!data.truncated) break;
            }
        } catch (e) { /* leave names as whatever was collected before the failure */ }
    }
    return names;
}
const qualifyTable = (short) => (short === "basic_info" ? BASIC_INFO_TABLE : `GIAStock_Data.dbo.${short}`);

// Converts a 2-digit fiscal-year suffix (as buildFormulaGrid's own "FY26"/"Q3 FY26" labels in
// taskpane.js carry — see parseFiscalLabel below) to a full 4-digit year, by picking whichever
// century keeps it within 10 years of the current year. This warehouse's annual data runs from
// 1991 to a handful of forward estimate years, so a flat "assume 2000s" guess would silently
// misread a genuine "FY91" (1991) as 2091 instead.
function fiscalTwoDigitToFullYear(twoDigit) {
    const now = new Date().getFullYear();
    const candidate2000s = 2000 + twoDigit;
    return candidate2000s > now + 10 ? 1900 + twoDigit : candidate2000s;
}
// Parses the EXACT label text buildFormulaGrid writes into a table's own Period column —
// "FY26"/"FY2026" (annual) or "Q3 FY26"/"Q3 FY2026" (quarterly) — so a formula can reference that
// CELL directly instead of baking its year/quarter in as separate literal arguments; editing the
// cell then re-targets the period on the next recalculation, the same live-reference property the
// company argument already has when it's a cell reference too (see buildFormulaGrid's own comment
// in taskpane.js). Returns null for anything else, INCLUDING a plain number — a hand-typed
// `=GIA.GET(company,metric,2026)` never reaches this at all (Excel passes that as a JS number, not
// a string), so the existing numeric path in get() below is completely unaffected either way.
function parseFiscalLabel(period) {
    if (typeof period !== "string") return null;
    const m = period.trim().match(/^(?:Q([1-4])\s+)?FY(\d{2}|\d{4})$/i);
    if (!m) return null;
    const quarterNo = m[1] ? parseInt(m[1], 10) : null;
    const yearDigits = m[2];
    const fiscalYear = yearDigits.length === 2 ? fiscalTwoDigitToFullYear(parseInt(yearDigits, 10)) : parseInt(yearDigits, 10);
    return { fiscalYear, quarterNo };
}
const classifyPeriodFormat = (example, dataFrequency) => {
    const ex = String(example || "").trim();
    if (/^\d{4}$/.test(ex)) return { format: PERIOD_FORMAT_PLAIN_YEAR };
    const monthYear = ex.match(/^([A-Za-z]{3})(\d{4})$/);
    if (monthYear) {
        const freq = String(dataFrequency || "").toLowerCase();
        if (freq.startsWith("year")) return { format: PERIOD_FORMAT_MONTH_YEAR_ANNUAL, monthPrefix: monthYear[1] };
        if (freq.startsWith("quarter")) return { format: PERIOD_FORMAT_MONTH_YEAR_QUARTERLY };
    }
    return null;
};
const literalPeriodValue = (periodInfo, requested) => {
    if (periodInfo.format === PERIOD_FORMAT_PLAIN_YEAR) return requested.quarterNo ? null : String(requested.fiscalYear);
    if (periodInfo.format === PERIOD_FORMAT_MONTH_YEAR_ANNUAL) {
        return requested.quarterNo ? null : `${periodInfo.monthPrefix}${requested.fiscalYear}`;
    }
    if (!requested.quarterNo) return null; // MONTH_YEAR_QUARTERLY needs an actual quarter
    const month = QUARTER_MONTH_BY_NO[requested.quarterNo];
    const calendarYear = requested.quarterNo === 4 ? requested.fiscalYear : requested.fiscalYear - 1;
    return `${month}${calendarYear}`;
};
// Parses a dbcatalog schema doc's Identifier+Data Columns tables into name -> { example } —
// covers BOTH sections since a metric's own column usually lives under Data Columns, while
// fincode/mode/year live under Identifier Columns.
const extractColumnDetails = (markdown) => {
    const details = new Map();
    let inSection = false, pastSeparator = false;
    for (const rawLine of markdown.split("\n")) {
        const line = rawLine.trim();
        if (/^###\s+(Identifier|Data)\s+Columns/i.test(line)) { inSection = true; pastSeparator = false; continue; }
        if (!inSection) continue;
        if (line.startsWith("#") || line === "---") { inSection = false; continue; }
        if (line === "") continue;
        if (!line.startsWith("|")) continue;
        if (/^\|[-\s|]+\|$/.test(line)) { pastSeparator = true; continue; }
        if (!pastSeparator) continue;
        const cells = line.split("|");
        const name = cells[1] ? cells[1].trim() : "";
        if (name) details.set(name, { example: cells[2] ? cells[2].trim() : "" });
    }
    return details;
};
const extractDataFrequency = (markdown) => {
    for (const rawLine of markdown.split("\n")) {
        const m = rawLine.match(/\|\s*\*\*Data Frequency\*\*\s*\|\s*([^|]+)\|/i);
        if (m) return m[1].trim();
    }
    return null;
};
// Scoped to this runtime's lifetime (the shared-runtime page, so it persists across every
// GIA.GET cell for the whole Excel session, not just one call) — every formula referencing the
// same table reuses one fetch instead of paying for it again.
// Caches the in-flight PROMISE rather than the resolved value, for the same reason
// realColumnsCache above does — see its own comment.
const schemaCache = new Map();
const loadTableSchema = (short) => {
    if (!schemaCache.has(short)) {
        const pending = loadTableSchemaUncached(short);
        // A REJECTED promise must not stay cached the way a resolved one does — otherwise a single
        // transient S3 blip would permanently break every formula on that table for the rest of
        // the session. (Caching the value only after the await, as this used to, got that right by
        // accident; caching the promise has to do it on purpose.)
        pending.catch(() => schemaCache.delete(short));
        schemaCache.set(short, pending);
    }
    return schemaCache.get(short);
};
const loadTableSchemaUncached = async (short) => {
    // dbcatalog's .md filenames are keyed by the FULLY-QUALIFIED table name (confirmed against
    // index.json, e.g. "GIAStock_Data.dbo.sc_year_data.md"), not the short name METRIC_REGISTRY's
    // "prefer"/table overrides use — fetching the short name 403s (S3 hides "no such key" behind
    // a 403 for anonymous requests, rather than a 404).
    const res = await fetch(`${SCHEMA_BASE_URL}${encodeURIComponent(qualifyTable(short))}.md`);
    if (!res.ok) throw new Error(`schema fetch for "${short}" returned ${res.status}`);
    const markdown = await res.text();
    return { details: extractColumnDetails(markdown), frequency: extractDataFrequency(markdown) };
};

/**
 * The single resolution point GET() calls on every invocation: given whatever the caller
 * supplied (a metric name and/or explicit table/column/periodLiteral overrides), returns the
 * fully-resolved `{ table, column, periodColumn, periodLiteral, modeColumn }` GET() needs to
 * build its query, or `{ error }`. Three override params let a caller skip part or all of this:
 * - `table` + `column` + `periodLiteral` ALL given: skips every lookup entirely (no network
 *   call) — the caller is fully responsible for correctness in this case, including mode and
 *   quarter (if any) are baked into `periodLiteral` directly, so `quarterNo` is ignored here.
 * - `table` + `column` given, `periodLiteral` missing: skips the metric->table lookup, but
 *   still fetches that table's schema to find its period column/format + mode column.
 * - Neither given: looks the metric up in METRIC_REGISTRY — the quarterly slot when `quarterNo`
 *   is given (falling back to annual if the metric has no quarterly form), otherwise the annual
 *   slot (falling back to quarterly) — then fetches that table's schema the same way.
 */
async function resolveGetArgs(metric, period, quarterNo, table, column, periodLiteral) {
    if (table && column && periodLiteral) {
        return { table: qualifyTable(shortTableName(table)), column, periodColumn: "", periodLiteral, modeColumn: "", estimateColumn: "" };
    }

    let short = table ? shortTableName(table) : null;
    let candidateColumns = column ? String(column).split("|") : null;

    if (!short || !candidateColumns) {
        if (!metric) return { error: "metric is required when table/column aren't both given" };
        const entry = (METRIC_REGISTRY.metrics || []).find((m) => m.name === metric);
        if (!entry) return { error: `"${metric}" isn't a known metric` };
        const slot = quarterNo ? (entry.quarterly || entry.annual) : (entry.annual || entry.quarterly);
        if (!slot || !slot.prefer) return { error: `"${metric}" has no resolvable table` };
        short = short || slot.prefer;
        candidateColumns = candidateColumns || slot.columns || [];
    }

    let schema;
    try { schema = await loadTableSchema(short); } catch (e) { return { error: `couldn't load "${short}"'s schema: ${e.message}` }; }

    if (short === "basic_info") {
        const col = candidateColumns.find((c) => schema.details.has(c));
        if (!col) return { error: `couldn't match a real column on basic_info` };
        return { table: BASIC_INFO_TABLE, column: col, periodColumn: "", periodLiteral: "", modeColumn: "", estimateColumn: "" };
    }

    let cols = candidateColumns.filter((c) => schema.details.has(c));
    if (!cols.length && Array.from(schema.details.keys()).some(isTemplatedColumnName)) {
        // Templated table (e.g. sc_year_data, sc_quater_data) — its real per-metric columns
        // never appear literally in the doc, so fall back to INFORMATION_SCHEMA.
        const realColumns = await fetchRealTableColumns(qualifyTable(short));
        cols = candidateColumns.filter((c) => realColumns.includes(c));
    }
    if (!cols.length) return { error: `couldn't match a real column on ${short}` };
    const names = Array.from(schema.details.keys());
    if (!names.some((n) => n.toLowerCase() === "fincode")) return { error: `${short} has no fincode column to join on` };
    const modeColumn = names.find((n) => n.toLowerCase() === "mode") || "";
    // Some tables (e.g. sc_cagr_data) store both actual historical values AND future consensus
    // estimates under the same period key, flagged by this column — mirrors taskpane.js's own
    // loadPeriodJoinMetadata/buildDeterministicQuery estimateFilter, so GET() can't silently
    // return a forecast instead of the real actual for a period both exist for.
    const estimateColumn = names.find((n) => n.toLowerCase() === "estimate") || "";
    const yearName = names.find((n) => n.toLowerCase() === "year");

    let resolvedPeriodLiteral = periodLiteral || "";
    let periodColumn = "";
    if (yearName) {
        periodColumn = yearName;
        if (!resolvedPeriodLiteral) {
            const periodInfo = classifyPeriodFormat(schema.details.get(yearName).example, schema.frequency);
            if (!periodInfo) return { error: `${short}'s period format wasn't recognized` };
            const fiscalYear = parseInt(period, 10);
            if (!Number.isFinite(fiscalYear)) return { error: "a numeric fiscal year is required" };
            const literal = literalPeriodValue(periodInfo, { fiscalYear, quarterNo: quarterNo || null });
            if (literal === null) {
                return { error: quarterNo ? `${short} doesn't have quarterly data for this metric` : "this metric only reports quarterly — a quarter (1-4) is required" };
            }
            resolvedPeriodLiteral = literal;
        }
    } else if (quarterNo) {
        return { error: `${short} has no period column to filter a quarter on` };
    }
    return { table: qualifyTable(short), column: cols.join("|"), periodColumn, periodLiteral: resolvedPeriodLiteral, modeColumn, estimateColumn };
}

// ── Batching ──────────────────────────────────────────────────────────────────────────────────
// GET() used to build a one-row "SELECT TOP 1 ..." and fetch it per cell, which made the number of
// network calls a function of the number of CELLS. A 50-company x 10-metric x 5-year table is 2,500
// cells, i.e. 2,500 round trips; the whole cost is network, not Excel's own calculation engine.
// Everything below turns that into a function of the number of distinct TABLES instead: requests
// arriving in the same recalculation burst are collected, companies are resolved once, and each
// table is read with a single multi-company/multi-period/multi-column query.
//
// Every tunable lives here because the backend's own metering behavior isn't known — these are the
// knobs that trade call count against latency.
const BATCH_WINDOW_MS = 20;        // Excel fires a recalculation's cells in one burst; this collects it
const MAX_BATCH_REQUESTS = 2000;   // flush early rather than let one flush grow unboundedly
const MAX_NAMES_PER_RESOLVE = 40;  // UNION ALL branches per company-resolution query
const MAX_MATCHES_PER_NAME = 10;   // candidate basic_info rows kept per name (see resolveCompanies)
const VALUE_CACHE_TTL_MS = 6 * 60 * 60 * 1000;
// Company financials are effectively immutable once reported, so a value is reused for the session;
// the TTL above only self-heals a session left open for days across a results announcement.
// /wallet/excel_run_query wraps every query in a row cap and reports it back as `limit` (observed:
// 5,000). This starts under that and re-tunes itself from the first real response, so the figure
// here never has to be exactly right.
let rowBudget = 4000;

const escapeSqlLiteral = (s) => String(s).replace(/'/g, "''");
const valueCache = new Map();    // "table|column|company|periodLiteral" -> { value, found, fetchedAt }
const companyCache = new Map();  // raw company argument -> [{ fincode, compName }, ...] (empty = no match)
let pendingRequests = [];
let flushTimer = null;

// Keyed on the company STRING the formula actually passed, not on a resolved fincode: which
// fincode wins isn't known until the rows come back (see the candidate walk in flushBatch).
const valueCacheKey = (r) => `${r.table}|${r.column}|${r.company}|${r.periodLiteral}`;

// Resolves each DISTINCT company string to its matching basic_info rows, once per session.
//
// Deliberately a UNION ALL of per-name LIKE branches rather than one OR-ed LIKE, so each name is
// matched exactly as the old per-cell query matched it and batching can't change what a given
// string resolves to. (Verified against the live backend: this shape survives the server's own
// "SELECT TOP N * FROM (<sql>) AS t" wrapper, which is also why it can't be written as a CTE.)
//
// Each branch keeps up to MAX_MATCHES_PER_NAME rows IN NATURAL ORDER rather than just TOP 1,
// because the old query applied its TOP 1 AFTER joining to the data table — so a name matching
// both an entity with no rows and one with rows (e.g. a "- DVR Ordinary" listing alongside the
// main one) resolved to the one that actually had data. Row 0 is still what a bare TOP 1 would
// have returned, so the common unambiguous case is unchanged; the rest are fallbacks that only
// come into play when an earlier candidate has no row for the period being asked for.
async function resolveCompanies(names) {
    const todo = names.filter((n) => !companyCache.has(n));
    for (let i = 0; i < todo.length; i += MAX_NAMES_PER_RESOLVE) {
        const chunk = todo.slice(i, i + MAX_NAMES_PER_RESOLVE);
        const sql = chunk.map((name) => {
            const safe = escapeSqlLiteral(name);
            return `SELECT TOP ${MAX_MATCHES_PER_NAME} '${safe}' AS q, b.fincode, b.CompName FROM ${BASIC_INFO_TABLE} b `
                + `WHERE b.CompName LIKE '%${safe}%' AND b.isdefault = 'yes'`;
        }).join("\nUNION ALL\n");
        const data = await runWalletQuery(sql, `[GIA.GET] resolve ${chunk.length} compan${chunk.length === 1 ? "y" : "ies"}`);
        const hits = new Map();
        for (const row of (data.rows || [])) {
            const q = String(row[0]);
            if (!hits.has(q)) hits.set(q, []);
            hits.get(q).push({ fincode: String(row[1]), compName: row[2] });
        }
        // A name with no rows genuinely matched nothing — cached as an empty list so it isn't
        // re-asked on every later recalculation.
        for (const name of chunk) companyCache.set(name, hits.get(name) || []);
    }
}

// One query per table: every fincode x every period x every column that table's pending requests
// asked for. Mirrors the SQL shape buildDeterministicQuery already runs in production
// (taskpane.js), minus the ranking/multi-table-join machinery GET() has no use for.
function buildBatchSql(group, fincodes) {
    const periodColumn = safeIdentifier(group.periodColumn);
    const modeColumn = safeIdentifier(group.modeColumn);
    const estimateColumn = safeIdentifier(group.estimateColumn);
    const inList = fincodes.map((f) => `'${escapeSqlLiteral(f)}'`).join(", ");
    const isBasicInfo = group.table === BASIC_INFO_TABLE;
    const alias = isBasicInfo ? "b" : "t1";

    const select = ["b.fincode"];
    if (!isBasicInfo && periodColumn) select.push(`${alias}.[${periodColumn}] AS [Period]`);
    // One output column per DISTINCT real column the group needs — a request whose metric resolved
    // to an "A|B" candidate list picks the first non-null of its own candidates from the row, which
    // is what the old per-cell COALESCE(...) did, just done once for every request that shares them.
    for (const c of group.columns) select.push(`${alias}.[${c}] AS [${c}]`);

    let sql = `SELECT ${select.join(", ")} FROM ${BASIC_INFO_TABLE} b`;
    if (!isBasicInfo) {
        sql += ` JOIN ${group.table} t1 ON t1.fincode = b.fincode`;
        // mode (Consolidated vs Standalone) is never a literal — it comes from basic_info's own
        // column, and only tables that actually have one get the join.
        if (modeColumn) sql += ` AND t1.[${modeColumn}] = b.mode`;
    }
    sql += ` WHERE b.fincode IN (${inList}) AND b.isdefault = 'yes'`;
    if (!isBasicInfo && periodColumn && group.periods.length) {
        sql += ` AND ${alias}.[${periodColumn}] IN (${group.periods.map((p) => `'${escapeSqlLiteral(p)}'`).join(", ")})`;
    }
    // Never let a future consensus-estimate row win over the real actual for the same period —
    // see the estimateColumn comment in resolveGetArgs.
    if (!isBasicInfo && estimateColumn) sql += ` AND ${alias}.[${estimateColumn}] = 'no'`;
    return sql;
}

// Runs one table-group over one set of fincodes and hands each row back keyed by (fincode, period).
// A response that still comes back `truncated` is split and re-run rather than trusted: the repo's
// standing rule for this endpoint (three separate implementations of it) is that a short read must
// never be mistaken for a complete one.
async function fetchGroupChunk(group, fincodes, depth = 0) {
    const sql = buildBatchSql(group, fincodes);
    const data = await runWalletQuery(sql, `[GIA.GET] ${shortTableName(group.table)} x${fincodes.length}`);
    if (Number.isFinite(data.limit) && data.limit > 0) rowBudget = Math.min(rowBudget, data.limit - 1);
    if (data.truncated && fincodes.length > 1 && depth < 6) {
        const mid = Math.ceil(fincodes.length / 2);
        const [a, b] = await Promise.all([
            fetchGroupChunk(group, fincodes.slice(0, mid), depth + 1),
            fetchGroupChunk(group, fincodes.slice(mid), depth + 1),
        ]);
        for (const [k, v] of b) a.set(k, v);
        return a;
    }
    const columns = (data.columns || []).map((c) => c.name);
    const fincodeIdx = 0; // selected first, by construction
    const periodIdx = columns.indexOf("Period");
    const byKey = new Map();
    for (const row of (data.rows || [])) {
        const period = periodIdx >= 0 ? String(row[periodIdx]) : "";
        byKey.set(`${String(row[fincodeIdx])}|${period}`, { row, columns });
    }
    return byKey;
}

async function flushBatch() {
    flushTimer = null;
    const batch = pendingRequests;
    pendingRequests = [];
    if (!batch.length) return;

    try {
        await resolveCompanies(Array.from(new Set(batch.map((r) => r.company))));
    } catch (e) {
        for (const r of batch) r.reject(e);
        return;
    }

    // Anything whose company matched nothing is done here — same answer the per-cell query gave.
    const live = [];
    for (const r of batch) {
        const matches = companyCache.get(r.company);
        if (!matches || !matches.length) { r.resolve({ found: false }); continue; }
        r.fincodes = matches.map((m) => m.fincode);
        const cached = valueCache.get(valueCacheKey(r));
        if (cached && Date.now() - cached.fetchedAt < VALUE_CACHE_TTL_MS) { r.resolve(cached); continue; }
        live.push(r);
    }
    if (!live.length) return;

    // Group by everything that has to be identical for one query to serve them all.
    const groups = new Map();
    for (const r of live) {
        const key = `${r.table}|${r.periodColumn}|${r.modeColumn}|${r.estimateColumn}`;
        let g = groups.get(key);
        if (!g) {
            g = { table: r.table, periodColumn: r.periodColumn, modeColumn: r.modeColumn,
                estimateColumn: r.estimateColumn, columns: new Set(), periods: new Set(),
                fincodes: new Set(), requests: [] };
            groups.set(key, g);
        }
        for (const c of r.candidates) g.columns.add(c);
        if (r.periodLiteral) g.periods.add(r.periodLiteral);
        for (const f of r.fincodes) g.fincodes.add(f);
        g.requests.push(r);
    }

    for (const g of groups.values()) {
        g.columns = Array.from(g.columns);
        g.periods = Array.from(g.periods);
        const fincodes = Array.from(g.fincodes);
        // Rows come back as fincodes x periods, so chunk the fincodes to stay under the row cap.
        const perCompany = Math.max(1, g.periods.length);
        const chunkSize = Math.max(1, Math.floor(rowBudget / perCompany));
        try {
            const found = new Map();
            for (let i = 0; i < fincodes.length; i += chunkSize) {
                const part = await fetchGroupChunk(g, fincodes.slice(i, i + chunkSize));
                for (const [k, v] of part) found.set(k, v);
            }
            for (const r of g.requests) {
                // Walk this request's own candidate companies in order and take the first that
                // actually has a row for the period asked for — what the old query's "TOP 1 after
                // the join" did (see resolveCompanies). Other requests' fincodes are in the same
                // response but are never considered here.
                let match = null;
                for (const f of r.fincodes) {
                    match = found.get(`${f}|${r.periodLiteral || ""}`);
                    if (match) break;
                }
                if (!match) { r.resolve({ found: false }); continue; }
                // First non-null across the request's own candidate columns — the COALESCE the
                // old per-cell SQL did, applied per request rather than per query.
                let value = null;
                for (const c of r.candidates) {
                    const idx = match.columns.indexOf(c);
                    if (idx >= 0 && match.row[idx] !== null && match.row[idx] !== undefined) { value = match.row[idx]; break; }
                }
                const result = { found: true, value, fetchedAt: Date.now() };
                valueCache.set(valueCacheKey(r), result);
                r.resolve(result);
            }
        } catch (e) {
            for (const r of g.requests) r.reject(e);
        }
    }
}

// Called by GET() in place of its old per-cell fetch. Identical pending keys share one promise, so
// duplicate cells never cost two lookups even inside a single burst.
function enqueueValue(request) {
    return new Promise((resolve, reject) => {
        pendingRequests.push({ ...request, resolve, reject });
        if (pendingRequests.length >= MAX_BATCH_REQUESTS) {
            if (flushTimer) { clearTimeout(flushTimer); flushTimer = null; }
            flushBatch();
        } else if (!flushTimer) {
            flushTimer = setTimeout(flushBatch, BATCH_WINDOW_MS);
        }
    });
}

// Cross-bundle handle, same pattern as taskpane.js's own window.__goiaChatBridge — the add-in runs
// a SharedRuntime, so the task pane shares this page and can drop these caches (e.g. behind a
// future "Refresh values" control) and then force a recalculation.
if (typeof window !== "undefined") {
    window.__giaGetBridge = {
        clearCache: () => { valueCache.clear(); companyCache.clear(); },
        stats: () => ({ values: valueCache.size, companies: companyCache.size, rowBudget }),
    };
}

/**
 * =GIA.GET(company, metric, period) — hand-typeable with just those 3. `quarter` is an optional
 * 4th value (1-4) for a quarterly read instead of annual; `table`/`column`/`periodLiteral` are
 * optional power-user overrides (see resolveGetArgs); a typical formula never needs them.
 * @customfunction
 * @param {string} companyName Company name to look up (loose match).
 * @param {string} metric Canonical metric name, e.g. "Revenue". Only optional when both `table`
 * and `column` are given directly.
 * @param {number|string} period Fiscal year, e.g. 2026 — or a cell reference to a label this app
 * itself writes, "FY26"/"Q3 FY26" (see parseFiscalLabel), which also supplies the quarter.
 * @param {number} [quarter] Quarter number 1-4, for a quarterly read instead of annual. Omit for
 * the annual figure, or when `period` is already a "Q.. FY.." label.
 * @param {string} [table] Advanced: skip metric resolution and use this real table directly,
 * e.g. "GIAStock_Data.dbo.sc_year_data".
 * @param {string} [column] Advanced: skip metric resolution and use this real column directly.
 * @param {string} [periodLiteral] Advanced: skip period-format resolution and filter on this
 * literal (already in the target table's own stored format) directly.
 * @returns {Promise<number|string>} The bare numeric value, or a "GET_ERROR:" message.
 */
async function get(companyName, metric, period, quarter, table, column, periodLiteral) {
    try {
        // Passed to the batcher RAW, not pre-escaped: resolveCompanies escapes it once when it
        // builds the LIKE, and escaping here too would double up every apostrophe.
        const rawCompany = String(companyName || "");
        if (!rawCompany) return "GET_ERROR: company name required";

        // A `period` referencing a cell that holds one of this app's own "FY26"/"Q3 FY26" labels
        // (see buildFormulaGrid's Period column in taskpane.js) carries its own quarter, if any —
        // an explicit `quarter` argument only matters for the plain-number `period` form below, so
        // a recognized label decides BOTH outright rather than being merged with it. A hand-typed
        // numeric period (Excel passes that as a real JS number, never a string) never matches here
        // at all, so this never touches the existing plain-number path's own behavior.
        const fiscalLabel = parseFiscalLabel(period);
        const effectivePeriod = fiscalLabel ? fiscalLabel.fiscalYear : period;

        let quarterNo = null;
        if (fiscalLabel) {
            quarterNo = fiscalLabel.quarterNo;
        } else if (quarter !== undefined && quarter !== null && quarter !== "") {
            quarterNo = parseInt(quarter, 10);
            if (!Number.isFinite(quarterNo) || quarterNo < 1 || quarterNo > 4) return "GET_ERROR: quarter must be 1-4";
        }

        const resolved = await resolveGetArgs(metric, effectivePeriod, quarterNo, table, column, periodLiteral);
        if (resolved.error) return `GET_ERROR: ${resolved.error}`;
        const safeTable = safeIdentifier(resolved.table);
        if (!safeTable || !/^GIAStock_Data\.dbo\.[A-Za-z_][A-Za-z0-9_]*$/.test(safeTable)) {
            return "GET_ERROR: invalid or disallowed table";
        }
        const candidates = String(resolved.column).split("|").map(safeColumnName).filter(Boolean);
        if (!candidates.length) return "GET_ERROR: invalid column";

        // Handed to the batcher rather than fetched here, so a sheet full of these costs one query
        // per TABLE instead of one per cell — see enqueueValue/flushBatch above. identity
        // resolution is WHOAMI's job to prove first — user_id/enterprise_id left unset here on
        // purpose (see runWalletQuery).
        const hit = await enqueueValue({
            company: rawCompany,
            table: safeTable,
            column: resolved.column,
            candidates,
            periodColumn: resolved.periodColumn || "",
            periodLiteral: resolved.periodLiteral || "",
            modeColumn: resolved.modeColumn || "",
            estimateColumn: resolved.estimateColumn || "",
        });
        if (!hit.found) return "GET_ERROR: no matching data";
        // Coerced to a real Number rather than passed through as-is: many numeric-looking columns
        // in this warehouse are stored as TEXT, so the backend can legitimately hand back
        // "1057219" as a string — returning that AS a string would make Excel treat the cell as
        // text (left-aligned, no numeric formatting applied), not the number it is.
        const value = Number(hit.value);
        return Number.isFinite(value) ? value : "GET_ERROR: non-numeric value returned";
    } catch (e) {
        return `GET_ERROR: ${e.message}`;
    }
}

/**
 * SPIKE ONLY. No network call — a hardcoded 2D array to prove dynamic-array spilling works at
 * all before GET even needs to.
 * @customfunction
 * @returns {string[][]} A 3-row x 2-column table.
 */
function getTable() {
    return [
        ["Company", "Revenue"],
        ["Reliance Industries Ltd.", "1057219"],
        ["Tata Motors Ltd.", "83855"],
    ];
}
