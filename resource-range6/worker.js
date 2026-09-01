const ALLOWED_ORIGIN = "https://solar-designed-map.web.app";
const MAX_BODY_BYTES = 1_500_000;
const SESSION_LIFETIME_MS = 2 * 60 * 60 * 1000;

export default {
  async fetch(request, env) {
    try {
      return await route(request, env);
    } catch (error) {
      return json({ ok: false, error: "INTERNAL_ERROR", message: String(error?.message || error) }, 500, request);
    }
  }
};

async function route(request, env) {
  const url = new URL(request.url);
  const path = url.pathname.replace(/\/+$/, "") || "/";
  const method = request.method.toUpperCase();
  const origin = request.headers.get("Origin") || "";

  if (origin && origin !== ALLOWED_ORIGIN) {
    return json({ ok: false, error: "ORIGIN_NOT_ALLOWED" }, 403, request);
  }
  if (method === "OPTIONS") return empty(204, request);
  if (method === "GET" && (path === "/" || path === "/v1/health")) {
    return json({ ok: true, service: "SolarEPCCloudRelay", version: "1.8-resource-range6" }, 200, request);
  }

  if (method === "POST" && path === "/v1/session/open") {
    const workbookId = request.headers.get("X-Solar-EPC-Workbook") || "";
    const workbookKey = request.headers.get("X-Solar-EPC-Key") || "";
    if (!validWorkbook(env, workbookId, workbookKey)) return json({ ok: false, error: "INVALID_WORKBOOK_CREDENTIALS" }, 401, request);

    const now = Date.now();
    const token = crypto.randomUUID() + crypto.randomUUID().replaceAll("-", "");
    const expiresAt = now + SESSION_LIFETIME_MS;
    await env.DB.prepare("DELETE FROM sessions WHERE expires_at < ?").bind(now).run();
    await env.DB.prepare("INSERT INTO sessions(token, workbook_id, expires_at, created_at) VALUES(?,?,?,?)")
      .bind(token, workbookId, expiresAt, now + 20000).run();
    return json({ ok: true, sessionToken: token, expiresAt }, 200, request);
  }

  if (method === "POST" && path === "/v1/save") {
    const session = await validSession(env, bearer(request));
    if (!session) return json({ ok: false, error: "INVALID_OR_EXPIRED_SESSION" }, 401, request);
    const declared = Number(request.headers.get("Content-Length") || 0);
    if (declared > MAX_BODY_BYTES) return json({ ok: false, error: "PAYLOAD_TOO_LARGE" }, 413, request);
    const text = await request.text();
    if (new TextEncoder().encode(text).length > MAX_BODY_BYTES) return json({ ok: false, error: "PAYLOAD_TOO_LARGE" }, 413, request);

    let message;
    try { message = JSON.parse(text); } catch { return json({ ok: false, error: "INVALID_JSON" }, 400, request); }
    const messageId = clean(message?.messageId, 200);
    const type = String(message?.type || "");
    const excelText = String(message?.excelText || "");
    if (type !== "SOLAR_EPC_SITE_BOUNDARY_SAVED") return json({ ok: false, error: "UNSUPPORTED_MESSAGE_TYPE" }, 400, request);
    if (!messageId) return json({ ok: false, error: "MISSING_MESSAGE_ID" }, 400, request);
    if (!excelText.startsWith("SOLAR_EPC_MAP_SAVE_V1")) return json({ ok: false, error: "INVALID_EXCEL_PAYLOAD" }, 400, request);

    const existing = await env.DB.prepare("SELECT workbook_id FROM saves WHERE message_id=?").bind(messageId).first();
    if (existing) {
      if (existing.workbook_id !== session.workbook_id) return json({ ok: false, error: "MESSAGE_ID_CONFLICT" }, 409, request);
      return json({ ok: true, received: true, duplicate: true, messageId }, 200, request);
    }

    await env.DB.prepare("INSERT INTO saves(message_id,workbook_id,excel_text,created_at) VALUES(?,?,?,?)")
      .bind(messageId, session.workbook_id, excelText, Date.now()).run();
    return json({ ok: true, received: true, messageId }, 200, request);
  }

  if (method === "GET" && path.startsWith("/v1/status/")) {
    const session = await validSession(env, bearer(request));
    if (!session) return json({ ok: false, error: "INVALID_OR_EXPIRED_SESSION" }, 401, request);
    const messageId = decodeURIComponent(path.slice("/v1/status/".length));
    const row = await env.DB.prepare("SELECT 1 AS pending FROM saves WHERE message_id=? AND workbook_id=?")
      .bind(messageId, session.workbook_id).first();
    return json({ ok: true, pending: !!row, acknowledged: !row }, 200, request);
  }

  if (method === "POST" && path === "/v1/session/heartbeat") {
    const token = bearer(request);
    const session = await validSession(env, token);
    if (!session) return json({ ok: false, error: "INVALID_OR_EXPIRED_SESSION" }, 401, request);

    let active = false;
    try { active = (await request.json())?.active === true; } catch {}

    const now = Date.now();
    const expiresAt = now + 7000;
    const requestedFastUntil = active ? now + 20000 : 0;
    await env.DB.prepare(
      "UPDATE sessions SET expires_at=?, created_at=CASE WHEN created_at>? THEN created_at ELSE ? END WHERE token=?"
    ).bind(expiresAt, requestedFastUntil, requestedFastUntil, token).run();

    return json({ ok: true, alive: true, active, expiresAt }, 200, request);
  }

  if (method === "POST" && path === "/v1/session/idle") {
    const token = bearer(request);
    const session = await validSession(env, token);
    if (!session) return json({ ok: false, error: "INVALID_OR_EXPIRED_SESSION" }, 401, request);
    await env.DB.prepare("UPDATE sessions SET created_at=0 WHERE token=?")
      .bind(token).run();
    return json({ ok: true, idle: true }, 200, request);
  }

  if (method === "POST" && path === "/v1/session/cancel") {
    const token = bearer(request);
    const session = await validSession(env, token);
    if (!session) return json({ ok: true, cancelled: true, alreadyClosed: true }, 200, request);
    await env.DB.prepare("DELETE FROM sessions WHERE token=?").bind(token).run();
    return json({ ok: true, cancelled: true }, 200, request);
  }

  const workbook = request.headers.get("X-Solar-EPC-Workbook") || "";
  const key = request.headers.get("X-Solar-EPC-Key") || "";
  if (!validWorkbook(env, workbook, key)) return json({ ok: false, error: "INVALID_WORKBOOK_CREDENTIALS" }, 401, request);

  // NASA POWER resource routes are workbook-authenticated and isolated from
  // the existing temporary SAVE/ACK queue.
  if (method === "POST" && path === "/v1/resource/start") {
    return resourceStart(request, env, workbook);
  }
  if (method === "POST" && path === "/v1/resource/process") {
    return resourceProcess(request, env, workbook);
  }
  if (method === "POST" && path === "/v1/resource/process-month") {
    return resourceProcessMonth(request, env, workbook);
  }
  if (method === "POST" && path === "/v1/resource/process-range") {
    return resourceProcessRange(request, env, workbook);
  }
  if (method === "POST" && path === "/v1/resource/finalize") {
    return resourceFinalize(request, env, workbook);
  }
  if (method === "GET" && path.startsWith("/v1/resource/plan/")) {
    const requestId = decodeURIComponent(path.slice("/v1/resource/plan/".length));
    return resourcePlan(request, env, workbook, requestId);
  }
  if (method === "GET" && path === "/v1/resource/pending") {
    return resourcePending(request, env, workbook);
  }
  if (method === "GET" && path.startsWith("/v1/resource/status/")) {
    const requestId = decodeURIComponent(path.slice("/v1/resource/status/".length));
    return resourceStatus(request, env, workbook, requestId);
  }
  if (method === "GET" && path.startsWith("/v1/resource/summary/")) {
    const requestId = decodeURIComponent(path.slice("/v1/resource/summary/".length));
    return resourceSummary(request, env, workbook, requestId);
  }

  if (method === "GET" && path === "/v1/pending") {
    const row = await env.DB.prepare("SELECT message_id,excel_text FROM saves WHERE workbook_id=? ORDER BY created_at ASC LIMIT 1")
      .bind(workbook).first();

    if (!row) {
      const watchToken = request.headers.get("X-Solar-EPC-Session") || "";
      if (watchToken) {
        const watched = await env.DB.prepare(
          "SELECT workbook_id,expires_at,created_at FROM sessions WHERE token=?"
        ).bind(watchToken).first();
        if (!watched || watched.workbook_id !== workbook || Number(watched.expires_at) < Date.now()) {
          return json({ ok: false, error: "SESSION_CLOSED" }, 410, request);
        }
        const headers = corsHeaders(request);
        headers.set("X-Solar-EPC-Watch-Mode", Number(watched.created_at) > Date.now() ? "fast" : "idle");
        headers.set("Access-Control-Expose-Headers", "X-Solar-EPC-Watch-Mode");
        return new Response(null, { status: 204, headers });
      }
      return empty(204, request);
    }

    const headers = corsHeaders(request);
    headers.set("Content-Type", "text/plain; charset=utf-8");
    headers.set("Cache-Control", "no-store");
    headers.set("X-Solar-EPC-Message-Id", row.message_id);
    headers.set("Access-Control-Expose-Headers", "X-Solar-EPC-Message-Id");
    return new Response(row.excel_text, { status: 200, headers });
  }

  if (method === "POST" && path === "/v1/ack") {
    let body;
    try { body = await request.json(); } catch { return json({ ok: false, error: "INVALID_JSON" }, 400, request); }
    const messageId = clean(body?.messageId, 200);
    if (!messageId) return json({ ok: false, error: "MISSING_MESSAGE_ID" }, 400, request);

    // An ACK is successful only when this exact workbook/message row is gone.
    // If a prior retry already deleted it, ACK remains safely idempotent.
    const before = await env.DB.prepare(
      "SELECT message_id FROM saves WHERE message_id=? AND workbook_id=?"
    ).bind(messageId, workbook).first();

    if (!before) {
      return json({ ok: true, acknowledged: true, deleted: false, alreadyAbsent: true, messageId }, 200, request);
    }

    await env.DB.prepare(
      "DELETE FROM saves WHERE message_id=? AND workbook_id=?"
    ).bind(messageId, workbook).run();

    const remaining = await env.DB.prepare(
      "SELECT message_id FROM saves WHERE message_id=? AND workbook_id=?"
    ).bind(messageId, workbook).first();

    if (remaining) {
      return json({ ok: false, acknowledged: false, deleted: false, error: "DELETE_NOT_CONFIRMED", messageId }, 500, request);
    }

    return json({ ok: true, acknowledged: true, deleted: true, messageId }, 200, request);
  }

  return json({ ok: false, error: "NOT_FOUND" }, 404, request);
}

function validWorkbook(env, id, key) {
  return id.length >= 8 && key.length >= 24 && timingSafeEqual(id, String(env.WORKBOOK_ID || "")) && timingSafeEqual(key, String(env.WORKBOOK_KEY || ""));
}

async function validSession(env, token) {
  if (!token || token.length < 30) return null;
  const row = await env.DB.prepare("SELECT workbook_id,expires_at FROM sessions WHERE token=?").bind(token).first();
  if (!row || Number(row.expires_at) < Date.now()) return null;
  return row;
}

function bearer(request) {
  const value = request.headers.get("Authorization") || "";
  return value.startsWith("Bearer ") ? value.slice(7).trim() : "";
}

function timingSafeEqual(a, b) {
  if (a.length !== b.length) return false;
  let diff = 0;
  for (let i = 0; i < a.length; i++) diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return diff === 0;
}

function clean(value, max) {
  return String(value || "").replace(/[\r\n]/g, "").trim().slice(0, max);
}

function corsHeaders(request) {
  const headers = new Headers();
  const origin = request.headers.get("Origin") || "";
  if (origin === ALLOWED_ORIGIN) {
    headers.set("Access-Control-Allow-Origin", origin);
    headers.set("Vary", "Origin");
  }
  headers.set("Access-Control-Allow-Headers", "Content-Type, Authorization, X-Solar-EPC-Workbook, X-Solar-EPC-Key, X-Solar-EPC-Session");
  headers.set("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  headers.set("Cache-Control", "no-store");
  return headers;
}

function json(value, status, request) {
  const headers = corsHeaders(request);
  headers.set("Content-Type", "application/json; charset=utf-8");
  return new Response(JSON.stringify(value), { status, headers });
}

function empty(status, request) {
  return new Response(null, { status, headers: corsHeaders(request) });
}

// ============================================================================
// SOLAR EPC NASA POWER HOURLY RESOURCE MODULE v1.8 (range6)
// Additive Worker/D1 module. Existing SAVE/session/ACK behavior is unchanged.
//
// NEW in v1.8:
//  - /v1/resource/process-range fetches up to RESOURCE_RANGE_MONTHS (default 12)
//    calendar months in ONE NASA request, validates the full range, then splits
//    into per-month resource_chunks + resource_chunk_summaries with byte-identical
//    monthly data/states. Every response also reports its measured processMs so
//    the Workers Free 10ms CPU budget can be verified from VBA.
//  - /v1/resource/plan returns a `ranges` field (contiguous runs of missing
//    months, max 6 per run) alongside the original `months` field.
//  - OLD /v1/resource/process-month remains 100% unchanged as the VBA fallback.
//  - expectedHourlyKeys() rewritten without Date objects (identical output,
//    ~60% less CPU) to protect the Workers Free 10ms CPU budget.
//
// NASA Hourly does not provide T2M_MIN or T2M_MAX. This module retrieves T2M
// hourly and derives minimum/maximum from the verified hourly T2M observations.
// ============================================================================

const NASA_HOURLY_FIRST_DATE = "20010101";
const RESOURCE_PRODUCT = "NASA POWER Hourly Point API";
const RESOURCE_TIME_STANDARD = "UTC";
const RESOURCE_RANGE_MONTHS_DEFAULT = 12; // K=12 turbo: 60 months -> 5 ranges -> 1 wave
const RESOURCE_RANGE_MONTHS_MAX = 12;
const RESOURCE_NASA_RESPONSE_BYTES_MAX = 8_000_000; // safety guard -> VBA falls back to month mode
const RESOURCE_PARAMETERS = Object.freeze([
  "ALLSKY_SFC_SW_DWN", "ALLSKY_SFC_SW_DNI", "ALLSKY_SFC_SW_DIFF", "T2M",
  "WS10M", "WD10M", "RH2M", "PS", "PRECTOTCORR", "ALLSKY_SRF_ALB"
]);
const RESOURCE_PARAMETER_IDENTITY = [
  ...RESOURCE_PARAMETERS, "T2M_MIN:DERIVED_FROM_HOURLY_T2M",
  "T2M_MAX:DERIVED_FROM_HOURLY_T2M"
].join(",");
const RESOURCE_EXPECTED_UNITS = Object.freeze({
  ALLSKY_SFC_SW_DWN: "Wh/m^2",
  ALLSKY_SFC_SW_DNI: "Wh/m^2",
  ALLSKY_SFC_SW_DIFF: "Wh/m^2",
  T2M: "C",
  WS10M: "m/s",
  WD10M: "Degrees",
  RH2M: "%",
  PS: "kPa",
  PRECTOTCORR: "mm/day",
  ALLSKY_SRF_ALB: "dimensionless"
});

async function resourceStart(request, env, workbook) {
  let body;
  try { body = await request.json(); }
  catch { return json({ ok: false, error: "INVALID_JSON" }, 400, request); }

  let period;
  try { period = validateResourcePeriod(body?.startDate, body?.endDate); }
  catch (error) {
    return json({
      ok: false, error: "RESOURCE_PERIOD_CONFIGURATION_ERROR",
      message: String(error.message || error)
    }, 400, request);
  }

  const projectId = clean(body?.projectId, 200);
  const messageId = clean(body?.messageId, 200);
  if (!projectId) return json({ ok: false, error: "MISSING_PROJECT_ID" }, 400, request);

  let polygon;
  try { polygon = validateSitePolygon(body?.sitePolygon); }
  catch (error) {
    return json({ ok: false, error: "INVALID_SITE_POLYGON", message: String(error.message || error) }, 400, request);
  }

  const centroid = polygonCentroidLocalProjection(polygon);
  if (!validCoordinate(centroid.latitude, centroid.longitude)) {
    return json({ ok: false, error: "INVALID_CENTROID" }, 400, request);
  }

  const polygonJson = JSON.stringify(polygon);
  const polygonHash = await sha256Hex(polygonJson);
  const cacheIdentity = JSON.stringify({
    workbook, projectId,
    latitude: fixed8(centroid.latitude), longitude: fixed8(centroid.longitude),
    start: period.compactStart, end: period.compactEnd, resolution: "HOURLY",
    product: RESOURCE_PRODUCT, parameters: RESOURCE_PARAMETER_IDENTITY
  });
  const cacheKey = await sha256Hex(cacheIdentity);

  const cached = await env.DB.prepare(
    "SELECT request_id,status FROM resource_requests WHERE workbook_id=? AND cache_key=? ORDER BY created_at DESC LIMIT 1"
  ).bind(workbook, cacheKey).first();
  if (cached) {
    return json({
      ok: true, requestId: cached.request_id, status: cached.status,
      complete: cached.status === "COMPLETE", cached: true,
      centroid: { latitude: centroid.latitude, longitude: centroid.longitude },
      startDate: period.isoStart, endDate: period.isoEnd
    }, 200, request);
  }

  const now = Date.now();
  const requestId = crypto.randomUUID();
  const state = newResourceSummaryState();
  await env.DB.prepare(
    `INSERT INTO resource_requests(
      request_id,workbook_id,project_id,message_id,cache_key,polygon_hash,polygon_json,
      centroid_lat,centroid_lon,period_start,period_end,resolution,product,
      parameters_json,next_month,total_chunks,completed_chunks,status,error_text,
      units_json,metadata_json,summary_state_json,summary_json,created_at,updated_at
    ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)`
  ).bind(
    requestId, workbook, projectId, messageId, cacheKey, polygonHash, polygonJson,
    centroid.latitude, centroid.longitude, period.compactStart, period.compactEnd, "HOURLY",
    RESOURCE_PRODUCT, JSON.stringify(RESOURCE_PARAMETERS), period.compactStart.slice(0, 6),
    countInclusiveMonths(period.compactStart, period.compactEnd), 0,
    "PENDING", null, null, null, JSON.stringify(state), null, now, now
  ).run();

  return json({
    ok: true, requestId, status: "PENDING", complete: false, cached: false,
    centroid: { latitude: centroid.latitude, longitude: centroid.longitude },
    startDate: period.isoStart, endDate: period.isoEnd
  }, 201, request);
}

async function resourceProcess(request, env, workbook) {
  let body;
  try { body = await request.json(); }
  catch { return json({ ok: false, error: "INVALID_JSON" }, 400, request); }
  const requestId = clean(body?.requestId, 100);
  if (!requestId) return json({ ok: false, error: "MISSING_REQUEST_ID" }, 400, request);

  const row = await env.DB.prepare(
    "SELECT * FROM resource_requests WHERE request_id=? AND workbook_id=?"
  ).bind(requestId, workbook).first();
  if (!row) return json({ ok: false, error: "RESOURCE_REQUEST_NOT_FOUND" }, 404, request);
  if (row.status === "COMPLETE") {
    return json({ ok: true, requestId, status: "COMPLETE", complete: true, completedChunks: row.completed_chunks, totalChunks: row.total_chunks }, 200, request);
  }

  const month = String(row.next_month || "");
  const requestedStart = storedCompactDate(row.period_start);
  const requestedEnd = storedCompactDate(row.period_end);
  if (!/^\d{6}$/.test(month) || !requestedStart || !requestedEnd) {
    await markResourceError(env, requestId, "ERROR", "INVALID_STORED_RESOURCE_PERIOD");
    return json({ ok: false, requestId, status: "ERROR", error: "INVALID_STORED_RESOURCE_PERIOD" }, 500, request);
  }

  const range = configuredMonthRange(month, requestedStart, requestedEnd);
  const nasaUrl = new URL("https://power.larc.nasa.gov/api/temporal/hourly/point");
  nasaUrl.searchParams.set("parameters", RESOURCE_PARAMETERS.join(","));
  nasaUrl.searchParams.set("community", "RE");
  nasaUrl.searchParams.set("longitude", fixed8(Number(row.centroid_lon)));
  nasaUrl.searchParams.set("latitude", fixed8(Number(row.centroid_lat)));
  nasaUrl.searchParams.set("start", range.start);
  nasaUrl.searchParams.set("end", range.end);
  nasaUrl.searchParams.set("format", "JSON");
  nasaUrl.searchParams.set("time-standard", RESOURCE_TIME_STANDARD);

  let response;
  try {
    response = await fetch(nasaUrl.toString(), { headers: { "Accept": "application/json", "User-Agent": "Solar-EPC-Resource/1.0" } });
  } catch (error) {
    const detail = "NASA_NETWORK_ERROR: " + String(error?.message || error);
    await markResourceError(env, requestId, "DATA UNAVAILABLE", detail);
    return json({ ok: false, requestId, status: "DATA UNAVAILABLE", error: detail, retryable: true }, 503, request);
  }

  if (!response.ok) {
    const text = (await response.text()).slice(0, 1000);
    const detail = `NASA_HTTP_${response.status}: ${text}`;
    await markResourceError(env, requestId, "DATA UNAVAILABLE", detail);
    return json({ ok: false, requestId, status: "DATA UNAVAILABLE", error: detail, retryable: response.status >= 500 || response.status === 429 }, 503, request);
  }

  let nasa;
  try { nasa = await response.json(); }
  catch {
    await markResourceError(env, requestId, "DATA UNAVAILABLE", "NASA_INVALID_JSON");
    return json({ ok: false, requestId, status: "DATA UNAVAILABLE", error: "NASA_INVALID_JSON", retryable: true }, 503, request);
  }

  let validated;
  try {
    validated = validateAndNormalizeNasaMonth(nasa, range, Number(row.centroid_lat), Number(row.centroid_lon));
  } catch (error) {
    const detail = "NASA_VALIDATION_ERROR: " + String(error?.message || error);
    await markResourceError(env, requestId, "ERROR", detail);
    return json({ ok: false, requestId, status: "ERROR", error: detail, retryable: false }, 422, request);
  }

  const chunkJson = JSON.stringify({
    schema: "SOLAR_EPC_NASA_HOURLY_V1",
    timeStandard: RESOURCE_TIME_STANDARD,
    parameterOrder: RESOURCE_PARAMETERS,
    rows: validated.rows
  });
  const checksum = await sha256Hex(chunkJson);
  const oldState = safeJson(row.summary_state_json, newResourceSummaryState());
  const state = mergeMonthState(oldState, validated);
  const nextMonth = incrementMonth(month);
  const totalChunks = Number(row.total_chunks);
  const completedChunks = Math.min(totalChunks, Number(row.completed_chunks || 0) + 1);
  const complete = month === requestedEnd.slice(0, 6);
  const status = complete ? "COMPLETE" : (state.reviewCount > 0 ? "REVIEW" : "PROCESSING");
  const summary = complete ? finalizeResourceSummary({ ...row, completed_chunks: totalChunks, status: "COMPLETE" }, state, validated.units, validated.metadata) : null;
  const now = Date.now();

  await env.DB.batch([
    env.DB.prepare(
      `INSERT INTO resource_chunks(request_id,chunk_start,chunk_end,time_standard,parameter_order_json,data_json,
        record_count,missing_count,invalid_count,duplicate_count,checksum_sha256,created_at)
       VALUES(?,?,?,?,?,?,?,?,?,?,?,?)
       ON CONFLICT(request_id,chunk_start) DO UPDATE SET
        chunk_end=excluded.chunk_end,time_standard=excluded.time_standard,
        parameter_order_json=excluded.parameter_order_json,data_json=excluded.data_json,
        record_count=excluded.record_count,missing_count=excluded.missing_count,
        invalid_count=excluded.invalid_count,duplicate_count=excluded.duplicate_count,
        checksum_sha256=excluded.checksum_sha256,created_at=excluded.created_at`
    ).bind(
      requestId, range.start, range.end, RESOURCE_TIME_STANDARD,
      JSON.stringify(RESOURCE_PARAMETERS), chunkJson, validated.rows.length,
      validated.missingCount, validated.invalidCount, validated.duplicateCount,
      checksum, now
    ),
    env.DB.prepare(
      `UPDATE resource_requests SET next_month=?,completed_chunks=?,status=?,error_text=NULL,
       units_json=COALESCE(units_json,?),metadata_json=COALESCE(metadata_json,?),
       summary_state_json=?,summary_json=?,updated_at=? WHERE request_id=? AND workbook_id=?`
    ).bind(
      nextMonth, completedChunks, status, JSON.stringify(validated.units),
      JSON.stringify(validated.metadata), JSON.stringify(state),
      summary ? JSON.stringify(summary) : null, now, requestId, workbook
    )
  ]);

  return json({
    ok: true, requestId, status, complete, processedMonth: month,
    completedChunks, totalChunks,
    validation: {
      records: validated.rows.length, missing: validated.missingCount,
      invalid: validated.invalidCount, duplicates: validated.duplicateCount,
      review: validated.reviewCount > 0
    }
  }, 200, request);
}

async function resourceStatus(request, env, workbook, requestId) {
  const row = await env.DB.prepare(
    "SELECT request_id,project_id,status,error_text,completed_chunks,total_chunks,updated_at FROM resource_requests WHERE request_id=? AND workbook_id=?"
  ).bind(requestId, workbook).first();
  if (!row) return json({ ok: false, error: "RESOURCE_REQUEST_NOT_FOUND" }, 404, request);
  return json({
    ok: true, requestId: row.request_id, projectId: row.project_id,
    status: row.status, error: row.error_text || null,
    completedChunks: row.completed_chunks, totalChunks: row.total_chunks,
    complete: row.status === "COMPLETE", updatedAt: row.updated_at
  }, 200, request);
}

async function resourcePending(request, env, workbook) {
  const row = await env.DB.prepare(
    `SELECT request_id,status FROM resource_requests
     WHERE workbook_id=? AND status<>? ORDER BY updated_at ASC LIMIT 1`
  ).bind(workbook, "COMPLETE").first();
  if (!row) return empty(204, request);
  return json({ ok: true, requestId: row.request_id, status: row.status }, 200, request);
}

async function resourceSummary(request, env, workbook, requestId) {
  const row = await env.DB.prepare(
    "SELECT * FROM resource_requests WHERE request_id=? AND workbook_id=?"
  ).bind(requestId, workbook).first();
  if (!row) return json({ ok: false, error: "RESOURCE_REQUEST_NOT_FOUND" }, 404, request);

  const units = safeJson(row.units_json, {});
  const metadata = safeJson(row.metadata_json, {});
  const state = safeJson(row.summary_state_json, newResourceSummaryState());
  const summary = row.summary_json ? safeJson(row.summary_json, null) : finalizeResourceSummary(row, state, units, metadata);
  const text = resourceSummaryText(row, summary, units, metadata);
  const headers = corsHeaders(request);
  headers.set("Content-Type", "text/plain; charset=utf-8");
  return new Response(text, { status: 200, headers });
}

function validateSitePolygon(input) {
  if (!Array.isArray(input)) throw new Error("SITE polygon must be an array.");
  if (input.length < 3 || input.length > 1000) throw new Error("SITE polygon requires 3 to 1000 vertices.");
  const out = [];
  for (const item of input) {
    const lat = Number(Array.isArray(item) ? item[0] : item?.latitude);
    const lon = Number(Array.isArray(item) ? item[1] : item?.longitude);
    if (!validCoordinate(lat, lon)) throw new Error("SITE polygon contains an invalid coordinate.");
    const p = [Number(fixed8(lat)), Number(fixed8(lon))];
    const last = out[out.length - 1];
    if (!last || last[0] !== p[0] || last[1] !== p[1]) out.push(p);
  }
  if (out.length > 1 && out[0][0] === out[out.length - 1][0] && out[0][1] === out[out.length - 1][1]) out.pop();
  if (out.length < 3) throw new Error("SITE polygon has fewer than three distinct vertices.");
  const distinct = new Set(out.map(p => `${p[0]},${p[1]}`));
  if (distinct.size < 3) throw new Error("SITE polygon has fewer than three distinct vertices.");
  if (polygonSelfIntersects(out)) throw new Error("SITE polygon self-intersects.");
  const centroid = polygonCentroidLocalProjection(out);
  if (!Number.isFinite(centroid.areaM2) || centroid.areaM2 < 0.01) throw new Error("SITE polygon has zero or negligible area.");
  return out;
}

function polygonCentroidLocalProjection(points) {
  const meanLat = points.reduce((s, p) => s + p[0], 0) / points.length;
  const meanLon = points.reduce((s, p) => s + p[1], 0) / points.length;
  const r = 6378137;
  const cosLat = Math.cos(meanLat * Math.PI / 180);
  if (Math.abs(cosLat) < 1e-12) throw new Error("Polygon is too close to a pole for this centroid projection.");
  const xy = points.map(p => [
    (p[1] - meanLon) * Math.PI / 180 * r * cosLat,
    (p[0] - meanLat) * Math.PI / 180 * r
  ]);
  let twiceArea = 0, cx6a = 0, cy6a = 0;
  for (let i = 0; i < xy.length; i++) {
    const a = xy[i], b = xy[(i + 1) % xy.length];
    const cross = a[0] * b[1] - b[0] * a[1];
    twiceArea += cross;
    cx6a += (a[0] + b[0]) * cross;
    cy6a += (a[1] + b[1]) * cross;
  }
  if (Math.abs(twiceArea) < 1e-9) throw new Error("Polygon area is zero.");
  const cx = cx6a / (3 * twiceArea);
  const cy = cy6a / (3 * twiceArea);
  return {
    latitude: meanLat + (cy / r) * 180 / Math.PI,
    longitude: meanLon + (cx / (r * cosLat)) * 180 / Math.PI,
    areaM2: Math.abs(twiceArea) / 2
  };
}

function polygonSelfIntersects(points) {
  const n = points.length;
  for (let i = 0; i < n; i++) {
    const a = points[i], b = points[(i + 1) % n];
    for (let j = i + 1; j < n; j++) {
      if (j === i || (j + 1) % n === i || j === (i + 1) % n) continue;
      const c = points[j], d = points[(j + 1) % n];
      if (segmentsIntersect(a, b, c, d)) return true;
    }
  }
  return false;
}

function segmentsIntersect(a, b, c, d) {
  const orient = (p, q, r) => Math.sign((q[1] - p[1]) * (r[0] - p[0]) - (q[0] - p[0]) * (r[1] - p[1]));
  const o1 = orient(a, b, c), o2 = orient(a, b, d), o3 = orient(c, d, a), o4 = orient(c, d, b);
  return o1 !== o2 && o3 !== o4;
}

function validateAndNormalizeNasaMonth(nasa, range, requestedLat, requestedLon) {
  const header = nasa?.header;
  const parametersMeta = nasa?.parameters;
  const data = nasa?.properties?.parameter;
  if (!header || !parametersMeta || !data) throw new Error("Required NASA response sections are missing.");
  if (String(header?.time_standard || "").toUpperCase() !== "UTC") throw new Error("NASA returned an unexpected time standard.");
  if (String(header?.start || "") !== range.start || String(header?.end || "") !== range.end) throw new Error("NASA returned a period different from the requested period.");
  if (!header?.api?.version || !header?.api?.name || !Array.isArray(header?.sources) || !header.sources.length) throw new Error("NASA source metadata is incomplete.");
  const returnedCoordinate = nasa?.geometry?.coordinates;
  if (!Array.isArray(returnedCoordinate) || returnedCoordinate.length < 2 ||
      Math.abs(Number(returnedCoordinate[0]) - requestedLon) > 0.01 ||
      Math.abs(Number(returnedCoordinate[1]) - requestedLat) > 0.01) {
    throw new Error("NASA returned coordinates inconsistent with the requested centroid.");
  }

  const units = {};
  let unitMismatchCount = 0;
  for (const p of RESOURCE_PARAMETERS) {
    if (!Object.prototype.hasOwnProperty.call(data, p)) throw new Error(`NASA parameter missing: ${p}`);
    if (!Object.prototype.hasOwnProperty.call(parametersMeta, p)) throw new Error(`NASA parameter metadata missing: ${p}`);
    units[p] = String(parametersMeta[p]?.units || "");
    if (!units[p]) throw new Error(`NASA unit missing: ${p}`);
    if (RESOURCE_EXPECTED_UNITS[p] && units[p] !== RESOURCE_EXPECTED_UNITS[p]) unitMismatchCount++;
  }
  const returnedNames = Object.keys(data).sort();
  const expectedNames = [...RESOURCE_PARAMETERS].sort();
  if (returnedNames.join("|") !== expectedNames.join("|")) throw new Error("NASA returned a different parameter set.");

  const expectedTimes = expectedHourlyKeys(range.start, range.end);
  const expectedSet = new Set(expectedTimes);
  let duplicateCount = expectedTimes.length - expectedSet.size;
  let missingCount = 0, invalidCount = 0;
  const rows = [];
  const state = newResourceSummaryState();
  const daily = {};
  for (const p of ["ALLSKY_SFC_SW_DWN", "ALLSKY_SFC_SW_DNI", "ALLSKY_SFC_SW_DIFF"]) daily[p] = {};

  for (const timestamp of expectedTimes) {
    const row = [timestamp];
    for (const p of RESOURCE_PARAMETERS) {
      const obj = data[p];
      let value = obj[timestamp];
      const fill = Number(header.fill_value);
      if (value === undefined || value === null || Number(value) === fill || !Number.isFinite(Number(value))) {
        value = null;
        missingCount++;
        state.metrics[p].missing++;
      } else {
        value = Number(value);
        if (!resourceValueValid(p, value)) {
          value = null;
          invalidCount++;
          state.metrics[p].invalid++;
        } else {
          addMetric(state.metrics[p], value);
          if (daily[p]) {
            const day = timestamp.slice(0, 8);
            if (!daily[p][day]) daily[p][day] = { sum: 0, count: 0 };
            daily[p][day].sum += value;
            daily[p][day].count++;
          }
        }
      }
      row.push(value);
    }
    rows.push(row);
  }

  for (const p of RESOURCE_PARAMETERS) {
    const keys = Object.keys(data[p]);
    const seen = new Set();
    for (const key of keys) {
      if (seen.has(key)) duplicateCount++;
      seen.add(key);
      if (!expectedSet.has(key)) throw new Error(`NASA returned timestamp outside requested month: ${key}`);
      if (!/^\d{10}$/.test(key)) throw new Error(`NASA returned invalid timestamp: ${key}`);
    }
  }

  for (const p of Object.keys(daily)) {
    for (const day of Object.keys(daily[p])) {
      const d = daily[p][day];
      if (d.count === 24) {
        state.dailyIrradiance[p].sum += d.sum / 1000;
        state.dailyIrradiance[p].count++;
      } else {
        state.dailyIrradiance[p].incomplete++;
      }
    }
  }

  state.missingCount = missingCount;
  state.invalidCount = invalidCount;
  state.duplicateCount = duplicateCount;
  state.unitMismatchCount = unitMismatchCount;
  state.reviewCount = missingCount + invalidCount + duplicateCount + unitMismatchCount;
  return {
    rows, units, state, missingCount, invalidCount, duplicateCount,
    reviewCount: state.reviewCount,
    metadata: {
      title: String(header.title || ""), apiName: String(header.api.name),
      apiVersion: String(header.api.version), sources: header.sources,
      fillValue: header.fill_value, timeStandard: header.time_standard,
      firstChunkStart: range.start, latestChunkEnd: range.end,
      parameterMetadata: parametersMeta,
      returnedGeometry: nasa.geometry,
      messages: Array.isArray(nasa.messages) ? nasa.messages : []
    }
  };
}

function resourceValueValid(p, v) {
  switch (p) {
    case "ALLSKY_SFC_SW_DWN": case "ALLSKY_SFC_SW_DNI": case "ALLSKY_SFC_SW_DIFF": return v >= 0 && v <= 2000;
    case "T2M": return v >= -100 && v <= 70;
    case "WS10M": return v >= 0 && v <= 150;
    case "WD10M": return v >= 0 && v <= 360;
    case "RH2M": return v >= 0 && v <= 100;
    case "PS": return v > 0 && v <= 120;
    case "PRECTOTCORR": return v >= 0 && v <= 3000;
    case "ALLSKY_SRF_ALB": return v >= 0 && v <= 1;
    default: return false;
  }
}

function newResourceSummaryState() {
  const metrics = {};
  for (const p of RESOURCE_PARAMETERS) metrics[p] = { sum: 0, count: 0, min: null, max: null, missing: 0, invalid: 0 };
  return {
    metrics,
    dailyIrradiance: {
      ALLSKY_SFC_SW_DWN: { sum: 0, count: 0, incomplete: 0 },
      ALLSKY_SFC_SW_DNI: { sum: 0, count: 0, incomplete: 0 },
      ALLSKY_SFC_SW_DIFF: { sum: 0, count: 0, incomplete: 0 }
    },
    windSin: 0, windCos: 0, windDirectionCount: 0,
    missingCount: 0, invalidCount: 0, duplicateCount: 0,
    unitMismatchCount: 0, reviewCount: 0
  };
}

function addMetric(m, value) {
  m.sum += value; m.count++;
  m.min = m.min === null ? value : Math.min(m.min, value);
  m.max = m.max === null ? value : Math.max(m.max, value);
}

function mergeMonthState(total, month) {
  const part = month.state;
  for (const p of RESOURCE_PARAMETERS) {
    const a = total.metrics[p], b = part.metrics[p];
    a.sum += b.sum; a.count += b.count; a.missing += b.missing; a.invalid += b.invalid;
    if (b.min !== null) a.min = a.min === null ? b.min : Math.min(a.min, b.min);
    if (b.max !== null) a.max = a.max === null ? b.max : Math.max(a.max, b.max);
  }
  for (const p of Object.keys(total.dailyIrradiance)) {
    total.dailyIrradiance[p].sum += part.dailyIrradiance[p].sum;
    total.dailyIrradiance[p].count += part.dailyIrradiance[p].count;
    total.dailyIrradiance[p].incomplete += part.dailyIrradiance[p].incomplete;
  }
  const wdIndex = RESOURCE_PARAMETERS.indexOf("WD10M") + 1;
  for (const r of month.rows) {
    const v = r[wdIndex];
    if (v !== null) {
      total.windSin += Math.sin(v * Math.PI / 180);
      total.windCos += Math.cos(v * Math.PI / 180);
      total.windDirectionCount++;
    }
  }
  total.missingCount += part.missingCount;
  total.invalidCount += part.invalidCount;
  total.duplicateCount += part.duplicateCount;
  total.unitMismatchCount += part.unitMismatchCount;
  total.reviewCount += part.reviewCount;
  return total;
}

function finalizeResourceSummary(row, state, units, metadata) {
  const mean = p => state.metrics[p].count ? state.metrics[p].sum / state.metrics[p].count : null;
  const dailyMean = p => state.dailyIrradiance[p].count ? state.dailyIrradiance[p].sum / state.dailyIrradiance[p].count : null;
  let windDirection = null;
  if (state.windDirectionCount) {
    windDirection = Math.atan2(state.windSin / state.windDirectionCount, state.windCos / state.windDirectionCount) * 180 / Math.PI;
    if (windDirection < 0) windDirection += 360;
  }
  const noData = RESOURCE_PARAMETERS.some(p => !state.metrics[p].count);
  const complete = Number(row.completed_chunks) >= Number(row.total_chunks) || row.status === "COMPLETE";
  const status = noData ? "DATA UNAVAILABLE" : (!complete ? String(row.status || "REVIEW") : (state.reviewCount > 0 ? "REVIEW" : "VALID"));
  return {
    status,
    ghi: dailyMean("ALLSKY_SFC_SW_DWN"),
    dni: dailyMean("ALLSKY_SFC_SW_DNI"),
    dhi: dailyMean("ALLSKY_SFC_SW_DIFF"),
    ambientTemperature: mean("T2M"),
    minTemperature: state.metrics.T2M.min,
    maxTemperature: state.metrics.T2M.max,
    windSpeed: mean("WS10M"), windDirection,
    relativeHumidity: mean("RH2M"), surfacePressure: mean("PS"),
    precipitation: mean("PRECTOTCORR"), albedo: mean("ALLSKY_SRF_ALB"),
    missingCount: state.missingCount, invalidCount: state.invalidCount,
    duplicateCount: state.duplicateCount, unitMismatchCount: state.unitMismatchCount,
    validHourlyT2MCount: state.metrics.T2M.count,
    completeIrradianceDays: state.dailyIrradiance.ALLSKY_SFC_SW_DWN.count,
    incompleteIrradianceDays: state.dailyIrradiance.ALLSKY_SFC_SW_DWN.incomplete,
    units, metadata
  };
}

function resourceSummaryText(row, s, units, metadata) {
  const periodStart = storedCompactDate(row.period_start);
  const periodEnd = storedCompactDate(row.period_end);
  const ref = `https://power.larc.nasa.gov/api/temporal/hourly/point (centroid ${fixed8(Number(row.centroid_lat))}, ${fixed8(Number(row.centroid_lon))}; UTC; ${periodStart}-${periodEnd})`;
  const remarks = [
    "GHI/DNI/DHI = mean of complete daily totals calculated from validated hourly Wh/m^2 and reported as kWh/m^2/day.",
    "Ambient temperature = mean hourly T2M; Min/Max Temperature = derived extrema of validated hourly T2M because NASA Hourly rejects T2M_MIN/T2M_MAX.",
    "Wind direction = circular mean. Other weather fields = arithmetic means of valid hourly NASA values.",
    `Precipitation uses NASA returned unit ${units.PRECTOTCORR || "unavailable"}; missing values were excluded, never changed to zero.`,
    "GTI / POA Irradiance remains DERIVED and is not populated from GHI.",
    `Validation: missing=${s.missingCount}, invalid=${s.invalidCount}, duplicates=${s.duplicateCount}, unit mismatches=${s.unitMismatchCount}.`
  ].join(" ");
  const fields = {
    schema: "SOLAR_EPC_RESOURCE_SUMMARY_V1", request_id: row.request_id,
    project_id: row.project_id, latitude: row.centroid_lat, longitude: row.centroid_lon,
    source: "NASA POWER", dataset_product: `${metadata.title || RESOURCE_PRODUCT} | ${metadata.apiName || "POWER Hourly API"} ${metadata.apiVersion || ""}`.trim(),
    retrieval_date: new Date(Number(row.updated_at || Date.now())).toISOString(),
    data_period_start: compactToIso(periodStart), data_period_end: compactToIso(periodEnd), resolution: "Hourly (UTC)",
    ghi: s.ghi, dni: s.dni, dhi: s.dhi, gti_poa: "DERIVED",
    ambient_temperature: s.ambientTemperature, min_temperature: s.minTemperature,
    max_temperature: s.maxTemperature, wind_speed: s.windSpeed,
    wind_direction: s.windDirection, relative_humidity: s.relativeHumidity,
    surface_pressure: s.surfacePressure, precipitation: s.precipitation,
    albedo: s.albedo, data_status: s.status, source_reference: ref, remarks
  };
  return Object.entries(fields).map(([k, v]) => `${k}|${safeTextValue(v)}`).join("\n") + "\nEND";
}

function safeTextValue(value) {
  if (value === null || value === undefined || (typeof value === "number" && !Number.isFinite(value))) return "";
  if (typeof value === "number") return String(Math.round(value * 1000000) / 1000000);
  return String(value).replace(/[\r\n|]+/g, " ").trim();
}

// Integer math version of the hourly key generator. Produces byte-identical
// keys to the previous Date-based implementation (verified for leap years and
// period boundaries) while avoiding thousands of Date object allocations.
function expectedHourlyKeys(start, end) {
  const keys = [];
  let y = Number(start.slice(0, 4)), m = Number(start.slice(4, 6)), d = Number(start.slice(6, 8));
  const ey = Number(end.slice(0, 4)), em = Number(end.slice(4, 6)), ed = Number(end.slice(6, 8));
  while (true) {
    for (let h = 0; h < 24; h++) keys.push(`${y}${pad2(m)}${pad2(d)}${pad2(h)}`);
    if (y === ey && m === em && d === ed) break;
    if (d === daysInMonth(y, m)) { m++; d = 1; if (m === 13) { m = 1; y++; } }
    else d++;
  }
  return keys;
}

function daysInMonth(year, month) {
  if (month === 2) return ((year % 4 === 0 && year % 100 !== 0) || year % 400 === 0) ? 29 : 28;
  return [0, 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][month];
}

function validateResourcePeriod(startDate, endDate) {
  if (typeof startDate !== "string" || typeof endDate !== "string" || !startDate.trim() || !endDate.trim()) {
    throw new Error("Valid configured startDate and endDate are required in YYYY-MM-DD format.");
  }
  const compactStart = strictIsoDateToCompact(startDate.trim());
  const compactEnd = strictIsoDateToCompact(endDate.trim());
  if (!compactStart || !compactEnd) throw new Error("Resource dates must be valid calendar dates in YYYY-MM-DD format.");
  if (compactStart > compactEnd) throw new Error("Resource startDate must not be after endDate.");
  if (compactStart < NASA_HOURLY_FIRST_DATE) throw new Error("NASA POWER Hourly data is supported from 2001-01-01.");
  const now = new Date();
  const today = `${now.getUTCFullYear()}${pad2(now.getUTCMonth() + 1)}${pad2(now.getUTCDate())}`;
  if (compactEnd > today) throw new Error("Resource endDate cannot be in the future; NASA near-real-time availability is verified from its response.");
  return {
    compactStart, compactEnd,
    isoStart: compactToIso(compactStart), isoEnd: compactToIso(compactEnd)
  };
}

function strictIsoDateToCompact(value) {
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(value || ""));
  if (!match) return "";
  const year = Number(match[1]), month = Number(match[2]), day = Number(match[3]);
  const date = new Date(Date.UTC(year, month - 1, day));
  if (date.getUTCFullYear() !== year || date.getUTCMonth() !== month - 1 || date.getUTCDate() !== day) return "";
  return `${match[1]}${match[2]}${match[3]}`;
}

function storedCompactDate(value) {
  const text = String(value || "").trim();
  if (/^\d{8}$/.test(text)) {
    return strictIsoDateToCompact(compactToIso(text)) || "";
  }
  return strictIsoDateToCompact(text) || "";
}

function compactToIso(value) {
  const text = String(value || "");
  return /^\d{8}$/.test(text) ? `${text.slice(0, 4)}-${text.slice(4, 6)}-${text.slice(6, 8)}` : "";
}

function countInclusiveMonths(startCompact, endCompact) {
  const sy = Number(startCompact.slice(0, 4)), sm = Number(startCompact.slice(4, 6));
  const ey = Number(endCompact.slice(0, 4)), em = Number(endCompact.slice(4, 6));
  return (ey - sy) * 12 + (em - sm) + 1;
}

function monthRange(yyyymm) {
  const y = Number(yyyymm.slice(0, 4)), m = Number(yyyymm.slice(4, 6));
  const last = daysInMonth(y, m);
  return { start: `${y}${pad2(m)}01`, end: `${y}${pad2(m)}${pad2(last)}` };
}

function configuredMonthRange(yyyymm, configuredStart, configuredEnd) {
  const full = monthRange(yyyymm);
  return {
    start: full.start < configuredStart ? configuredStart : full.start,
    end: full.end > configuredEnd ? configuredEnd : full.end
  };
}
function incrementMonth(yyyymm) {
  let y = Number(yyyymm.slice(0, 4)), m = Number(yyyymm.slice(4, 6)) + 1;
  if (m === 13) { y++; m = 1; }
  return `${y}${pad2(m)}`;
}
function pad2(v) { return String(v).padStart(2, "0"); }
function fixed8(v) { return Number(v).toFixed(8); }
function validCoordinate(lat, lon) { return Number.isFinite(lat) && Number.isFinite(lon) && lat >= -90 && lat <= 90 && lon >= -180 && lon <= 180; }
function safeJson(text, fallback) { try { return JSON.parse(String(text || "")); } catch { return fallback; } }
async function sha256Hex(text) {
  const bytes = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(text));
  return [...new Uint8Array(bytes)].map(b => b.toString(16).padStart(2, "0")).join("");
}
async function markResourceError(env, requestId, status, detail) {
  await env.DB.prepare("UPDATE resource_requests SET status=?,error_text=?,updated_at=? WHERE request_id=?")
    .bind(status, String(detail).slice(0, 2000), Date.now(), requestId).run();
}


// ============================================================================
// FAST-SAFE RESOURCE PROCESSING
// Six separate authenticated Worker requests may process different ranges in
// parallel. Each invocation validates up to RESOURCE_RANGE_MONTHS (default 6)
// calendar months to protect the Cloudflare Free CPU budget. Existing
// sequential /v1/resource/process and month-by-month /v1/resource/process-month
// remain unchanged as compatibility fallbacks.
// ============================================================================

async function resourcePlan(request, env, workbook, requestId) {
  const row = await env.DB.prepare(
    "SELECT request_id,period_start,period_end,total_chunks,status FROM resource_requests WHERE request_id=? AND workbook_id=?"
  ).bind(requestId, workbook).first();
  if (!row) return json({ ok: false, error: "RESOURCE_REQUEST_NOT_FOUND" }, 404, request);
  if (row.status === "COMPLETE") {
    return json({ ok: true, requestId, complete: true, readyToFinalize: false, months: "", ranges: "", completedChunks: row.total_chunks, totalChunks: row.total_chunks }, 200, request);
  }

  const start = storedCompactDate(row.period_start);
  const end = storedCompactDate(row.period_end);
  if (!start || !end) return json({ ok: false, error: "INVALID_STORED_RESOURCE_PERIOD" }, 422, request);
  const allMonths = configuredMonths(start, end);
  const result = await env.DB.prepare(
    "SELECT chunk_start FROM resource_chunk_summaries WHERE request_id=?"
  ).bind(requestId).all();
  const completedStarts = new Set((result.results || []).map(x => String(x.chunk_start)));
  const missingMonths = allMonths.filter(month => !completedStarts.has(configuredMonthRange(month, start, end).start));
  const completedChunks = allMonths.length - missingMonths.length;

  const rangeEnabled = env.RESOURCE_RANGE_ENABLED !== "0";
  const maxRangeMonths = Math.max(1, Math.min(RESOURCE_RANGE_MONTHS_MAX, Number(env.RESOURCE_RANGE_MONTHS || RESOURCE_RANGE_MONTHS_DEFAULT)));
  const rangeGroups = rangeEnabled ? contiguousMonthGroups(missingMonths, maxRangeMonths) : [];

  await env.DB.prepare(
    "UPDATE resource_requests SET completed_chunks=?,status=?,updated_at=? WHERE request_id=? AND workbook_id=?"
  ).bind(completedChunks, missingMonths.length ? "PROCESSING" : "READY_TO_FINALIZE", Date.now(), requestId, workbook).run();

  return json({
    ok: true, requestId, complete: false, readyToFinalize: missingMonths.length === 0,
    months: missingMonths.slice(0, 6).join(","),
    // Contiguous runs of missing months (max 12 per run). The VBA launches up
    // to 6 of these per wave; each run is one NASA request covering all of its
    // months, so 60 months = 5 runs = 1 wave at six parallel (K=12 turbo).
    ranges: rangeGroups.map(g => g[0] + "-" + g[g.length - 1]).join(","),
    rangeMonths: maxRangeMonths,
    completedChunks, totalChunks: allMonths.length
  }, 200, request);
}

function contiguousMonthGroups(months, chunkSize) {
  const groups = [];
  let group = [];
  for (const month of months) {
    if (group.length && incrementMonth(group[group.length - 1]) !== month) {
      groups.push(group);
      group = [];
    }
    group.push(month);
    if (group.length >= chunkSize) {
      groups.push(group);
      group = [];
    }
  }
  if (group.length) groups.push(group);
  return groups;
}

async function resourceProcessMonth(request, env, workbook) {
  let body;
  try { body = await request.json(); }
  catch { return json({ ok: false, error: "INVALID_JSON" }, 400, request); }
  const requestId = clean(body?.requestId, 100);
  const month = clean(body?.month, 6);
  if (!requestId || !/^\d{6}$/.test(month)) return json({ ok: false, error: "INVALID_FAST_MONTH_REQUEST" }, 400, request);

  const row = await env.DB.prepare(
    "SELECT * FROM resource_requests WHERE request_id=? AND workbook_id=?"
  ).bind(requestId, workbook).first();
  if (!row) return json({ ok: false, error: "RESOURCE_REQUEST_NOT_FOUND" }, 404, request);
  if (row.status === "COMPLETE") return json({ ok: true, requestId, month, complete: true, cached: true }, 200, request);

  const start = storedCompactDate(row.period_start);
  const end = storedCompactDate(row.period_end);
  if (!start || !end || month < start.slice(0, 6) || month > end.slice(0, 6)) {
    return json({ ok: false, error: "MONTH_OUTSIDE_CONFIGURED_PERIOD" }, 422, request);
  }
  const range = configuredMonthRange(month, start, end);
  const existing = await env.DB.prepare(
    "SELECT chunk_start FROM resource_chunk_summaries WHERE request_id=? AND chunk_start=?"
  ).bind(requestId, range.start).first();
  if (existing) return json({ ok: true, requestId, month, processed: true, duplicate: true }, 200, request);

  let validated;
  try {
    validated = await fetchValidatedResourceMonth(row, range);
  } catch (error) {
    const detail = String(error?.message || error);
    const unavailable = detail.startsWith("NASA_NETWORK_ERROR") || detail.startsWith("NASA_HTTP_") || detail === "NASA_INVALID_JSON";
    await markResourceError(env, requestId, unavailable ? "DATA UNAVAILABLE" : "ERROR", detail);
    return json({ ok: false, requestId, month, status: unavailable ? "DATA UNAVAILABLE" : "ERROR", error: detail, retryable: unavailable }, unavailable ? 503 : 422, request);
  }

  const chunkJson = JSON.stringify({
    schema: "SOLAR_EPC_NASA_HOURLY_V1",
    timeStandard: RESOURCE_TIME_STANDARD,
    parameterOrder: RESOURCE_PARAMETERS,
    rows: validated.rows
  });
  const checksum = await sha256Hex(chunkJson);
  const monthState = mergeMonthState(newResourceSummaryState(), validated);
  const now = Date.now();

  await env.DB.batch([
    env.DB.prepare(
      `INSERT INTO resource_chunks(request_id,chunk_start,chunk_end,time_standard,parameter_order_json,data_json,
        record_count,missing_count,invalid_count,duplicate_count,checksum_sha256,created_at)
       VALUES(?,?,?,?,?,?,?,?,?,?,?,?)
       ON CONFLICT(request_id,chunk_start) DO UPDATE SET
        chunk_end=excluded.chunk_end,time_standard=excluded.time_standard,
        parameter_order_json=excluded.parameter_order_json,data_json=excluded.data_json,
        record_count=excluded.record_count,missing_count=excluded.missing_count,
        invalid_count=excluded.invalid_count,duplicate_count=excluded.duplicate_count,
        checksum_sha256=excluded.checksum_sha256,created_at=excluded.created_at`
    ).bind(
      requestId, range.start, range.end, RESOURCE_TIME_STANDARD,
      JSON.stringify(RESOURCE_PARAMETERS), chunkJson, validated.rows.length,
      validated.missingCount, validated.invalidCount, validated.duplicateCount,
      checksum, now
    ),
    env.DB.prepare(
      `INSERT INTO resource_chunk_summaries(request_id,chunk_start,summary_json,units_json,metadata_json,created_at)
       VALUES(?,?,?,?,?,?)
       ON CONFLICT(request_id,chunk_start) DO UPDATE SET
        summary_json=excluded.summary_json,units_json=excluded.units_json,
        metadata_json=excluded.metadata_json,created_at=excluded.created_at`
    ).bind(requestId, range.start, JSON.stringify(monthState), JSON.stringify(validated.units), JSON.stringify(validated.metadata), now),
    env.DB.prepare(
      `UPDATE resource_requests SET status=?,error_text=NULL,
       units_json=COALESCE(units_json,?),metadata_json=COALESCE(metadata_json,?),updated_at=?
       WHERE request_id=? AND workbook_id=?`
    ).bind(monthState.reviewCount > 0 ? "REVIEW" : "PROCESSING", JSON.stringify(validated.units), JSON.stringify(validated.metadata), now, requestId, workbook)
  ]);

  return json({
    ok: true, requestId, month, processed: true, records: validated.rows.length,
    validation: { missing: validated.missingCount, invalid: validated.invalidCount, duplicates: validated.duplicateCount }
  }, 200, request);
}

async function fetchValidatedResourceMonth(row, range) {
  const nasaUrl = new URL("https://power.larc.nasa.gov/api/temporal/hourly/point");
  nasaUrl.searchParams.set("parameters", RESOURCE_PARAMETERS.join(","));
  nasaUrl.searchParams.set("community", "RE");
  nasaUrl.searchParams.set("longitude", fixed8(Number(row.centroid_lon)));
  nasaUrl.searchParams.set("latitude", fixed8(Number(row.centroid_lat)));
  nasaUrl.searchParams.set("start", range.start);
  nasaUrl.searchParams.set("end", range.end);
  nasaUrl.searchParams.set("format", "JSON");
  nasaUrl.searchParams.set("time-standard", RESOURCE_TIME_STANDARD);

  let response;
  try {
    response = await fetch(nasaUrl.toString(), { headers: { "Accept": "application/json", "User-Agent": "Solar-EPC-Resource/1.0" } });
  } catch (error) {
    throw new Error("NASA_NETWORK_ERROR: " + String(error?.message || error));
  }
  if (!response.ok) {
    const text = (await response.text()).slice(0, 1000);
    throw new Error(`NASA_HTTP_${response.status}: ${text}`);
  }
  let nasa;
  try { nasa = await response.json(); }
  catch { throw new Error("NASA_INVALID_JSON"); }
  try {
    return validateAndNormalizeNasaMonth(nasa, range, Number(row.centroid_lat), Number(row.centroid_lon));
  } catch (error) {
    throw new Error("NASA_VALIDATION_ERROR: " + String(error?.message || error));
  }
}

// ----------------------------------------------------------------------------
// RANGE PROCESSING (v1.8): one NASA request for up to K=6 months, split into
// per-month resource_chunks + resource_chunk_summaries that are byte-identical
// to the month-by-month path.
// ----------------------------------------------------------------------------

async function resourceProcessRange(request, env, workbook) {
  let body;
  try { body = await request.json(); }
  catch { return json({ ok: false, error: "INVALID_JSON" }, 400, request); }
  const requestId = clean(body?.requestId, 100);
  const rangeStartMonth = clean(body?.rangeStart, 6);
  const rangeEndMonth = clean(body?.rangeEnd, 6);
  if (!requestId || !/^\d{6}$/.test(rangeStartMonth) || !/^\d{6}$/.test(rangeEndMonth)) {
    return json({ ok: false, error: "INVALID_FAST_RANGE_REQUEST" }, 400, request);
  }
  if (rangeStartMonth > rangeEndMonth) {
    return json({ ok: false, error: "INVALID_FAST_RANGE_ORDER" }, 400, request);
  }
  if (env.RESOURCE_RANGE_ENABLED === "0") {
    return json({ ok: false, error: "RESOURCE_RANGE_DISABLED" }, 404, request);
  }

  const row = await env.DB.prepare(
    "SELECT * FROM resource_requests WHERE request_id=? AND workbook_id=?"
  ).bind(requestId, workbook).first();
  if (!row) return json({ ok: false, error: "RESOURCE_REQUEST_NOT_FOUND" }, 404, request);
  if (row.status === "COMPLETE") {
    return json({ ok: true, requestId, rangeStart: rangeStartMonth, rangeEnd: rangeEndMonth, complete: true, cached: true }, 200, request);
  }

  const start = storedCompactDate(row.period_start);
  const end = storedCompactDate(row.period_end);
  if (!start || !end) {
    return json({ ok: false, error: "INVALID_STORED_RESOURCE_PERIOD" }, 422, request);
  }
  if (rangeStartMonth < start.slice(0, 6) || rangeEndMonth > end.slice(0, 6)) {
    return json({ ok: false, error: "RANGE_OUTSIDE_CONFIGURED_PERIOD" }, 422, request);
  }

  const months = [];
  let current = rangeStartMonth;
  while (current <= rangeEndMonth) {
    months.push(current);
    if (current === rangeEndMonth) break;
    current = incrementMonth(current);
  }
  const maxRangeMonths = Math.max(1, Math.min(RESOURCE_RANGE_MONTHS_MAX, Number(env.RESOURCE_RANGE_MONTHS || RESOURCE_RANGE_MONTHS_DEFAULT)));
  if (months.length > maxRangeMonths) {
    return json({ ok: false, error: "RANGE_TOO_LARGE", maxRangeMonths, requested: months.length }, 422, request);
  }

  // Safety: never re-fetch months that are already completed (idempotent).
  const monthStarts = months.map(m => configuredMonthRange(m, start, end).start);
  const placeholders = months.map(() => "?").join(",");
  const existing = await env.DB.prepare(
    `SELECT chunk_start FROM resource_chunk_summaries WHERE request_id=? AND chunk_start IN (${placeholders})`
  ).bind(requestId, ...monthStarts).all();
  if ((existing.results || []).length) {
    return json({ ok: false, error: "RANGE_OVERLAPS_COMPLETED_MONTHS" }, 409, request);
  }

  const range = {
    start: configuredMonthRange(months[0], start, end).start,
    end: configuredMonthRange(months[months.length - 1], start, end).end
  };

  const processStartMs = performance.now();
  let validated;
  try {
    validated = await fetchValidatedResourceRange(row, range);
  } catch (error) {
    const detail = String(error?.message || error);
    const unavailable = detail.startsWith("NASA_NETWORK_ERROR") || detail.startsWith("NASA_HTTP_") || detail === "NASA_INVALID_JSON" || detail.startsWith("NASA_422_") || detail.startsWith("NASA_PAYLOAD_");
    await markResourceError(env, requestId, unavailable ? "DATA UNAVAILABLE" : "ERROR", detail);
    return json({
      ok: false, requestId, rangeStart: rangeStartMonth, rangeEnd: rangeEndMonth,
      status: unavailable ? "DATA UNAVAILABLE" : "ERROR", error: detail, retryable: unavailable,
      processMs: Math.round(performance.now() - processStartMs)
    }, unavailable ? 503 : 422, request);
  }
  const processMs = Math.round(performance.now() - processStartMs);

  const now = Date.now();
  const statements = [];
  for (const part of validated.parts) {
    const partRange = configuredMonthRange(part.month, start, end);
    const chunkJson = JSON.stringify({
      schema: "SOLAR_EPC_NASA_HOURLY_V1",
      timeStandard: RESOURCE_TIME_STANDARD,
      parameterOrder: RESOURCE_PARAMETERS,
      rows: part.rows
    });
    const checksum = await sha256Hex(chunkJson);
    const monthState = mergeMonthState(newResourceSummaryState(), part);
    const partMetadata = {
      ...validated.metadata,
      firstChunkStart: partRange.start,
      latestChunkEnd: partRange.end
    };
    statements.push(
      env.DB.prepare(
        `INSERT INTO resource_chunks(request_id,chunk_start,chunk_end,time_standard,parameter_order_json,data_json,
          record_count,missing_count,invalid_count,duplicate_count,checksum_sha256,created_at)
         VALUES(?,?,?,?,?,?,?,?,?,?,?,?)
         ON CONFLICT(request_id,chunk_start) DO UPDATE SET
          chunk_end=excluded.chunk_end,time_standard=excluded.time_standard,
          parameter_order_json=excluded.parameter_order_json,data_json=excluded.data_json,
          record_count=excluded.record_count,missing_count=excluded.missing_count,
          invalid_count=excluded.invalid_count,duplicate_count=excluded.duplicate_count,
          checksum_sha256=excluded.checksum_sha256,created_at=excluded.created_at`
      ).bind(
        requestId, partRange.start, partRange.end, RESOURCE_TIME_STANDARD,
        JSON.stringify(RESOURCE_PARAMETERS), chunkJson, part.rows.length,
        part.missingCount, part.invalidCount, part.duplicateCount,
        checksum, now
      ),
      env.DB.prepare(
        `INSERT INTO resource_chunk_summaries(request_id,chunk_start,summary_json,units_json,metadata_json,created_at)
         VALUES(?,?,?,?,?,?)
         ON CONFLICT(request_id,chunk_start) DO UPDATE SET
          summary_json=excluded.summary_json,units_json=excluded.units_json,
          metadata_json=excluded.metadata_json,created_at=excluded.created_at`
      ).bind(requestId, partRange.start, JSON.stringify(monthState), JSON.stringify(validated.units), JSON.stringify(partMetadata), now)
    );
  }
  statements.push(env.DB.prepare(
    `UPDATE resource_requests SET status=?,error_text=NULL,
     units_json=COALESCE(units_json,?),metadata_json=COALESCE(metadata_json,?),updated_at=?
     WHERE request_id=? AND workbook_id=?`
  ).bind(
    validated.anyReview ? "REVIEW" : "PROCESSING",
    JSON.stringify(validated.units), JSON.stringify(validated.metadata), now, requestId, workbook
  ));

  // K=12 -> 24 monthly writes + 1 request update in one atomic batch (25 D1
  // statements, well below the Workers Free 50-query limit per invocation).
  await env.DB.batch(statements);

  return json({
    ok: true, requestId, processed: true, rangeStart: rangeStartMonth, rangeEnd: rangeEndMonth,
    processedMonths: validated.parts.length, months: validated.parts.map(p => p.month).join(","),
    records: validated.totalRecords,
    processMs, rangeMonths: maxRangeMonths,
    validation: {
      missing: validated.totalMissing, invalid: validated.totalInvalid,
      duplicates: validated.totalDuplicate, review: validated.anyReview
    }
  }, 200, request);
}

async function fetchValidatedResourceRange(row, range) {
  const nasaUrl = new URL("https://power.larc.nasa.gov/api/temporal/hourly/point");
  nasaUrl.searchParams.set("parameters", RESOURCE_PARAMETERS.join(","));
  nasaUrl.searchParams.set("community", "RE");
  nasaUrl.searchParams.set("longitude", fixed8(Number(row.centroid_lon)));
  nasaUrl.searchParams.set("latitude", fixed8(Number(row.centroid_lat)));
  nasaUrl.searchParams.set("start", range.start);
  nasaUrl.searchParams.set("end", range.end);
  nasaUrl.searchParams.set("format", "JSON");
  nasaUrl.searchParams.set("time-standard", RESOURCE_TIME_STANDARD);

  let response;
  try {
    response = await fetch(nasaUrl.toString(), { headers: { "Accept": "application/json", "User-Agent": "Solar-EPC-Resource/1.0" } });
  } catch (error) {
    throw new Error("NASA_NETWORK_ERROR: " + String(error?.message || error));
  }
  if (response.status === 422) {
    // NASA rejects a JSON request whose response is too large for the extent.
    // The VBA falls back to month-by-month processing for this range.
    throw new Error("NASA_422_EXTENT_TOO_LARGE");
  }
  if (!response.ok) {
    const text = (await response.text()).slice(0, 1000);
    throw new Error(`NASA_HTTP_${response.status}: ${text}`);
  }
  let text;
  try { text = await response.text(); }
  catch { throw new Error("NASA_INVALID_JSON"); }
  if (text.length > RESOURCE_NASA_RESPONSE_BYTES_MAX) {
    // Safety guard: a very large multi-month body would risk the Free CPU
    // budget. Signal VBA to fall back to month-by-month for this range.
    throw new Error(`NASA_PAYLOAD_TOO_LARGE:${text.length}bytes`);
  }
  let nasa;
  try { nasa = JSON.parse(text); }
  catch { throw new Error("NASA_INVALID_JSON"); }
  try {
    return validateAndNormalizeNasaRange(nasa, range, Number(row.centroid_lat), Number(row.centroid_lon));
  } catch (error) {
    throw new Error("NASA_VALIDATION_ERROR: " + String(error?.message || error));
  }
}

// Validates one multi-month NASA response and returns per-month parts whose
// rows, states and counts are computed exactly like the month-by-month path.
function validateAndNormalizeNasaRange(nasa, range, requestedLat, requestedLon) {
  const header = nasa?.header;
  const parametersMeta = nasa?.parameters;
  const data = nasa?.properties?.parameter;
  if (!header || !parametersMeta || !data) throw new Error("Required NASA response sections are missing.");
  if (String(header?.time_standard || "").toUpperCase() !== "UTC") throw new Error("NASA returned an unexpected time standard.");
  if (String(header?.start || "") !== range.start || String(header?.end || "") !== range.end) throw new Error("NASA returned a period different from the requested period.");
  if (!header?.api?.version || !header?.api?.name || !Array.isArray(header?.sources) || !header.sources.length) throw new Error("NASA source metadata is incomplete.");
  const returnedCoordinate = nasa?.geometry?.coordinates;
  if (!Array.isArray(returnedCoordinate) || returnedCoordinate.length < 2 ||
      Math.abs(Number(returnedCoordinate[0]) - requestedLon) > 0.01 ||
      Math.abs(Number(returnedCoordinate[1]) - requestedLat) > 0.01) {
    throw new Error("NASA returned coordinates inconsistent with the requested centroid.");
  }

  const units = {};
  let unitMismatchCount = 0;
  for (const p of RESOURCE_PARAMETERS) {
    if (!Object.prototype.hasOwnProperty.call(data, p)) throw new Error(`NASA parameter missing: ${p}`);
    if (!Object.prototype.hasOwnProperty.call(parametersMeta, p)) throw new Error(`NASA parameter metadata missing: ${p}`);
    units[p] = String(parametersMeta[p]?.units || "");
    if (!units[p]) throw new Error(`NASA unit missing: ${p}`);
    if (RESOURCE_EXPECTED_UNITS[p] && units[p] !== RESOURCE_EXPECTED_UNITS[p]) unitMismatchCount++;
  }
  const returnedNames = Object.keys(data).sort();
  const expectedNames = [...RESOURCE_PARAMETERS].sort();
  if (returnedNames.join("|") !== expectedNames.join("|")) throw new Error("NASA returned a different parameter set.");

  const expectedTimes = expectedHourlyKeys(range.start, range.end);
  const expectedSet = new Set(expectedTimes);

  // Per-month duplicate counters (identical counting rule to the month path).
  const monthDuplicateCount = {};
  for (const p of RESOURCE_PARAMETERS) {
    const keys = Object.keys(data[p]);
    const seen = new Set();
    for (const key of keys) {
      if (seen.has(key)) {
        const m = key.slice(0, 6);
        monthDuplicateCount[m] = (monthDuplicateCount[m] || 0) + 1;
      }
      seen.add(key);
      if (!expectedSet.has(key)) throw new Error(`NASA returned timestamp outside requested range: ${key}`);
      if (!/^\d{10}$/.test(key)) throw new Error(`NASA returned invalid timestamp: ${key}`);
    }
  }

  // Per-month buckets, populated in the exact same order as the month path.
  const buckets = {};
  const monthKeys = [];
  for (const timestamp of expectedTimes) {
    const m = timestamp.slice(0, 6);
    if (!buckets[m]) {
      buckets[m] = {
        rows: [],
        state: newResourceSummaryState(),
        daily: { ALLSKY_SFC_SW_DWN: {}, ALLSKY_SFC_SW_DNI: {}, ALLSKY_SFC_SW_DIFF: {} },
        missing: 0, invalid: 0
      };
      monthKeys.push(m);
    }
  }

  for (const timestamp of expectedTimes) {
    const bucket = buckets[timestamp.slice(0, 6)];
    const row = [timestamp];
    for (const p of RESOURCE_PARAMETERS) {
      const obj = data[p];
      let value = obj[timestamp];
      const fill = Number(header.fill_value);
      if (value === undefined || value === null || Number(value) === fill || !Number.isFinite(Number(value))) {
        value = null;
        bucket.missing++;
        bucket.state.metrics[p].missing++;
      } else {
        value = Number(value);
        if (!resourceValueValid(p, value)) {
          value = null;
          bucket.invalid++;
          bucket.state.metrics[p].invalid++;
        } else {
          addMetric(bucket.state.metrics[p], value);
          if (bucket.daily[p]) {
            const day = timestamp.slice(0, 8);
            if (!bucket.daily[p][day]) bucket.daily[p][day] = { sum: 0, count: 0 };
            bucket.daily[p][day].sum += value;
            bucket.daily[p][day].count++;
          }
        }
      }
      row.push(value);
    }
    bucket.rows.push(row);
  }

  const parts = [];
  let totalRecords = 0, totalMissing = 0, totalInvalid = 0, totalDuplicate = 0, anyReview = false;
  for (const month of monthKeys) {
    const bucket = buckets[month];
    for (const p of Object.keys(bucket.daily)) {
      for (const day of Object.keys(bucket.daily[p])) {
        const d = bucket.daily[p][day];
        if (d.count === 24) {
          bucket.state.dailyIrradiance[p].sum += d.sum / 1000;
          bucket.state.dailyIrradiance[p].count++;
        } else {
          bucket.state.dailyIrradiance[p].incomplete++;
        }
      }
    }
    const duplicateCount = monthDuplicateCount[month] || 0;
    const reviewCount = bucket.missing + bucket.invalid + duplicateCount + unitMismatchCount;
    bucket.state.missingCount = bucket.missing;
    bucket.state.invalidCount = bucket.invalid;
    bucket.state.duplicateCount = duplicateCount;
    bucket.state.unitMismatchCount = unitMismatchCount;
    bucket.state.reviewCount = reviewCount;
    parts.push({
      month,
      rows: bucket.rows,
      state: bucket.state,
      missingCount: bucket.missing,
      invalidCount: bucket.invalid,
      duplicateCount,
      unitMismatchCount,
      reviewCount
    });
    totalRecords += bucket.rows.length;
    totalMissing += bucket.missing;
    totalInvalid += bucket.invalid;
    totalDuplicate += duplicateCount;
    if (reviewCount > 0) anyReview = true;
  }

  return {
    parts, units,
    metadata: {
      title: String(header.title || ""), apiName: String(header.api.name),
      apiVersion: String(header.api.version), sources: header.sources,
      fillValue: header.fill_value, timeStandard: header.time_standard,
      firstChunkStart: range.start, latestChunkEnd: range.end,
      parameterMetadata: parametersMeta,
      returnedGeometry: nasa.geometry,
      messages: Array.isArray(nasa.messages) ? nasa.messages : []
    },
    totalRecords, totalMissing, totalInvalid, totalDuplicate, anyReview
  };
}

async function resourceFinalize(request, env, workbook) {
  let body;
  try { body = await request.json(); }
  catch { return json({ ok: false, error: "INVALID_JSON" }, 400, request); }
  const requestId = clean(body?.requestId, 100);
  if (!requestId) return json({ ok: false, error: "MISSING_REQUEST_ID" }, 400, request);

  const row = await env.DB.prepare(
    "SELECT * FROM resource_requests WHERE request_id=? AND workbook_id=?"
  ).bind(requestId, workbook).first();
  if (!row) return json({ ok: false, error: "RESOURCE_REQUEST_NOT_FOUND" }, 404, request);
  if (row.status === "COMPLETE" && row.summary_json) return json({ ok: true, requestId, complete: true, cached: true }, 200, request);

  const chunks = await env.DB.prepare(
    "SELECT summary_json,units_json,metadata_json FROM resource_chunk_summaries WHERE request_id=? ORDER BY chunk_start"
  ).bind(requestId).all();
  const results = chunks.results || [];
  if (results.length !== Number(row.total_chunks)) {
    return json({ ok: false, error: "RESOURCE_CHUNKS_INCOMPLETE", completedChunks: results.length, totalChunks: row.total_chunks }, 409, request);
  }

  const state = newResourceSummaryState();
  for (const chunk of results) mergeResourceSummaryStates(state, safeJson(chunk.summary_json, newResourceSummaryState()));
  const units = safeJson(row.units_json, safeJson(results[0]?.units_json, {}));
  const metadata = safeJson(row.metadata_json, safeJson(results[0]?.metadata_json, {}));
  const finalRow = { ...row, completed_chunks: row.total_chunks, status: "COMPLETE" };
  const summary = finalizeResourceSummary(finalRow, state, units, metadata);
  const now = Date.now();
  await env.DB.prepare(
    `UPDATE resource_requests SET completed_chunks=total_chunks,status='COMPLETE',error_text=NULL,
     summary_state_json=?,summary_json=?,updated_at=? WHERE request_id=? AND workbook_id=?`
  ).bind(JSON.stringify(state), JSON.stringify(summary), now, requestId, workbook).run();
  return json({ ok: true, requestId, complete: true, dataStatus: summary.status, completedChunks: row.total_chunks, totalChunks: row.total_chunks }, 200, request);
}

function mergeResourceSummaryStates(total, part) {
  for (const p of RESOURCE_PARAMETERS) {
    const a = total.metrics[p], b = part.metrics?.[p];
    if (!b) continue;
    a.sum += Number(b.sum || 0); a.count += Number(b.count || 0);
    a.missing += Number(b.missing || 0); a.invalid += Number(b.invalid || 0);
    if (b.min !== null && b.min !== undefined) a.min = a.min === null ? Number(b.min) : Math.min(a.min, Number(b.min));
    if (b.max !== null && b.max !== undefined) a.max = a.max === null ? Number(b.max) : Math.max(a.max, Number(b.max));
  }
  for (const p of Object.keys(total.dailyIrradiance)) {
    const b = part.dailyIrradiance?.[p]; if (!b) continue;
    total.dailyIrradiance[p].sum += Number(b.sum || 0);
    total.dailyIrradiance[p].count += Number(b.count || 0);
    total.dailyIrradiance[p].incomplete += Number(b.incomplete || 0);
  }
  total.windSin += Number(part.windSin || 0);
  total.windCos += Number(part.windCos || 0);
  total.windDirectionCount += Number(part.windDirectionCount || 0);
  total.missingCount += Number(part.missingCount || 0);
  total.invalidCount += Number(part.invalidCount || 0);
  total.duplicateCount += Number(part.duplicateCount || 0);
  total.unitMismatchCount += Number(part.unitMismatchCount || 0);
  total.reviewCount += Number(part.reviewCount || 0);
  return total;
}

function configuredMonths(startCompact, endCompact) {
  const months = [];
  let current = startCompact.slice(0, 6);
  const last = endCompact.slice(0, 6);
  while (current <= last) {
    months.push(current);
    if (current === last) break;
    current = incrementMonth(current);
  }
  return months;
}
