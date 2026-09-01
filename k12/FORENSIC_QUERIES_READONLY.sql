-- ============================================================================
-- READ-ONLY FORENSIC QUERIES — Solar EPC NASA POWER resource dataset
-- PART 1: CHECK #2  (ALLSKY_SRF_ALB fill vs ALLSKY_SFC_SW_DWN)
-- PART 2: CENTROID IDENTITY (request 1dce0c1c... vs 18.18043580,74.61004963)
-- All statements are SELECT only. No INSERT/UPDATE/DELETE. No status changes.
-- Replace <DB_NAME>, <WORKBOOK_ID> and the three request IDs as appropriate.
-- Works in wrangler:  npx wrangler d1 execute <DB_NAME> --remote --command "..."
-- ============================================================================

-- ============================================================================
-- PART 1 — CHECK #2: classify every ALB fill by same-timestamp DWN
-- row layout inside data_json:
--   rows[i] = [ts, DWN, DNI, DIFF, T2M, WS10M, WD10M, RH2M, PS, PRECTOTCORR, ALB]
--   -> ALB index 10, DWN index 1 (derived from parameterOrder in the export too)
-- ============================================================================

-- 1.0 Sanity: chunk inventory for the request
-- SELECT COUNT(*) AS chunks,
--        SUM(record_count)      AS records_stored,
--        SUM(missing_count)     AS missing_stored,
--        SUM(invalid_count)     AS invalid_stored,
--        SUM(duplicate_count)   AS duplicates_stored
-- FROM resource_chunks
-- WHERE request_id='1dce0c1c-0306-4765-97f8-567a49bed34b';

-- 1.1 THE classification query (exactly one bucket per ALB fill)
WITH rec AS (
  SELECT je.value AS row
  FROM resource_chunks rc,
       json_each(json_extract(rc.data_json,'$.rows')) AS je
  WHERE rc.request_id = '1dce0c1c-0306-4765-97f8-567a49bed34b'
)
SELECT
  SUM(CASE WHEN json_extract(row,'$[10]') IS NULL THEN 1 ELSE 0 END) AS alb_missing_total,
  SUM(CASE WHEN json_extract(row,'$[10]') IS NULL
            AND json_extract(row,'$[1]') = 0 THEN 1 ELSE 0 END)     AS night_na_candidate,
  SUM(CASE WHEN json_extract(row,'$[10]') IS NULL
            AND json_extract(row,'$[1]') > 0 THEN 1 ELSE 0 END)     AS daytime_albedo_missing,
  SUM(CASE WHEN json_extract(row,'$[10]') IS NULL
            AND json_extract(row,'$[1]') IS NULL THEN 1 ELSE 0 END) AS dwn_unavailable
FROM rec;

-- 1.2 Optional guard: ALB fill where DWN looks invalid/negative (should be 0;
--     stored DWN is already normalized null for NASA fill, so this catches anomalies)
-- SELECT json_extract(je.value,'$[0]') AS ts, json_extract(je.value,'$[1]') AS dwn
-- FROM resource_chunks rc, json_each(json_extract(rc.data_json,'$.rows')) je
-- WHERE rc.request_id='1dce0c1c-0306-4765-97f8-567a49bed34b'
--   AND json_extract(je.value,'$[10]') IS NULL
--   AND json_extract(je.value,'$[1]') IS NOT NULL
--   AND json_extract(je.value,'$[1]') <= 0
-- ORDER BY ts;

-- 1.3 ALL DAYTIME_ALBEDO_MISSING timestamps (diagnostic list)
-- SELECT json_extract(je.value,'$[0]') AS ts, json_extract(je.value,'$[1]') AS dwn
-- FROM resource_chunks rc, json_each(json_extract(rc.data_json,'$.rows')) je
-- WHERE rc.request_id='1dce0c1c-0306-4765-97f8-567a49bed34b'
--   AND json_extract(je.value,'$[10]') IS NULL
--   AND json_extract(je.value,'$[1]') > 0
-- ORDER BY ts;

-- 1.4 EXPORT for the classifier (parses rows exactly, builds run lengths)
-- npx wrangler d1 execute <DB_NAME> --remote \
--   --command "SELECT request_id,chunk_start,parameter_order_json,data_json \
--              FROM resource_chunks \
--              WHERE request_id='1dce0c1c-0306-4765-97f8-567a49bed34b' \
--              ORDER BY chunk_start" --json > check2_chunks.json
-- node k12/forensic_check2.mjs check2_chunks.json 1dce0c1c-0306-4765-97f8-567a49bed34b

-- ============================================================================
-- PART 2 — CENTROID IDENTITY
-- ============================================================================

-- 2.1 Find every request for each of the three centroids (also gives the
--     request_id that corresponds to 18.18043580, 74.61004963)
SELECT request_id, workbook_id, project_id, status,
       centroid_lat, centroid_lon, period_start, period_end,
       cache_key, polygon_hash, completed_chunks, total_chunks,
       created_at, updated_at
FROM resource_requests
WHERE ABS(centroid_lat - 17.99038747069333) < 0.00000001
   OR ABS(centroid_lat - 17.990776332796173) < 0.00000001
   OR ABS(centroid_lat - 18.18043580) < 0.00000001
ORDER BY centroid_lat, created_at;

-- 2.2 Cache-key comparison across the three requests (expected: all DIFFERENT)
-- SELECT request_id, centroid_lat, centroid_lon, cache_key, polygon_hash
-- FROM resource_requests
-- WHERE request_id IN ('1dce0c1c-0306-4765-97f8-567a49bed34b',
--                      '<REQ_B>', '<REQ_C_18.18043580>')
-- ORDER BY centroid_lat;

-- 2.3 Chunk/row inventory per request (expected: 60 chunks & 43,824 rows each
--     for a 60-month request; range path: 12 per request / 4,381 per range)
-- SELECT request_id, COUNT(*) AS chunks,
--        SUM(record_count) AS rows, SUM(missing_count) AS missing_total,
--        SUM(invalid_count) AS invalid_total, SUM(duplicate_count) AS dup_total,
--        COUNT(DISTINCT checksum_sha256) AS distinct_checksums
-- FROM resource_chunks
-- WHERE request_id IN ('1dce0c1c-0306-4765-97f8-567a49bed34b', '<REQ_C>')
-- GROUP BY request_id;

-- 2.4 Are chunk payload checksums literally shared between requests?
--     (If the same NASA cell was answered, hashes are expected to be equal —
--      this is NOT a cache leak; the app computes checksums from its own fetch.)
-- SELECT request_id, chunk_start, checksum_sha256
-- FROM resource_chunks
-- WHERE request_id IN ('1dce0c1c-0306-4765-97f8-567a49bed34b', '<REQ_C>')
-- ORDER BY chunk_start;

-- 2.5 Summary isolation / identity (each request must own exactly one row)
-- SELECT request_id,
--        length(summary_state_json) AS state_len,
--        length(summary_json)       AS summary_len,
--        json_extract(summary_json,'$.data_status') AS status_field
-- FROM resource_requests
-- WHERE request_id IN ('1dce0c1c-0306-4765-97f8-567a49bed34b', '<REQ_C>');

-- 2.6 Raw hourly equality between request 1 and request 3 (T2M etc.) —
--     NULL-safe comparison at every matching timestamp:
-- WITH a AS (
--   SELECT json_extract(je.value,'$[0]') AS ts,
--          json_extract(je.value,'$[4]') AS t2m,
--          json_extract(je.value,'$[5]') AS ws,
--          json_extract(je.value,'$[7]') AS rh,
--          json_extract(je.value,'$[8]') AS ps,
--          json_extract(je.value,'$[9]') AS prec
--   FROM resource_chunks rc, json_each(json_extract(rc.data_json,'$.rows')) je
--   WHERE rc.request_id='1dce0c1c-0306-4765-97f8-567a49bed34b'
-- ), b AS (
--   SELECT json_extract(je.value,'$[0]') AS ts,
--          json_extract(je.value,'$[4]') AS t2m,
--          json_extract(je.value,'$[5]') AS ws,
--          json_extract(je.value,'$[7]') AS rh,
--          json_extract(je.value,'$[8]') AS ps,
--          json_extract(je.value,'$[9]') AS prec
--   FROM resource_chunks rc, json_each(json_extract(rc.data_json,'$.rows')) je
--   WHERE rc.request_id='<REQ_C>'
-- )
-- SELECT COUNT(*) AS matched,
--        SUM(CASE WHEN a.t2m IS NOT b.t2m THEN 1 ELSE 0 END) AS t2m_diff,
--        SUM(CASE WHEN a.ws  IS NOT b.ws  THEN 1 ELSE 0 END) AS ws_diff,
--        SUM(CASE WHEN a.rh  IS NOT b.rh  THEN 1 ELSE 0 END) AS rh_diff,
--        SUM(CASE WHEN a.ps  IS NOT b.ps  THEN 1 ELSE 0 END) AS ps_diff,
--        SUM(CASE WHEN a.prec IS NOT b.prec THEN 1 ELSE 0 END) AS prec_diff
-- FROM a JOIN b USING (ts);

-- 2.7 EXPORT both requests' chunks for the row-level comparator:
-- npx wrangler d1 execute <DB_NAME> --remote \
--   --command "SELECT request_id,chunk_start,parameter_order_json,data_json,checksum_sha256 \
--              FROM resource_chunks WHERE request_id='1dce0c1c-0306-4765-97f8-567a49bed34b' \
--              ORDER BY chunk_start" --json > chunks_req1.json
-- npx wrangler d1 execute <DB_NAME> --remote \
--   --command "SELECT request_id,chunk_start,parameter_order_json,data_json,checksum_sha256 \
--              FROM resource_chunks WHERE request_id='<REQ_C>' \
--              ORDER BY chunk_start" --json > chunks_reqC.json
-- npx wrangler d1 execute <DB_NAME> --remote \
--   --command "SELECT request_id,workbook_id,project_id,status,centroid_lat,centroid_lon, \
--              period_start,period_end,cache_key,polygon_hash,completed_chunks,total_chunks \
--              FROM resource_requests ORDER BY created_at" --json > requests_meta.json
-- node k12/forensic_centroid_chunks.mjs chunks_req1.json chunks_reqC.json requests_meta.json
