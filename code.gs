/**
 * Technical Productivity Assessment — Sheets ➜ Supabase sync
 * ---------------------------------------------------------
 * Pushes three source tabs into the _tpa staging tables in the
 * `weekly-report` Supabase project, then calls the merge functions.
 *
 * SETUP (once):
 *   1. Extensions ➜ Apps Script, paste this file in.
 *   2. Project Settings ➜ Script Properties ➜ add:
 *        SUPABASE_SERVICE_KEY = <the service_role key>
 *      Never paste the key into this file.
 *   3. Run `syncAll` once and approve the authorisation prompt.
 *   4. Run `installTriggers` to schedule it hourly.
 *
 * ENTRY POINTS:
 *   syncAll()          – normal run, last SYNC_WINDOW_DAYS of data
 *   runFullBackfill()  – ignores the window, loads all history
 *   installTriggers()  – hourly schedule
 */

// ───────────────────────────── config ─────────────────────────────

var SUPABASE_URL     = 'https://imvvwqnmidyrfzncwzjn.supabase.co';
var SPREADSHEET_ID   = '14tKu1JmTcvuDTgb7RxLm-yB_SyIVDDILbR9Jg3wd70M';
var SYNC_WINDOW_DAYS = 60;    // rows older than this are skipped on a normal run
var BATCH_SIZE       = 500;   // rows per HTTP request

/**
 * Each source tab: where its header row is, which staging table it
 * feeds, which merge function drains it, and how its columns map.
 *
 * `from` is matched against the sheet header after normalising
 * (lowercase, punctuation stripped, whitespace collapsed), so minor
 * header edits upstream will not break the sync.
 */
var SOURCES = [
  {
    key:        'production',
    tab:        'CICO DB',
    headerRow:  2,
    staging:    'production_job_card_staging_tpa',
    mergeFn:    'merge_production_tpa',
    dateColumn: 'clock_in_at',
    columns: [
      { to: 'id',               from: 'id',               type: 'int'  },
      { to: 'repair_item_id',   from: 'repair_item_id',   type: 'int'  },
      { to: 'repair_item_name', from: 'repair_item_name', type: 'text' },
      { to: 'item_name',        from: 'item_name',        type: 'text' },
      { to: 'car_id',           from: 'car_id',           type: 'int'  },
      { to: 'clock_in_at',      from: 'clock_in_at',      type: 'ts'   },
      { to: 'clock_out_at',     from: 'clock_out_at',     type: 'ts'   },
      { to: 'clock_in_email',   from: 'clock_in_email',   type: 'text' },
      { to: 'clock_out_email',  from: 'clock_out_email',  type: 'text' },
      { to: 'status',           from: 'status',           type: 'text' },
      { to: 'stage',            from: 'stage',            type: 'text' },
      { to: 'estimated_hours',  from: 'estimated_hours',  type: 'num'  },
      { to: 'repeated_job',     from: 'repeated_job',     type: 'text' },
      { to: 'invoiced_at',      from: 'invoiced_at',      type: 'date' }
    ]
  },
  {
    key:        'aftersales',
    tab:        'A/S Clockin-out',
    headerRow:  2,
    staging:    'aftersales_job_card_staging_tpa',
    mergeFn:    'merge_aftersales_tpa',
    dateColumn: 'clock_in_at',
    columns: [
      { to: 'id',                    from: 'id',                    type: 'int'  },
      { to: 'correct_item_id',       from: 'correct_item_id',       type: 'int'  },
      { to: 'inquiry_id',            from: 'inquiry_id',            type: 'int'  },
      { to: 'car_id',                from: 'car_id',                type: 'int'  },
      { to: 'vin_no',                from: 'vin_no',                type: 'text' },
      { to: 'item_name',             from: 'item_name',             type: 'text' },
      // The source calls the clock-in timestamp `created_at`.
      { to: 'clock_in_at',           from: 'created_at',            type: 'ts'   },
      { to: 'clock_out_at',          from: 'clock_out_at',          type: 'ts'   },
      { to: 'clocked_in_by',         from: 'clocked_in_by',         type: 'text' },
      { to: 'clocked_out_by',        from: 'clocked_out_by',        type: 'text' },
      { to: 'tech',                  from: 'tech',                  type: 'text' },
      { to: 'type',                  from: 'type',                  type: 'text' },
      { to: 'quality_wall_approval', from: 'quality_wall_approval', type: 'text' },
      { to: 'status',                from: 'status',                type: 'text' },
      { to: 'estimated_hours',       from: 'estimated_hours',       type: 'num'  }
    ]
  },
  {
    key:        'inspection',
    tab:        'post insp',
    headerRow:  1,
    staging:    'inspection_staging_tpa',
    mergeFn:    'merge_inspection_tpa',
    dateColumn: 'inspected_at',
    columns: [
      { to: 'car_id',             from: 'car_id',                              type: 'int'  },
      { to: 'inspected_at',       from: 'date',                                type: 'ts'   },
      { to: 'inspection_type',    from: 'inspection_type',                     type: 'text' },
      { to: 'tech_name',          from: 'name',                                type: 'text' },
      { to: 'new_findings',       from: 'no of new findings in post inspection', type: 'int' },
      { to: 'post_not_necessary', from: 'post insp not_necessary',             type: 'int'  },
      { to: 'pre_not_necessary',  from: 'pre insp not_necessary',              type: 'int'  },
      { to: 'rework',             from: 'rework',                              type: 'int'  },
      { to: 'days_to_return',     from: 'after how many days it came back',    type: 'int'  }
    ]
  }
];

// ─────────────────────────── entry points ───────────────────────────

function syncAll()         { return runSync_(SYNC_WINDOW_DAYS); }
function runFullBackfill() { return runSync_(null); }

function installTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'syncAll') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('syncAll').timeBased().everyHours(1).create();
  Logger.log('Hourly syncAll trigger installed.');
}

// ──────────────────────────── the sync ────────────────────────────

function runSync_(windowDays) {
  var ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
  var tz      = ss.getSpreadsheetTimeZone();
  var cutoff  = windowDays === null
              ? null
              : new Date(Date.now() - windowDays * 86400000);
  var summary = [];

  SOURCES.forEach(function (src) {
    var started = new Date();
    try {
      var rows     = readTab_(ss, src, tz, cutoff);
      var upserted = pushAndMerge_(src, rows);
      summary.push(src.key + ': ' + rows.length + ' read, ' + upserted + ' upserted');
      logSync_(src.tab, started, rows.length, upserted, 'ok', null);
    } catch (err) {
      summary.push(src.key + ': FAILED — ' + err.message);
      logSync_(src.tab, started, null, null, 'error', String(err.message).slice(0, 900));
    }
  });

  // Create employee + alias rows for anyone new in the data.
  try {
    var resolved = rpc_('resolve_employees_tpa', {});
    var r = (resolved && resolved[0]) || {};
    summary.push('employees: +' + (r.created_employees || 0) +
                 ' people, +' + (r.created_aliases || 0) + ' aliases');
  } catch (err) {
    summary.push('employees: FAILED — ' + err.message);
  }

  Logger.log(summary.join('\n'));
  return summary.join('\n');
}

/** Reads one tab and returns an array of objects ready for its staging table. */
function readTab_(ss, src, tz, cutoff) {
  var sheet = ss.getSheetByName(src.tab);
  if (!sheet) throw new Error('Tab not found: ' + src.tab);

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow <= src.headerRow) return [];

  var header = sheet.getRange(src.headerRow, 1, 1, lastCol).getValues()[0].map(normalise_);
  var index  = {};
  src.columns.forEach(function (c) {
    var i = findColumn_(header, normalise_(c.from));
    if (i === -1) throw new Error('Column "' + c.from + '" not found on tab ' + src.tab);
    index[c.to] = i;
  });

  var firstDataRow = src.headerRow + 1;
  var values = sheet.getRange(firstDataRow, 1, lastRow - src.headerRow, lastCol).getValues();
  var out    = [];

  for (var r = 0; r < values.length; r++) {
    var row = values[r];
    var obj = {};
    var keep = false;

    for (var c = 0; c < src.columns.length; c++) {
      var col = src.columns[c];
      var val = coerce_(row[index[col.to]], col.type, tz);
      obj[col.to] = val;
      if (val !== null && val !== '') keep = true;
    }
    if (!keep) continue;

    if (cutoff) {
      var stamp = obj[src.dateColumn];
      if (!stamp) continue;
      if (new Date(String(stamp).replace(' ', 'T')) < cutoff) continue;
    }
    out.push(obj);
  }
  return out;
}

/** Uploads rows to staging in batches, then runs the merge function. */
function pushAndMerge_(src, rows) {
  if (!rows.length) return 0;
  for (var i = 0; i < rows.length; i += BATCH_SIZE) {
    request_('POST', '/rest/v1/' + src.staging, rows.slice(i, i + BATCH_SIZE), 'return=minimal');
  }
  var result = rpc_(src.mergeFn, {});
  return typeof result === 'number' ? result : rows.length;
}

function logSync_(tab, started, read, upserted, status, message) {
  try {
    request_('POST', '/rest/v1/sync_log_tpa', [{
      source_tab:    tab,
      started_at:    started.toISOString(),
      finished_at:   new Date().toISOString(),
      rows_read:     read,
      rows_upserted: upserted,
      status:        status,
      message:       message
    }], 'return=minimal');
  } catch (e) {
    Logger.log('sync_log write failed: ' + e.message);
  }
}

// ───────────────────────────── helpers ─────────────────────────────

function serviceKey_() {
  var k = PropertiesService.getScriptProperties().getProperty('SUPABASE_SERVICE_KEY');
  if (!k) throw new Error('SUPABASE_SERVICE_KEY is not set in Script Properties.');
  return k;
}

function request_(method, path, payload, prefer) {
  var key = serviceKey_();
  var res = UrlFetchApp.fetch(SUPABASE_URL + path, {
    method:            method.toLowerCase(),
    contentType:       'application/json',
    headers:           { apikey: key, Authorization: 'Bearer ' + key, Prefer: prefer || 'return=representation' },
    payload:           JSON.stringify(payload),
    muteHttpExceptions: true
  });
  var code = res.getResponseCode();
  if (code >= 300) throw new Error(method + ' ' + path + ' → ' + code + ' ' + res.getContentText().slice(0, 400));
  var body = res.getContentText();
  return body ? JSON.parse(body) : null;
}

function rpc_(fn, args) {
  return request_('POST', '/rest/v1/rpc/' + fn, args, 'return=representation');
}

/** Lowercase, strip punctuation, collapse whitespace. */
function normalise_(v) {
  return String(v == null ? '' : v)
    .toLowerCase()
    .replace(/[^a-z0-9_ ]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

/** Exact match first, then a prefix match so truncated headers still resolve. */
function findColumn_(header, target) {
  var i = header.indexOf(target);
  if (i !== -1) return i;
  for (var j = 0; j < header.length; j++) {
    if (header[j] && (header[j].indexOf(target) === 0 || target.indexOf(header[j]) === 0)) return j;
  }
  return -1;
}

function coerce_(value, type, tz) {
  if (value === '' || value === null || value === undefined) return null;

  if (type === 'ts' || type === 'date') {
    var d = (value instanceof Date) ? value : new Date(value);
    if (isNaN(d.getTime())) return null;
    return Utilities.formatDate(d, tz, type === 'ts' ? 'yyyy-MM-dd HH:mm:ss' : 'yyyy-MM-dd');
  }

  if (type === 'int' || type === 'num') {
    var n = (typeof value === 'number') ? value : Number(String(value).replace(/[^0-9.\-]/g, ''));
    if (!isFinite(n)) return null;
    return type === 'int' ? Math.round(n) : n;
  }

  return String(value).trim() || null;
}
