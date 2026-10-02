/**
 * Dronetiks booth questionnaire — Google Apps Script web app.
 *
 * The spreadsheet this script is attached to holds:
 *   - the attendee list (the organiser's Excel file, imported as a tab)
 *   - a "Responses" tab that this script creates and fills, one row per lead
 *
 * Setup and deployment steps are in README.md.
 */

const SETTINGS = {
  // Tab holding the attendee list. If no tab has this name, the first tab
  // whose header row has first/last-name columns is used, so an imported
  // Excel tab works without renaming it.
  attendeesSheet: 'Participants',
  responsesSheet: 'Responses',
  // Only needed for a standalone script. Leave empty when the script was
  // created from the spreadsheet (Extensions > Apps Script).
  spreadsheetId: '',
  // Optional booth PIN: add a Script Property with this name
  // (Project Settings > Script Properties). When set, each device must enter
  // it once before attendees are shown or answers are accepted.
  pinProperty: 'BOOTH_PIN',
};

/**
 * The questionnaire. Edit questions here; the page and the Responses columns
 * follow automatically.
 *   header   column name in the Responses tab (keep it stable once the event
 *            starts: a new header becomes a new column)
 *   fill     attendee-list field that pre-fills the input
 *   keep     value stays filled in for the next lead
 *   default  initial value ('today' = the device's current date)
 */
const FORM = {
  title: 'Dronetiks Trade Show Assessment Questionnaire',
  subtitle: 'Lead-qualification survey for warehouse and distribution center prospects. ' +
    'Search for the attendee to pre-fill Lead Information.',
  sections: [
    {
      title: 'Lead Information',
      fields: [
        { id: 'date', header: 'Date', label: 'Date', type: 'date', default: 'today' },
        { id: 'eventName', header: 'Event Name', label: 'Event Name', type: 'text', default: 'Edge_2026', keep: true },
        { id: 'firstName', header: 'First Name', label: 'First Name', type: 'text', required: true, fill: 'first' },
        { id: 'lastName', header: 'Last Name', label: 'Last Name', type: 'text', required: true, fill: 'last' },
        { id: 'company', header: 'Company Name', label: 'Company Name', type: 'text', required: true, fill: 'company' },
        { id: 'jobTitle', header: 'Job Title', label: 'Job Title', type: 'text', fill: 'title' },
        { id: 'email', header: 'Corporate Email', label: 'Corporate Email', type: 'email', required: true, fill: 'email', hint: 'example@example.com' },
      ],
    },
    {
      title: 'Facility Footprint & Scale',
      fields: [
        {
          id: 'squareFootage', header: 'Square Footage', type: 'radio', required: true,
          label: 'Total square footage of your primary distribution or fulfillment center',
          options: [
            'Under 50,000 sq. ft. (Low priority — manual methods typically suffice)',
            '50,000 – 150,000 sq. ft. (Good target — standard pilot size)',
            '150,000 – 500,000 sq. ft. (High value — severe manual audit pain points)',
            '500,000+ sq. ft. (Enterprise tier — massive automated recurring revenue potential)',
          ],
        },
        {
          id: 'rackHeight', header: 'Rack Height', type: 'radio', required: true,
          label: 'Maximum vertical clearance / top shelf height of pallet racks',
          options: [
            'Under 15 feet (Easy human reach)',
            '15 to 25 feet (Requires scissor lifts or high-reach stock pickers)',
            '26 to 40 feet (Severe human safety risk — Prime Drone Candidate)',
            '40+ feet (High-bay distribution center — Immediate Sales Priority)',
          ],
        },
        {
          id: 'aisles', header: 'Rack Aisles', type: 'radio', required: true,
          label: 'Average total inventory rack aisles per facility',
          options: ['1 – 20 aisles', '21 – 50 aisles', '51 – 100 aisles', '100+ aisles'],
        },
        {
          id: 'rarelyAudited', header: 'Rarely Audited Zones', type: 'textarea',
          label: 'Are there specific high-bay racks or narrow-aisle zones that your human teams rarely audit ' +
            'due to high reach safety constraints or forklift collision risks?',
        },
        {
          id: 'lighting', header: 'Aisle Lighting', type: 'textarea',
          label: 'Do your aisles utilize motion-activated LED lighting systems that shut off when pickers leave, ' +
            'or do you maintain baseline ambient lighting overnight?',
        },
      ],
    },
    {
      title: 'Operational Bottlenecks',
      fields: [
        {
          id: 'cycleCount', header: 'Cycle Count Frequency', type: 'radio', required: true,
          label: 'How often do you perform wall-to-wall cycle counting or full physical inventory checks?',
          options: ['Daily / Weekly rolling checks', 'Monthly', 'Quarterly / Bi-annually', 'Annually only'],
        },
        {
          id: 'lowActivity', header: 'Lowest-Activity Windows', type: 'textarea',
          label: "When are your facility's lowest-activity windows or shift changes during the 24-hour cycle? " +
            '(e.g., 2:00 AM to 5:00 AM? Low-volume weekends?)',
        },
        {
          id: 'trackingMethod', header: 'Location Tracking Method', type: 'radio', required: true,
          label: 'What is your primary methodology for tracking warehouse location assets?',
          options: [
            'Standard 1D/2D Barcode scanning',
            'Manual paper pick sheets & clipboards',
            'Existing Passive RFID tag infrastructure',
            'Active RFID / RTLS (Real-Time Location Systems)',
          ],
        },
        {
          id: 'wms', header: 'WMS', type: 'radio', required: true,
          label: 'What core Warehouse Management System (WMS) do you use?',
          options: [
            'SAP EWM / Oracle WMS',
            'Manhattan Active / Blue Yonder',
            'Infor / NetSuite WMS',
            'Proprietary / Custom Built-in System',
          ],
        },
      ],
    },
  ],
};

const LEADING_COLUMNS = ['Submitted At', 'Updated At', 'Captured By'];
const TRAILING_COLUMNS = ['Access Code', 'Zip', 'Country', 'Registration Status', 'Attendee Key', 'Response ID'];

// Attendee-list headers are matched after lower-casing, dropping
// non-alphanumerics and a leading "attendee" ("AttendeeFirstName" -> "firstname").
const ATTENDEE_COLUMNS = {
  code: ['accesscode', 'code', 'badgeid', 'registrationid', 'id'],
  first: ['firstname', 'first'],
  last: ['lastname', 'last', 'surname'],
  title: ['jobtitle', 'title'],
  company: ['company', 'companyname', 'organization', 'organisation'],
  email: ['email', 'emailaddress', 'corporateemail'],
  zip: ['zippostalcode', 'zip', 'zipcode', 'postalcode'],
  country: ['country'],
  status: ['addedcancelled', 'status', 'registrationstatus'],
};

const EMAIL_PATTERN = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
const MAX_ANSWER_LENGTH = 5000;

// ---- Web app entry points ---------------------------------------------------

function doGet() {
  const page = HtmlService.createTemplateFromFile('Index');
  page.formJson = JSON.stringify(FORM).replace(/</g, '\\u003c');
  return page.evaluate()
    .setTitle('Dronetiks Booth Questionnaire')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Inlines another HTML file of the project; a missing file is skipped. */
function include_(name) {
  try {
    return HtmlService.createHtmlOutputFromFile(name).getContent();
  } catch (e) {
    return '';
  }
}

/** Attendee list plus which attendees already have a response. */
function getBootstrap(pin) {
  const denied = checkPin_(pin);
  if (denied) return denied;
  const ss = spreadsheet_();
  return { ok: true, attendees: readAttendees_(ss), submitted: submittedIndex_(ss) };
}

/** Saved answers for one attendee, so a repeat visit edits the same row. */
function getResponse(pin, attendeeKey) {
  const denied = checkPin_(pin);
  if (denied) return denied;
  const ss = spreadsheet_();
  const sheet = ss.getSheetByName(SETTINGS.responsesSheet);
  if (!attendeeKey || !sheet || sheet.getLastRow() < 2) return { ok: true, found: false };

  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(String);
  const keyCol = headers.indexOf('Attendee Key');
  if (keyCol < 0) return { ok: true, found: false };

  const tz = ss.getSpreadsheetTimeZone();
  for (let r = data.length - 1; r >= 1; r--) {
    if (String(data[r][keyCol]) !== attendeeKey) continue;
    const row = data[r];
    const values = {};
    fields_().forEach(function (f) {
      const c = headers.indexOf(f.header);
      if (c >= 0) values[f.id] = toText_(row[c], tz, f.type === 'date' ? 'yyyy-MM-dd' : null);
    });
    return {
      ok: true,
      found: true,
      values: values,
      at: toText_(row[headers.indexOf('Updated At')], tz),
      by: toText_(row[headers.indexOf('Captured By')], tz),
    };
  }
  return { ok: true, found: false };
}

/**
 * Saves one questionnaire. Listed attendees get one row each (a repeat
 * submission updates it); walk-ins are keyed by the page's response ID, so a
 * retried submission never creates a duplicate row.
 */
function submitResponse(pin, payload) {
  const denied = checkPin_(pin);
  if (denied) return denied;
  payload = payload || {};
  const cleaned = clean_(payload.values || {});
  if (cleaned.errors.length) return { ok: false, error: cleaned.errors.join('\n') };
  const attendeeKey = String(payload.attendeeKey || '');
  const responseId = String(payload.responseId || '') || Utilities.getUuid();

  const ss = spreadsheet_();
  const attendee = attendeeKey ? findAttendee_(ss, attendeeKey) : null;
  const now = new Date();
  const record = { 'Updated At': now, 'Captured By': String(payload.capturedBy || '').trim().slice(0, 100) };
  fields_().forEach(function (f) {
    record[f.header] = cleaned.values[f.id];
  });
  record['Attendee Key'] = attendeeKey;
  if (attendee) {
    record['Access Code'] = attendee.code;
    record['Zip'] = attendee.zip;
    record['Country'] = attendee.country;
    record['Registration Status'] = attendee.status;
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let rowNum;
  try {
    const target = responsesSheet_(ss);
    const sheet = target.sheet;
    const headers = target.headers;
    rowNum = findRow_(sheet, headers, attendeeKey ? 'Attendee Key' : 'Response ID', attendeeKey || responseId);
    if (rowNum) {
      writeCells_(sheet, headers, rowNum, record);
    } else {
      record['Submitted At'] = now;
      record['Response ID'] = responseId;
      sheet.appendRow(headers.map(function (h) {
        return h in record ? cell_(h, record[h]) : '';
      }));
    }
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return { ok: true, updated: !!rowNum, at: toText_(now, ss.getSpreadsheetTimeZone()) };
}

// ---- Helpers ----------------------------------------------------------------

function spreadsheet_() {
  return SETTINGS.spreadsheetId
    ? SpreadsheetApp.openById(SETTINGS.spreadsheetId)
    : SpreadsheetApp.getActiveSpreadsheet();
}

function checkPin_(pin) {
  const expected = PropertiesService.getScriptProperties().getProperty(SETTINGS.pinProperty);
  if (!expected || String(pin || '').trim() === String(expected).trim()) return null;
  return { ok: false, authError: true, error: pin ? 'That PIN is not correct.' : 'Enter the booth PIN.' };
}

function fields_() {
  return FORM.sections.reduce(function (all, s) {
    return all.concat(s.fields);
  }, []);
}

/** Trims answers and checks them against the form definition. */
function clean_(input) {
  const values = {};
  const errors = [];
  fields_().forEach(function (f) {
    const v = input[f.id] == null ? '' : String(input[f.id]).trim();
    values[f.id] = v;
    if (!v) {
      if (f.required) errors.push(f.label + ' is required.');
    } else if (v.length > MAX_ANSWER_LENGTH) {
      errors.push(f.label + ' is too long.');
    } else if (f.options && f.options.indexOf(v) === -1) {
      errors.push(f.label + ': "' + v + '" is not one of the options.');
    } else if (f.type === 'email' && !EMAIL_PATTERN.test(v)) {
      errors.push(f.label + ' is not a valid email address.');
    } else if (f.type === 'date' && !/^\d{4}-\d{2}-\d{2}$/.test(v)) {
      errors.push(f.label + ' is not a valid date.');
    }
  });
  return { values: values, errors: errors };
}

function normHeader_(h) {
  return String(h).toLowerCase().replace(/[^a-z0-9]/g, '').replace(/^attendee/, '');
}

function mapColumns_(headerRow) {
  const cols = {};
  const names = headerRow.map(normHeader_);
  Object.keys(ATTENDEE_COLUMNS).forEach(function (field) {
    ATTENDEE_COLUMNS[field].some(function (alias) {
      const i = names.indexOf(alias);
      if (i >= 0) cols[field] = i;
      return i >= 0;
    });
  });
  return cols;
}

function readAttendees_(ss) {
  const named = ss.getSheetByName(SETTINGS.attendeesSheet);
  const candidates = named ? [named] : ss.getSheets().filter(function (s) {
    return s.getName() !== SETTINGS.responsesSheet;
  });
  for (let s = 0; s < candidates.length; s++) {
    const values = candidates[s].getDataRange().getValues();
    for (let r = 0; r < Math.min(values.length, 10); r++) {
      const cols = mapColumns_(values[r]);
      if (cols.first != null && cols.last != null) return toAttendees_(values.slice(r + 1), cols);
    }
  }
  return [];
}

function toAttendees_(rows, cols) {
  const seen = {};
  const out = [];
  rows.forEach(function (row) {
    const get = function (field) {
      const v = cols[field] == null ? null : row[cols[field]];
      return v == null ? '' : String(v).trim();
    };
    const a = {
      code: get('code'), first: get('first'), last: get('last'), title: get('title'),
      company: get('company'), email: get('email'), zip: get('zip'), country: get('country'), status: get('status'),
    };
    if (!a.first && !a.last) return;
    // Excel stores US zips as numbers, which drops leading zeros (07401 -> 7401).
    if (/^\d{3,4}$/.test(a.zip) && (!a.country || /^(us|usa|united states)/i.test(a.country))) {
      a.zip = ('00000' + a.zip).slice(-5);
    }
    let key = a.code ? 'C:' + a.code : 'N:' + [a.first, a.last, a.company].join('|').toLowerCase();
    while (seen[key]) key += '+';
    seen[key] = true;
    a.key = key;
    out.push(a);
  });
  return out;
}

function findAttendee_(ss, key) {
  const list = readAttendees_(ss);
  for (let i = 0; i < list.length; i++) if (list[i].key === key) return list[i];
  return null;
}

function submittedIndex_(ss) {
  const sheet = ss.getSheetByName(SETTINGS.responsesSheet);
  const index = {};
  if (!sheet || sheet.getLastRow() < 2) return index;
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(String);
  const keyCol = headers.indexOf('Attendee Key');
  const atCol = headers.indexOf('Updated At');
  const byCol = headers.indexOf('Captured By');
  if (keyCol < 0) return index;
  const tz = ss.getSpreadsheetTimeZone();
  for (let r = 1; r < data.length; r++) {
    const key = String(data[r][keyCol]);
    if (key) index[key] = { at: atCol < 0 ? '' : toText_(data[r][atCol], tz), by: byCol < 0 ? '' : toText_(data[r][byCol], tz) };
  }
  return index;
}

/** Returns the Responses tab, creating it and any missing header columns. */
function responsesSheet_(ss) {
  const sheet = ss.getSheetByName(SETTINGS.responsesSheet) || ss.insertSheet(SETTINGS.responsesSheet);
  const lastCol = sheet.getLastColumn();
  const headers = lastCol ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String) : [];
  const wanted = LEADING_COLUMNS.concat(fields_().map(function (f) { return f.header; }), TRAILING_COLUMNS);
  const missing = wanted.filter(function (h) { return headers.indexOf(h) === -1; });
  if (missing.length) {
    const needed = headers.length + missing.length - sheet.getMaxColumns();
    if (needed > 0) sheet.insertColumnsAfter(sheet.getMaxColumns(), needed);
    sheet.getRange(1, headers.length + 1, 1, missing.length).setValues([missing]).setFontWeight('bold');
    if (!headers.length) sheet.setFrozenRows(1);
    Array.prototype.push.apply(headers, missing);
  }
  return { sheet: sheet, headers: headers };
}

function findRow_(sheet, headers, column, value) {
  const col = headers.indexOf(column);
  const n = sheet.getLastRow() - 1;
  if (col < 0 || n < 1 || !value) return 0;
  const values = sheet.getRange(2, col + 1, n, 1).getValues();
  for (let i = values.length - 1; i >= 0; i--) {
    if (String(values[i][0]) === value) return i + 2;
  }
  return 0;
}

/** Overwrites only this app's cells in a row, leaving any columns the team added alone. */
function writeCells_(sheet, headers, rowNum, record) {
  const cols = headers
    .map(function (h, i) { return h in record ? i : -1; })
    .filter(function (i) { return i >= 0; });
  let start = 0;
  while (start < cols.length) {
    let end = start;
    while (end + 1 < cols.length && cols[end + 1] === cols[end] + 1) end++;
    const run = cols.slice(start, end + 1).map(function (i) { return cell_(headers[i], record[headers[i]]); });
    sheet.getRange(rowNum, cols[start] + 1, 1, run.length).setValues([run]);
    start = end + 1;
  }
}

/**
 * Text answers are written with a leading apostrophe so Sheets keeps them as
 * typed: no formulas from "=..." input, no "1/2" turning into a date, no lost
 * zip-code zeros. The Date answer is left as yyyy-mm-dd so it becomes a date.
 */
function cell_(header, value) {
  if (isDate_(value) || typeof value !== 'string' || value === '' || header === dateHeader_()) return value;
  return "'" + value;
}

function dateHeader_() {
  const f = fields_().filter(function (x) { return x.type === 'date'; })[0];
  return f ? f.header : null;
}

function isDate_(v) {
  return Object.prototype.toString.call(v) === '[object Date]';
}

function toText_(v, tz, format) {
  if (isDate_(v)) return Utilities.formatDate(v, tz, format || 'MMM d, h:mm a');
  return v == null ? '' : String(v);
}
