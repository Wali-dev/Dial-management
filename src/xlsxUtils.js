import * as XLSX from 'xlsx'

// Column that stores the "coming date" status in the spreadsheet.
export const STATUS_COL = 'coming_date'
// Column that stores the overall lead status; updated whenever the coming date changes.
export const LEAD_STATUS_COL = 'lead_status'
export const LEAD_STATUS_SCHEDULED = 'scheduled'

// Candidate header names for each field we display. First match wins.
const FIELD_MATCHERS = {
  name: [/পূর্ণ_নাম/, /full[_ ]?name/i, /^name$/i],
  email: [/^email$/i, /e-?mail/i],
  phone: [/^phone_number$/i, /phone/i, /mobile/i],
  address: [/ঠিকানা/, /address/i],
  education: [/শিক্ষাগত/, /education/i, /qualification/i],
  age: [/^age$/i, /বয়স/],
  currentStatus: [/^current_status$/i, /occupation/i],
}

function findColumn(headers, patterns) {
  for (const re of patterns) {
    const idx = headers.findIndex((h) => re.test(String(h).trim()))
    if (idx !== -1) return idx
  }
  return -1
}

// Normalise phone numbers like 8801912078506 / 01912078506 / "p:+8801..." to +8801...
export function normalisePhone(raw) {
  let s = String(raw ?? '').trim()
  if (!s) return ''
  s = s.replace(/^p:/i, '').replace(/[^\d+]/g, '')
  if (s.startsWith('+')) return s
  if (s.startsWith('880')) return '+' + s
  if (s.startsWith('0') && s.length === 11) return '+88' + s
  return s
}

// Local Bangladeshi format for dialing (+8801XXXXXXXXX -> 01XXXXXXXXX), so the
// phone's dialer doesn't add any prefix in front of the "+" country code.
export function dialNumber(phone) {
  const s = String(phone ?? '')
  if (s.startsWith('+880')) return '0' + s.slice(4).replace(/^0+/, '')
  return s.replace(/^\+/, '')
}

/**
 * Parse an ArrayBuffer into { workbook, sheetName, headers, leads, statusColIdx }.
 * Each lead keeps its sheet row index so we can write the status back later.
 */
export function parseWorkbook(buffer) {
  const workbook = XLSX.read(buffer, { type: 'array' })
  const sheetName = workbook.SheetNames[0]
  const ws = workbook.Sheets[sheetName]
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: true })
  if (rows.length === 0) throw new Error('The first sheet is empty.')

  const headers = rows[0].map((h) => String(h))
  const cols = Object.fromEntries(
    Object.entries(FIELD_MATCHERS).map(([k, pats]) => [k, findColumn(headers, pats)]),
  )
  if (cols.phone === -1) throw new Error('Could not find a phone column in the sheet.')
  const statusColIdx = headers.findIndex((h) => h.trim() === STATUS_COL)
  const leadStatusColIdx = headers.findIndex((h) => h.trim() === LEAD_STATUS_COL)

  const leads = []
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r]
    if (!row || row.every((c) => c === '' || c == null)) continue
    const get = (i) => (i === -1 ? '' : String(row[i] ?? '').trim())
    leads.push({
      rowIndex: r, // 0-based sheet row (header is row 0)
      name: get(cols.name),
      email: get(cols.email),
      address: get(cols.address),
      phone: normalisePhone(row[cols.phone]),
      education: get(cols.education),
      age: get(cols.age),
      currentStatus: get(cols.currentStatus),
      status: statusColIdx === -1 ? '' : get(statusColIdx),
      leadStatus: get(leadStatusColIdx),
      originalLeadStatus: get(leadStatusColIdx),
    })
  }
  return { workbook, sheetName, headers, leads, statusColIdx, leadStatusColIdx }
}

/** Write a column of values into the sheet (adding the column if needed). Returns its index. */
function writeColumn(ws, range, headers, colIdx, header, leads, value) {
  let col = colIdx
  if (col === -1) col = Math.max(range.e.c + 1, headers.length)
  ws[XLSX.utils.encode_cell({ r: 0, c: col })] = { t: 's', v: header }
  for (const lead of leads) {
    const addr = XLSX.utils.encode_cell({ r: lead.rowIndex, c: col })
    const v = value(lead)
    if (v) ws[addr] = { t: 's', v }
    else delete ws[addr]
  }
  if (col > range.e.c) range.e.c = col
  return col
}

/** Write coming dates and lead statuses into the sheet and return xlsx bytes. */
export function buildUpdatedWorkbook(
  { workbook, sheetName, headers, statusColIdx, leadStatusColIdx },
  leads,
) {
  const ws = workbook.Sheets[sheetName]
  const range = XLSX.utils.decode_range(ws['!ref'])
  const col = writeColumn(ws, range, headers, statusColIdx, STATUS_COL, leads, (l) => l.status)
  const leadCol = writeColumn(
    ws,
    range,
    headers,
    leadStatusColIdx,
    LEAD_STATUS_COL,
    leads,
    (l) => l.leadStatus,
  )
  ws['!ref'] = XLSX.utils.encode_range(range)

  return { bytes: XLSX.write(workbook, { bookType: 'xlsx', type: 'array' }), col, leadCol }
}

/**
 * Download a new .xlsx containing only the given leads: every original column from the
 * source sheet, with coming_date / lead_status reflecting the current (possibly unsaved) values.
 */
export function exportLeads({ workbook, sheetName }, leads, fileName) {
  const ws = workbook.Sheets[sheetName]
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: true })
  const headers = rows[0].map((h) => String(h))
  let statusIdx = headers.findIndex((h) => h.trim() === STATUS_COL)
  if (statusIdx === -1) statusIdx = headers.push(STATUS_COL) - 1
  let leadStatusIdx = headers.findIndex((h) => h.trim() === LEAD_STATUS_COL)
  if (leadStatusIdx === -1) leadStatusIdx = headers.push(LEAD_STATUS_COL) - 1

  const out = [headers]
  for (const lead of leads) {
    const row = [...(rows[lead.rowIndex] || [])]
    while (row.length < headers.length) row.push('')
    row[statusIdx] = lead.status
    row[leadStatusIdx] = lead.leadStatus
    out.push(row)
  }

  const book = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet(out), 'Students')
  XLSX.writeFile(book, fileName)
}
