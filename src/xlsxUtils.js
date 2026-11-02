import * as XLSX from 'xlsx'

// Column that stores the "coming date" status in the spreadsheet.
export const STATUS_COL = 'coming_date'
// Column that stores the overall lead status; updated whenever the coming date changes.
export const LEAD_STATUS_COL = 'lead_status'
export const LEAD_STATUS_SCHEDULED = 'scheduled'
// Extra column added to exports that combine leads from several files.
export const SOURCE_FILE_COL = 'source_file'
// Fields a manually added student can have, and the header used if the sheet lacks one.
const MANUAL_FIELDS = { name: 'name', phone: 'phone_number', email: 'email' }

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
  return { workbook, sheetName, headers, leads, statusColIdx, leadStatusColIdx, fieldCols: cols }
}

/** Row index (0-based) for the next row appended to the workbook's sheet. */
export function nextRowIndex({ workbook, sheetName }, leads) {
  const ref = workbook.Sheets[sheetName]['!ref']
  const lastRow = ref ? XLSX.utils.decode_range(ref).e.r : 0
  return Math.max(lastRow, ...leads.map((l) => l.rowIndex)) + 1
}

/** Build a lead for a student added by hand. Every field is optional. */
export function makeManualLead({ name = '', phone = '', email = '' }, fileId, rowIndex, status) {
  return {
    rowIndex,
    fileId,
    isNew: true, // its name/phone/email cells still need writing to the sheet
    name: name.trim(),
    email: email.trim(),
    phone: normalisePhone(phone),
    address: '',
    education: '',
    age: '',
    currentStatus: '',
    status: status || '',
    leadStatus: status ? LEAD_STATUS_SCHEDULED : '',
    originalLeadStatus: '',
  }
}

/** Write a column of values into the sheet (adding the column if needed). Returns its index. */
function writeColumn(ws, range, headers, colIdx, header, leads, value) {
  let col = colIdx
  if (col === -1) {
    col = Math.max(range.e.c + 1, headers.length)
    ws[XLSX.utils.encode_cell({ r: 0, c: col })] = { t: 's', v: header }
  }
  for (const lead of leads) {
    const addr = XLSX.utils.encode_cell({ r: lead.rowIndex, c: col })
    const v = value(lead)
    if (v) ws[addr] = { t: 's', v }
    else delete ws[addr]
  }
  if (col > range.e.c) range.e.c = col
  for (const lead of leads) if (lead.rowIndex > range.e.r) range.e.r = lead.rowIndex
  return col
}

/**
 * Write coming dates, lead statuses and any manually added students into the sheet.
 * Returns the xlsx bytes plus the updated column layout to merge back into the book meta.
 */
export function buildUpdatedWorkbook(
  { workbook, sheetName, headers, statusColIdx, leadStatusColIdx, fieldCols = {} },
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
  const nextHeaders = [...headers]
  nextHeaders[col] = STATUS_COL
  nextHeaders[leadCol] = LEAD_STATUS_COL
  const nextFieldCols = { ...fieldCols }
  const manual = leads.filter((l) => l.isNew)
  if (manual.length) {
    for (const [field, header] of Object.entries(MANUAL_FIELDS)) {
      const existing = fieldCols[field] ?? -1
      const c = writeColumn(ws, range, nextHeaders, existing, header, manual, (l) => l[field])
      nextFieldCols[field] = c
      if (existing === -1) nextHeaders[c] = header
    }
  }
  ws['!ref'] = XLSX.utils.encode_range(range)

  return {
    bytes: XLSX.write(workbook, { bookType: 'xlsx', type: 'array' }),
    meta: {
      statusColIdx: col,
      leadStatusColIdx: leadCol,
      headers: nextHeaders,
      fieldCols: nextFieldCols,
    },
  }
}

/**
 * Download a new .xlsx containing only the given leads, which may come from several
 * source workbooks. Columns are the union of every source sheet's headers (matched by
 * name), with coming_date / lead_status reflecting the current (possibly unsaved) values.
 * When the leads span more than one file a source_file column is added.
 *
 * `files` is a Map of fileId -> { book, fileName }.
 */
export function exportLeads(files, leads, fileName) {
  const sheets = new Map() // fileId -> { rows, colMap }
  const headers = []
  const headerIdx = new Map()
  // Recognised fields (name, phone, ...) share one column even when files name them
  // differently; the first file's header wins.
  const fieldHeader = new Map()
  const addHeader = (h) => {
    if (!headerIdx.has(h)) headerIdx.set(h, headers.push(h) - 1)
    return headerIdx.get(h)
  }

  for (const lead of leads) {
    if (sheets.has(lead.fileId)) continue
    const { workbook, sheetName, fieldCols = {} } = files.get(lead.fileId).book
    const fieldOf = new Map(Object.entries(fieldCols).map(([field, idx]) => [idx, field]))
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
      header: 1,
      defval: '',
      raw: true,
    })
    // Map each source column to its position in the combined header row.
    const colMap = (rows[0] || []).map((h, i) => {
      let name = String(h).trim()
      const field = fieldOf.get(i)
      if (field && name) {
        if (!fieldHeader.has(field)) fieldHeader.set(field, name)
        name = fieldHeader.get(field)
      }
      return name ? addHeader(name) : -1
    })
    sheets.set(lead.fileId, { rows, colMap })
  }
  const statusIdx = addHeader(STATUS_COL)
  const leadStatusIdx = addHeader(LEAD_STATUS_COL)
  const sourceIdx = sheets.size > 1 ? addHeader(SOURCE_FILE_COL) : -1

  const out = [headers]
  for (const lead of leads) {
    const { rows, colMap } = sheets.get(lead.fileId)
    const src = rows[lead.rowIndex] || []
    const row = new Array(headers.length).fill('')
    colMap.forEach((to, from) => {
      if (to !== -1 && src[from] !== undefined) row[to] = src[from]
    })
    row[statusIdx] = lead.status
    row[leadStatusIdx] = lead.leadStatus
    if (sourceIdx !== -1) row[sourceIdx] = files.get(lead.fileId).fileName
    if (lead.isNew) {
      // Manually added: the sheet may not hold this row yet, so take values from the lead.
      for (const [field, header] of Object.entries(MANUAL_FIELDS)) {
        if (!fieldHeader.has(field)) fieldHeader.set(field, header)
        row[addHeader(fieldHeader.get(field))] = lead[field]
      }
    }
    out.push(row)
  }

  const book = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet(out), 'Students')
  XLSX.writeFile(book, fileName)
}
