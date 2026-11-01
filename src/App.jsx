import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import {
  buildUpdatedWorkbook,
  exportLeads,
  LEAD_STATUS_COL,
  LEAD_STATUS_SCHEDULED,
  parseWorkbook,
  STATUS_COL,
} from './xlsxUtils'
import SchedulePage from './SchedulePage'
import StudentsPage from './StudentsPage'
import LeadCard from './LeadCard'
import { DEFAULT_SCHEDULE_FILTER, filterLabel, filterLeads, filterSlug } from './scheduleFilter'
import './App.css'

const DATE_BUTTONS_KEY = 'dialmanagement.dateButtons'
const DEFAULT_DATE_BUTTONS = ['16 Nov', '23 Nov']
const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
const canWriteInPlace = typeof window !== 'undefined' && 'showOpenFilePicker' in window

function loadDateButtons() {
  try {
    const saved = JSON.parse(localStorage.getItem(DATE_BUTTONS_KEY))
    if (Array.isArray(saved)) return saved
  } catch {
    /* ignore */
  }
  return DEFAULT_DATE_BUTTONS
}

const SCHEDULE_HASH = '#/schedule'
const STUDENTS_HASH = '#/students'
const currentPage = () => {
  if (window.location.hash === SCHEDULE_HASH) return 'schedule'
  if (window.location.hash === STUDENTS_HASH) return 'students'
  return 'dashboard'
}

export default function App() {
  const [book, setBook] = useState(null) // parsed workbook meta
  const [leads, setLeads] = useState([])
  const [index, setIndex] = useState(0)
  const [fileName, setFileName] = useState('')
  const [fileHandle, setFileHandle] = useState(null)
  const [dirty, setDirty] = useState(false)
  const [message, setMessage] = useState(null) // { type, text }
  const [dateButtons, setDateButtons] = useState(loadDateButtons)
  const [page, setPage] = useState(currentPage)
  const [scheduleFilter, setScheduleFilter] = useState(DEFAULT_SCHEDULE_FILTER)
  // Students page: lead indices captured when it was opened, and the current position.
  const [studentIndices, setStudentIndices] = useState(null)
  const [studentPos, setStudentPos] = useState(0)
  const fileInputRef = useRef(null)

  // Simple hash routing so the browser back button returns to the dashboard.
  useEffect(() => {
    const onHash = () => setPage(currentPage())
    window.addEventListener('hashchange', onHash)
    return () => window.removeEventListener('hashchange', onHash)
  }, [])

  const goToSchedule = () => {
    window.location.hash = SCHEDULE_HASH
  }
  const goToStudents = () => {
    setStudentIndices(filterLeads(leads, scheduleFilter).map((r) => r.index))
    setStudentPos(0)
    window.location.hash = STUDENTS_HASH
  }
  const exportFiltered = (rows) => {
    if (!book || rows.length === 0) return
    const base = (fileName || 'leads.xlsx').replace(/\.xlsx?$/i, '')
    const name = `${base}-${filterSlug(scheduleFilter)}.xlsx`
    try {
      exportLeads(book, rows.map((r) => r.lead), name)
      flash('ok', `Exported ${rows.length} students to ${name}`)
    } catch (err) {
      flash('error', `Export failed: ${err.message}`)
    }
  }
  const goToDashboard = () => {
    if (window.location.hash) window.location.hash = ''
    setPage('dashboard')
  }
  const openLeadFromSchedule = (i) => {
    setIndex(i)
    goToDashboard()
  }

  useEffect(() => {
    try {
      localStorage.setItem(DATE_BUTTONS_KEY, JSON.stringify(dateButtons))
    } catch {
      /* ignore */
    }
  }, [dateButtons])

  // Warn before closing the tab with unsaved changes.
  useEffect(() => {
    if (!dirty) return
    const handler = (e) => {
      e.preventDefault()
      e.returnValue = ''
    }
    window.addEventListener('beforeunload', handler)
    return () => window.removeEventListener('beforeunload', handler)
  }, [dirty])

  const flash = (type, text) => {
    setMessage({ type, text })
    setTimeout(() => setMessage(null), 3500)
  }

  const loadBuffer = (buffer, name, handle = null) => {
    try {
      const parsed = parseWorkbook(buffer)
      setBook(parsed)
      setLeads(parsed.leads)
      setIndex(0)
      setFileName(name)
      setFileHandle(handle)
      setDirty(false)
      // Add any dates found in the file that aren't buttons yet.
      const found = [...new Set(parsed.leads.map((l) => l.status).filter(Boolean))]
      setDateButtons((prev) => [...prev, ...found.filter((d) => !prev.includes(d))])
      flash('ok', `Loaded ${parsed.leads.length} leads from ${name}`)
    } catch (err) {
      flash('error', err.message || 'Could not read that file.')
    }
  }

  const openFile = async () => {
    if (!canWriteInPlace) {
      fileInputRef.current?.click()
      return
    }
    try {
      const [handle] = await window.showOpenFilePicker({
        types: [{ description: 'Excel workbook', accept: { [XLSX_MIME]: ['.xlsx'] } }],
      })
      const file = await handle.getFile()
      loadBuffer(await file.arrayBuffer(), file.name, handle)
    } catch (err) {
      if (err.name !== 'AbortError') flash('error', err.message)
    }
  }

  const onFileInput = async (e) => {
    const file = e.target.files?.[0]
    if (!file) return
    loadBuffer(await file.arrayBuffer(), file.name, null)
    e.target.value = ''
  }

  const lead = leads[index]
  const total = leads.length
  // Fall back to the live filter if the students page was reached via browser history.
  const students = studentIndices ?? filterLeads(leads, scheduleFilter).map((r) => r.index)
  const doneCount = useMemo(() => leads.filter((l) => l.status).length, [leads])

  const go = useCallback(
    (delta) => setIndex((i) => Math.min(Math.max(i + delta, 0), Math.max(total - 1, 0))),
    [total],
  )

  const setStatusAt = (target, value) => {
    setLeads((prev) =>
      prev.map((l, i) => {
        if (i !== target) return l
        const status = l.status === value ? '' : value
        // Keep lead_status in sync: scheduled when a date is set, original value when cleared.
        const leadStatus = status ? LEAD_STATUS_SCHEDULED : l.originalLeadStatus
        return { ...l, status, leadStatus }
      }),
    )
    setDirty(true)
  }

  // Returns true when the button was added.
  const addDateButton = (label) => {
    if (dateButtons.includes(label)) {
      flash('error', `"${label}" already exists`)
      return false
    }
    setDateButtons((prev) => [...prev, label])
    return true
  }

  const removeDateButton = (label) => setDateButtons((prev) => prev.filter((b) => b !== label))

  const save = async () => {
    if (!book) return
    const { bytes, col, leadCol } = buildUpdatedWorkbook(book, leads)
    try {
      if (fileHandle) {
        const writable = await fileHandle.createWritable()
        await writable.write(new Blob([bytes], { type: XLSX_MIME }))
        await writable.close()
        flash('ok', `Saved to ${fileName}`)
      } else {
        const url = URL.createObjectURL(new Blob([bytes], { type: XLSX_MIME }))
        const a = document.createElement('a')
        a.href = url
        a.download = fileName || 'leads.xlsx'
        a.click()
        URL.revokeObjectURL(url)
        flash('ok', `Downloaded updated ${a.download}`)
      }
      // The workbook now contains the status column; remember where it is.
      setBook((b) => {
        const headers = [...b.headers]
        headers[col] = STATUS_COL
        headers[leadCol] = LEAD_STATUS_COL
        return { ...b, statusColIdx: col, leadStatusColIdx: leadCol, headers }
      })
      setDirty(false)
    } catch (err) {
      if (err.name !== 'AbortError') flash('error', `Save failed: ${err.message}`)
    }
  }

  // Keyboard: ←/→ navigate, 1-9 pick a date button, Ctrl+S save.
  useEffect(() => {
    const onKey = (e) => {
      if (e.target.tagName === 'INPUT') return
      if (page === 'schedule') return
      const onStudents = page === 'students'
      const target = onStudents ? students[studentPos] : index
      const move = (d) =>
        onStudents
          ? setStudentPos((p) => Math.min(Math.max(p + d, 0), Math.max(students.length - 1, 0)))
          : go(d)
      if ((e.ctrlKey || e.metaKey) && e.key === 's') {
        e.preventDefault()
        save()
      } else if (e.key === 'ArrowRight') move(1)
      else if (e.key === 'ArrowLeft') move(-1)
      else if (/^[1-9]$/.test(e.key) && dateButtons[Number(e.key) - 1] && leads[target]) {
        setStatusAt(target, dateButtons[Number(e.key) - 1])
      }
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  })

  const jumpToNextPending = () => {
    const next = leads.findIndex((l, i) => i > index && !l.status)
    if (next !== -1) setIndex(next)
    else flash('ok', 'No pending leads after this one')
  }

  return (
    <div className="app">
      <header className="topbar">
        <div className="brand">
          <span className="brand-mark">☎</span> Dial Management
        </div>
        <div className="topbar-actions">
          {fileName && (
            <span className="file-chip" title={fileName}>
              {fileName}
              {dirty && <span className="dot" title="Unsaved changes" />}
            </span>
          )}
          <button className="btn" onClick={openFile}>
            {book ? 'Open another' : 'Open .xlsx'}
          </button>
          {book && page !== 'schedule' && (
            <button className="btn" onClick={goToSchedule}>
              See schedule
            </button>
          )}
          {book && (
            <button className="btn btn-primary" onClick={save} disabled={!dirty && !!fileHandle}>
              {fileHandle ? 'Save to file' : 'Download updated file'}
            </button>
          )}
          <input
            ref={fileInputRef}
            type="file"
            accept=".xlsx,.xls"
            hidden
            onChange={onFileInput}
          />
        </div>
      </header>

      {message && <div className={`toast toast-${message.type}`}>{message.text}</div>}

      {!book ? (
        <main className="empty">
          <div className="drop" onClick={openFile}>
            <div className="drop-icon">📄</div>
            <h1>Open a leads spreadsheet</h1>
            <p>
              Choose the .xlsx export. Everything stays in your browser — nothing is uploaded.
            </p>
            <button className="btn btn-primary">Choose file</button>
            {!canWriteInPlace && (
              <p className="hint">
                Your browser can’t write to files directly, so saving will download an updated
                copy. Use Chrome or Edge to save in place.
              </p>
            )}
          </div>
        </main>
      ) : page === 'schedule' ? (
        <SchedulePage
          leads={leads}
          dateButtons={dateButtons}
          filter={scheduleFilter}
          onFilterChange={setScheduleFilter}
          onBack={goToDashboard}
          onOpenLead={openLeadFromSchedule}
          onExport={exportFiltered}
          onSeeStudents={goToStudents}
        />
      ) : page === 'students' ? (
        <StudentsPage
          leads={leads}
          indices={students}
          pos={Math.min(studentPos, Math.max(students.length - 1, 0))}
          onPosChange={setStudentPos}
          label={filterLabel(scheduleFilter)}
          dateButtons={dateButtons}
          onSetStatus={setStatusAt}
          onAddDateButton={addDateButton}
          onRemoveDateButton={removeDateButton}
          onBack={goToSchedule}
          onOpenLead={openLeadFromSchedule}
          onExport={() => exportFiltered(students.map((i) => ({ lead: leads[i], index: i })))}
        />
      ) : total === 0 ? (
        <main className="empty">
          <p>No leads found in this file.</p>
        </main>
      ) : (
        <main className="workspace">
          <section className="progress">
            <div>
              Lead <strong>{index + 1}</strong> of {total}
            </div>
            <div className="bar">
              <div className="bar-fill" style={{ width: `${(doneCount / total) * 100}%` }} />
            </div>
            <div>
              {doneCount} scheduled · {total - doneCount} pending
            </div>
          </section>

          <LeadCard
            lead={lead}
            dateButtons={dateButtons}
            onSetStatus={(value) => setStatusAt(index, value)}
            onAddDateButton={addDateButton}
            onRemoveDateButton={removeDateButton}
          />

          <nav className="nav">
            <button className="btn btn-lg" onClick={() => go(-1)} disabled={index === 0}>
              ← Previous
            </button>
            <div className="jump">
              <input
                type="number"
                min={1}
                max={total}
                value={index + 1}
                onChange={(e) => {
                  const n = Number(e.target.value)
                  if (n >= 1 && n <= total) setIndex(n - 1)
                }}
              />
              <button className="link" onClick={jumpToNextPending}>
                Next pending
              </button>
            </div>
            <button className="btn btn-lg" onClick={() => go(1)} disabled={index === total - 1}>
              Next →
            </button>
          </nav>
          <p className="shortcuts">
            Shortcuts: ← → to navigate · 1–9 to pick a date · Ctrl+S to save
          </p>
        </main>
      )}
    </div>
  )
}
