import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import {
  buildUpdatedWorkbook,
  exportLeads,
  LEAD_STATUS_SCHEDULED,
  makeManualLead,
  nextRowIndex,
  parseWorkbook,
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
let nextFileId = 1

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
  // Every loaded workbook: { id, book, fileName, fileHandle, dirty }.
  const [files, setFiles] = useState([])
  // Leads from all files combined, in file order. Each lead carries its `fileId`.
  const [leads, setLeads] = useState([])
  const [index, setIndex] = useState(0)
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
  const filesById = useMemo(() => new Map(files.map((f) => [f.id, f])), [files])
  const dirty = files.some((f) => f.dirty)
  const hasFiles = files.length > 0

  const exportFiltered = (rows) => {
    if (!hasFiles || rows.length === 0) return
    const sources = new Set(rows.map((r) => r.lead.fileId))
    const base =
      sources.size === 1
        ? filesById.get([...sources][0]).fileName.replace(/\.xlsx?$/i, '')
        : 'leads-combined'
    const name = `${base}-${filterSlug(scheduleFilter)}.xlsx`
    try {
      exportLeads(filesById, rows.map((r) => r.lead), name)
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

  // Rebuild the combined lead list from the file list, reusing current leads for kept files.
  const applyFiles = (nextFiles, parsedLeads = new Map()) => {
    const byFile = new Map()
    for (const l of leads) {
      if (!byFile.has(l.fileId)) byFile.set(l.fileId, [])
      byFile.get(l.fileId).push(l)
    }
    setFiles(nextFiles)
    setLeads(nextFiles.flatMap((f) => parsedLeads.get(f.id) ?? byFile.get(f.id) ?? []))
    setIndex(0)
    setStudentIndices(null)
    setStudentPos(0)
  }

  // Add workbooks to the set. A file with the same name as a loaded one replaces it.
  const addFiles = (entries) => {
    let nextFiles = [...files]
    const parsedLeads = new Map()
    const errors = []
    let added = 0
    for (const { buffer, name, handle } of entries) {
      const existing = nextFiles.find((f) => f.fileName === name)
      if (existing?.dirty) {
        errors.push(`${name} has unsaved changes — save it before reloading`)
        continue
      }
      try {
        const parsed = parseWorkbook(buffer)
        const id = nextFileId++
        parsedLeads.set(
          id,
          parsed.leads.map((l) => ({ ...l, fileId: id })),
        )
        const file = { id, book: parsed, fileName: name, fileHandle: handle, dirty: false }
        nextFiles = existing
          ? nextFiles.map((f) => (f.id === existing.id ? file : f))
          : [...nextFiles, file]
        added += parsed.leads.length
      } catch (err) {
        errors.push(`${name}: ${err.message || 'could not read file'}`)
      }
    }
    if (parsedLeads.size > 0) {
      applyFiles(nextFiles, parsedLeads)
      // Add any dates found in the files that aren't buttons yet.
      const found = [
        ...new Set([...parsedLeads.values()].flat().map((l) => l.status).filter(Boolean)),
      ]
      setDateButtons((prev) => [...prev, ...found.filter((d) => !prev.includes(d))])
    }
    if (errors.length) flash('error', errors.join(' · '))
    else if (parsedLeads.size > 0) {
      const names = entries.map((e) => e.name).join(', ')
      flash('ok', `Loaded ${added} leads from ${names}`)
    }
  }

  const removeFile = (id) => {
    const file = filesById.get(id)
    if (file?.dirty && !window.confirm(`${file.fileName} has unsaved changes. Remove it anyway?`))
      return
    applyFiles(files.filter((f) => f.id !== id))
  }

  const openFile = async () => {
    if (!canWriteInPlace) {
      fileInputRef.current?.click()
      return
    }
    try {
      const handles = await window.showOpenFilePicker({
        multiple: true,
        types: [{ description: 'Excel workbook', accept: { [XLSX_MIME]: ['.xlsx'] } }],
      })
      const entries = await Promise.all(
        handles.map(async (handle) => {
          const file = await handle.getFile()
          return { buffer: await file.arrayBuffer(), name: file.name, handle }
        }),
      )
      addFiles(entries)
    } catch (err) {
      if (err.name !== 'AbortError') flash('error', err.message)
    }
  }

  const onFileInput = async (e) => {
    const picked = [...(e.target.files || [])]
    e.target.value = ''
    if (picked.length === 0) return
    const entries = await Promise.all(
      picked.map(async (file) => ({ buffer: await file.arrayBuffer(), name: file.name, handle: null })),
    )
    addFiles(entries)
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
    const fileId = leads[target]?.fileId
    setFiles((prev) => prev.map((f) => (f.id === fileId ? { ...f, dirty: true } : f)))
    setLeads((prev) =>
      prev.map((l, i) => {
        if (i !== target) return l
        const status = l.status === value ? '' : value
        // Keep lead_status in sync: scheduled when a date is set, original value when cleared.
        const leadStatus = status ? LEAD_STATUS_SCHEDULED : l.originalLeadStatus
        return { ...l, status, leadStatus }
      }),
    )
  }

  // Add a student by hand to one of the loaded files. All fields are optional, but a
  // completely blank entry is ignored. Returns true when the student was added.
  const addStudent = ({ name, phone, email, fileId, status }) => {
    if (![name, phone, email].some((v) => v.trim())) {
      flash('error', 'Enter at least a name, phone or email')
      return false
    }
    const file = filesById.get(fileId) ?? files[0]
    const rowIndex = nextRowIndex(
      file.book,
      leads.filter((l) => l.fileId === file.id),
    )
    const lead = makeManualLead({ name, phone, email }, file.id, rowIndex, status)
    // Append at the end so indices held by the students page stay valid.
    setLeads((prev) => [...prev, lead])
    setFiles((prev) => prev.map((f) => (f.id === file.id ? { ...f, dirty: true } : f)))
    const where = files.length > 1 ? ` to ${file.fileName}` : ''
    flash('ok', `Added ${lead.name || lead.phone || lead.email}${where}${status ? ` · coming ${status}` : ''}`)
    return true
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

  const writeFile = async (file) => {
    const { bytes, meta } = buildUpdatedWorkbook(
      file.book,
      leads.filter((l) => l.fileId === file.id),
    )
    if (file.fileHandle) {
      const writable = await file.fileHandle.createWritable()
      await writable.write(new Blob([bytes], { type: XLSX_MIME }))
      await writable.close()
    } else {
      const url = URL.createObjectURL(new Blob([bytes], { type: XLSX_MIME }))
      const a = document.createElement('a')
      a.href = url
      a.download = file.fileName || 'leads.xlsx'
      a.click()
      URL.revokeObjectURL(url)
    }
    // The workbook may have gained columns; remember where they are.
    setFiles((prev) =>
      prev.map((f) => (f.id === file.id ? { ...f, book: { ...f.book, ...meta }, dirty: false } : f)),
    )
  }

  // Saves every file with changes. With nothing changed, files that can't be written in
  // place are downloaded so the user can still grab a copy.
  const save = async () => {
    if (!hasFiles) return
    const dirtyFiles = files.filter((f) => f.dirty)
    const targets = dirtyFiles.length ? dirtyFiles : files.filter((f) => !f.fileHandle)
    const saved = []
    for (const file of targets) {
      try {
        await writeFile(file)
        saved.push(file.fileName)
      } catch (err) {
        if (err.name !== 'AbortError') {
          flash('error', `Saving ${file.fileName} failed: ${err.message}`)
          return
        }
      }
    }
    if (saved.length) flash('ok', `Saved ${saved.join(', ')}`)
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
          {files.map((f) => (
            <span key={f.id} className="file-chip" title={f.fileName}>
              <span className="file-chip-name">{f.fileName}</span>
              {f.dirty && <span className="dot" title="Unsaved changes" />}
              <button
                className="file-chip-remove"
                onClick={() => removeFile(f.id)}
                aria-label={`Remove ${f.fileName}`}
                title="Remove file"
              >
                ×
              </button>
            </span>
          ))}
          <button className="btn" onClick={openFile}>
            {hasFiles ? '+ Add files' : 'Open .xlsx'}
          </button>
          {hasFiles && page !== 'schedule' && (
            <button className="btn" onClick={goToSchedule}>
              See schedule
            </button>
          )}
          {hasFiles && (
            <button
              className="btn btn-primary"
              onClick={save}
              disabled={!dirty && files.every((f) => f.fileHandle)}
            >
              {files.every((f) => f.fileHandle)
                ? files.length > 1
                  ? 'Save to files'
                  : 'Save to file'
                : files.length > 1
                  ? 'Download updated files'
                  : 'Download updated file'}
            </button>
          )}
          <input
            ref={fileInputRef}
            type="file"
            accept=".xlsx,.xls"
            multiple
            hidden
            onChange={onFileInput}
          />
        </div>
      </header>

      {message && <div className={`toast toast-${message.type}`}>{message.text}</div>}

      {!hasFiles ? (
        <main className="empty">
          <div className="drop" onClick={openFile}>
            <div className="drop-icon">📄</div>
            <h1>Open leads spreadsheets</h1>
            <p>
              Choose one or more .xlsx exports — their leads are combined. Everything stays in
              your browser — nothing is uploaded.
            </p>
            <button className="btn btn-primary">Choose files</button>
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
          files={files}
          dateButtons={dateButtons}
          filter={scheduleFilter}
          onFilterChange={setScheduleFilter}
          onBack={goToDashboard}
          onOpenLead={openLeadFromSchedule}
          onExport={exportFiltered}
          onSeeStudents={goToStudents}
          onAddStudent={addStudent}
        />
      ) : page === 'students' ? (
        <StudentsPage
          leads={leads}
          filesById={files.length > 1 ? filesById : null}
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
          <p>No leads found in {files.length > 1 ? 'these files' : 'this file'}.</p>
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
            sourceFile={files.length > 1 ? filesById.get(lead.fileId)?.fileName : null}
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
