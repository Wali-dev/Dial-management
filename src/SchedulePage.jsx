import { useMemo, useState } from 'react'
import { dialNumber } from './xlsxUtils'
import { ALL, UNSCHEDULED, filterLabel, filterLeads } from './scheduleFilter'

/**
 * Lists candidates by their coming date. Filter chips show a count per date and the
 * headline total always reflects the active filter. The filter lives in App so the
 * students page can show the same set.
 */
export default function SchedulePage({
  leads,
  files,
  dateButtons,
  filter,
  onFilterChange,
  onBack,
  onOpenLead,
  onExport,
  onSeeStudents,
  onAddStudent,
}) {
  const [adding, setAdding] = useState(false)
  // Count leads per coming date, ordered like the dashboard's date buttons,
  // followed by any dates that only exist in the file.
  const dateCounts = useMemo(() => {
    const counts = new Map(dateButtons.map((d) => [d, 0]))
    for (const l of leads) {
      if (l.status) counts.set(l.status, (counts.get(l.status) || 0) + 1)
    }
    return [...counts.entries()]
  }, [leads, dateButtons])

  const scheduledCount = useMemo(() => leads.filter((l) => l.status).length, [leads])
  const unscheduledCount = leads.length - scheduledCount
  const rows = useMemo(() => filterLeads(leads, filter), [leads, filter])
  const multiFile = files.length > 1
  const fileNames = useMemo(() => new Map(files.map((f) => [f.id, f.fileName])), [files])

  const setDate = (date) => onFilterChange({ ...filter, date })

  const chip = (value, label, count) => (
    <button
      key={value}
      className={`chip ${filter.date === value ? 'active' : ''}`}
      onClick={() => setDate(value)}
      aria-pressed={filter.date === value}
    >
      {label}
      <span className="chip-count">{count}</span>
    </button>
  )

  return (
    <main className="schedule">
      <div className="schedule-head">
        <button className="link" onClick={onBack}>
          ← Back to dashboard
        </button>
        <h1>Schedule</h1>
      </div>

      <section className="summary">
        <div className="summary-total">
          <span className="summary-number">{rows.length}</span>
          <span className="summary-label">
            {rows.length === 1 ? 'student' : 'students'} · {filterLabel(filter)}
          </span>
        </div>
        <div className="summary-side">
          <div className="summary-meta">
            {scheduledCount} scheduled · {unscheduledCount} not scheduled · {leads.length} total
            {multiFile && ` · from ${files.length} files`}
          </div>
          <div className="summary-actions">
            <button className="btn" onClick={() => onExport(rows)} disabled={rows.length === 0}>
              Export filtered (.xlsx)
            </button>
            <button className="btn" onClick={() => setAdding((v) => !v)} aria-expanded={adding}>
              + Add student
            </button>
            <button className="btn btn-primary" onClick={onSeeStudents} disabled={rows.length === 0}>
              See students
            </button>
          </div>
        </div>
      </section>

      {adding && (
        <AddStudentForm
          files={files}
          // Students added while viewing a date get that date, so they show up right away.
          status={filter.date !== ALL && filter.date !== UNSCHEDULED ? filter.date : ''}
          onAdd={onAddStudent}
          onClose={() => setAdding(false)}
        />
      )}

      <section className="filters">
        <div className="chips">
          {chip(ALL, 'All scheduled', scheduledCount)}
          {dateCounts.map(([date, count]) => chip(date, date, count))}
          {chip(UNSCHEDULED, 'Not scheduled', unscheduledCount)}
        </div>
        <input
          className="search"
          type="search"
          value={filter.query}
          onChange={(e) => onFilterChange({ ...filter, query: e.target.value })}
          placeholder="Search name, phone or address"
        />
      </section>

      {rows.length === 0 ? (
        <p className="schedule-empty">No students match this filter.</p>
      ) : (
        <div className="table-wrap">
          <table className="schedule-table">
            <thead>
              <tr>
                <th>#</th>
                <th>Name</th>
                <th>Phone</th>
                <th>Coming date</th>
                <th>Education</th>
                <th>Address</th>
                {multiFile && <th>File</th>}
              </tr>
            </thead>
            <tbody>
              {rows.map(({ lead, index }, i) => (
                <tr key={`${lead.fileId}:${lead.rowIndex}`} onClick={() => onOpenLead(index)} title="Open in dashboard">
                  <td className="num">{i + 1}</td>
                  <td className="strong">{lead.name || 'Unnamed lead'}</td>
                  <td className="num">
                    {lead.phone ? (
                      <a href={`tel:${dialNumber(lead.phone)}`} onClick={(e) => e.stopPropagation()}>
                        {dialNumber(lead.phone)}
                      </a>
                    ) : (
                      '—'
                    )}
                  </td>
                  <td>
                    {lead.status ? (
                      <span className="badge">{lead.status}</span>
                    ) : (
                      <span className="badge badge-muted">Not set</span>
                    )}
                  </td>
                  <td>{lead.education || '—'}</td>
                  <td className="wrap">{lead.address || '—'}</td>
                  {multiFile && <td className="muted">{fileNames.get(lead.fileId)}</td>}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </main>
  )
}

const EMPTY_STUDENT = { name: '', phone: '', email: '' }

/** Inline form for adding a student by hand. Every field is optional. */
function AddStudentForm({ files, status, onAdd, onClose }) {
  const [values, setValues] = useState(EMPTY_STUDENT)
  const [fileId, setFileId] = useState(files[0]?.id)
  const set = (key) => (e) => setValues((v) => ({ ...v, [key]: e.target.value }))

  const submit = (e) => {
    e.preventDefault()
    if (onAdd({ ...values, fileId, status })) setValues(EMPTY_STUDENT)
  }

  return (
    <form className="add-student" onSubmit={submit}>
      <div className="add-student-fields">
        <input value={values.name} onChange={set('name')} placeholder="Name" aria-label="Name" autoFocus />
        <input
          value={values.phone}
          onChange={set('phone')}
          placeholder="Phone"
          aria-label="Phone"
          type="tel"
          inputMode="tel"
        />
        <input
          value={values.email}
          onChange={set('email')}
          placeholder="Email"
          aria-label="Email"
          type="email"
        />
        {files.length > 1 && (
          <select
            value={fileId}
            onChange={(e) => setFileId(Number(e.target.value))}
            aria-label="Add to file"
            title="File the student is saved into"
          >
            {files.map((f) => (
              <option key={f.id} value={f.id}>
                {f.fileName}
              </option>
            ))}
          </select>
        )}
      </div>
      <div className="add-student-actions">
        <span className="add-student-hint">
          {status ? `Will be scheduled for ${status}` : 'Will be added as not scheduled'} · all fields
          optional
        </span>
        <button type="button" className="link" onClick={onClose}>
          Close
        </button>
        <button type="submit" className="btn btn-primary">
          Add student
        </button>
      </div>
    </form>
  )
}
