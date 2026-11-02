import LeadCard from './LeadCard'

/**
 * Dashboard-style view that steps through only the students chosen on the schedule page.
 * `indices` is the list of lead indices captured when the page was opened, so changing a
 * student's date here doesn't make them vanish from the list mid-session.
 */
export default function StudentsPage({
  leads,
  filesById,
  indices,
  pos,
  onPosChange,
  label,
  dateButtons,
  onSetStatus,
  onAddDateButton,
  onRemoveDateButton,
  onBack,
  onOpenLead,
  onExport,
}) {
  const total = indices.length
  const leadIndex = indices[pos]
  const lead = leads[leadIndex]

  const header = (
    <div className="students-head">
      <button className="link" onClick={onBack}>
        ← Back to schedule
      </button>
      <div className="students-title">
        <h1>Students</h1>
        <span className="badge">{label}</span>
        <span className="students-count">
          {total} {total === 1 ? 'student' : 'students'}
        </span>
      </div>
      <button className="btn" onClick={onExport} disabled={total === 0}>
        Export these (.xlsx)
      </button>
    </div>
  )

  if (!lead) {
    return (
      <main className="workspace">
        {header}
        <p className="schedule-empty">No students match this filter.</p>
      </main>
    )
  }

  const go = (delta) => onPosChange(Math.min(Math.max(pos + delta, 0), total - 1))

  return (
    <main className="workspace">
      {header}

      <section className="progress">
        <div>
          Student <strong>{pos + 1}</strong> of {total}
        </div>
        <div className="bar">
          <div className="bar-fill" style={{ width: `${((pos + 1) / total) * 100}%` }} />
        </div>
        <button className="link" onClick={() => onOpenLead(leadIndex)}>
          Open in dashboard
        </button>
      </section>

      <LeadCard
        lead={lead}
        sourceFile={filesById?.get(lead.fileId)?.fileName}
        dateButtons={dateButtons}
        onSetStatus={(value) => onSetStatus(leadIndex, value)}
        onAddDateButton={onAddDateButton}
        onRemoveDateButton={onRemoveDateButton}
      />

      <nav className="nav">
        <button className="btn btn-lg" onClick={() => go(-1)} disabled={pos === 0}>
          ← Previous
        </button>
        <div className="jump">
          <input
            type="number"
            min={1}
            max={total}
            value={pos + 1}
            onChange={(e) => {
              const n = Number(e.target.value)
              if (n >= 1 && n <= total) onPosChange(n - 1)
            }}
          />
        </div>
        <button className="btn btn-lg" onClick={() => go(1)} disabled={pos === total - 1}>
          Next →
        </button>
      </nav>
      <p className="shortcuts">Shortcuts: ← → to navigate · 1–9 to pick a date · Ctrl+S to save</p>
    </main>
  )
}
