import { useMemo } from 'react'
import { dialNumber } from './xlsxUtils'
import { ALL, UNSCHEDULED, filterLabel, filterLeads } from './scheduleFilter'

/**
 * Lists candidates by their coming date. Filter chips show a count per date and the
 * headline total always reflects the active filter. The filter lives in App so the
 * students page can show the same set.
 */
export default function SchedulePage({
  leads,
  dateButtons,
  filter,
  onFilterChange,
  onBack,
  onOpenLead,
  onExport,
  onSeeStudents,
}) {
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
          </div>
          <div className="summary-actions">
            <button className="btn" onClick={() => onExport(rows)} disabled={rows.length === 0}>
              Export filtered (.xlsx)
            </button>
            <button className="btn btn-primary" onClick={onSeeStudents} disabled={rows.length === 0}>
              See students
            </button>
          </div>
        </div>
      </section>

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
              </tr>
            </thead>
            <tbody>
              {rows.map(({ lead, index }, i) => (
                <tr key={lead.rowIndex} onClick={() => onOpenLead(index)} title="Open in dashboard">
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
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </main>
  )
}
