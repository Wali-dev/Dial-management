import { useState } from 'react'
import { QRCodeSVG } from 'qrcode.react'
import { dialNumber } from './xlsxUtils'

/** The dial card: QR code, lead details and coming-date buttons. Shared by dashboard and students page. */
export default function LeadCard({ lead, dateButtons, onSetStatus, onAddDateButton, onRemoveDateButton }) {
  const [newLabel, setNewLabel] = useState('')
  const [editMode, setEditMode] = useState(false)

  const submitNewLabel = (e) => {
    e.preventDefault()
    const label = newLabel.trim()
    if (label && onAddDateButton(label)) setNewLabel('')
  }

  return (
    <section className="card">
      <div className="qr">
        {lead.phone ? (
          <>
            <QRCodeSVG value={`tel:${dialNumber(lead.phone)}`} size={240} marginSize={2} />
            <a className="phone" href={`tel:${dialNumber(lead.phone)}`}>
              {dialNumber(lead.phone)}
            </a>
            <span className="qr-hint">Scan with your phone to dial</span>
          </>
        ) : (
          <div className="no-phone">No phone number</div>
        )}
      </div>

      <div className="details">
        <h2 className="name">{lead.name || 'Unnamed lead'}</h2>
        <dl>
          <dt>Email</dt>
          <dd>{lead.email ? <a href={`mailto:${lead.email}`}>{lead.email}</a> : '—'}</dd>
          <dt>Address</dt>
          <dd>{lead.address || '—'}</dd>
          <dt>Education</dt>
          <dd>{lead.education || '—'}</dd>
          <dt>Age</dt>
          <dd>{lead.age || '—'}</dd>
          <dt>Current status</dt>
          <dd>{lead.currentStatus || '—'}</dd>
          <dt>Lead status</dt>
          <dd>
            {lead.leadStatus ? (
              <span className="badge">{lead.leadStatus}</span>
            ) : (
              <span className="badge badge-muted">Not set</span>
            )}
          </dd>
          <dt>Coming date</dt>
          <dd>
            {lead.status ? (
              <span className="badge">Coming {lead.status}</span>
            ) : (
              <span className="badge badge-muted">Not set</span>
            )}
          </dd>
        </dl>

        <div className="dates">
          <div className="dates-head">
            <h3>When will they come?</h3>
            <button className="link" onClick={() => setEditMode((v) => !v)}>
              {editMode ? 'Done' : 'Edit buttons'}
            </button>
          </div>
          <div className="date-buttons">
            {dateButtons.map((label, i) => (
              <span key={label} className="date-wrap">
                <button
                  className={`date-btn ${lead.status === label ? 'active' : ''}`}
                  onClick={() => onSetStatus(label)}
                  title={i < 9 ? `Shortcut: ${i + 1}` : undefined}
                >
                  {label}
                </button>
                {editMode && (
                  <button
                    className="remove"
                    onClick={() => onRemoveDateButton(label)}
                    aria-label={`Remove ${label}`}
                  >
                    ×
                  </button>
                )}
              </span>
            ))}
            {lead.status && (
              <button className="date-btn clear" onClick={() => onSetStatus(lead.status)}>
                Clear
              </button>
            )}
          </div>
          <form className="add-date" onSubmit={submitNewLabel}>
            <input
              value={newLabel}
              onChange={(e) => setNewLabel(e.target.value)}
              placeholder="New button, e.g. 30 Nov"
            />
            <button className="btn" type="submit">
              + Add button
            </button>
          </form>
        </div>
      </div>
    </section>
  )
}
