import { dialNumber } from './xlsxUtils'

export const ALL = '__all__'
export const UNSCHEDULED = '__unscheduled__'
export const DEFAULT_SCHEDULE_FILTER = { date: ALL, query: '' }

/** Returns [{ lead, index }] for the leads matching the schedule filter. */
export function filterLeads(leads, { date, query }) {
  const q = query.trim().toLowerCase()
  return leads
    .map((lead, index) => ({ lead, index }))
    .filter(({ lead }) => {
      if (date === ALL) return !!lead.status
      if (date === UNSCHEDULED) return !lead.status
      return lead.status === date
    })
    .filter(
      ({ lead }) =>
        !q ||
        lead.name.toLowerCase().includes(q) ||
        lead.phone.includes(q) ||
        dialNumber(lead.phone).includes(q) ||
        lead.address.toLowerCase().includes(q),
    )
}

export function filterLabel({ date, query }) {
  const base = date === ALL ? 'All scheduled' : date === UNSCHEDULED ? 'Not scheduled' : `Coming ${date}`
  return query.trim() ? `${base} · matching “${query.trim()}”` : base
}

/** File-name-safe slug of the filter, e.g. "coming-16-nov". */
export function filterSlug({ date }) {
  const base = date === ALL ? 'all-scheduled' : date === UNSCHEDULED ? 'not-scheduled' : `coming-${date}`
  return base.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '')
}
