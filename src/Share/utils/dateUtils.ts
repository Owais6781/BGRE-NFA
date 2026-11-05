/**
 * Safely parses various date formats used across the NFA system:
 * - DD/MM/YYYY or DD-MM-YYYY (e.g. from en-IN locale "04/11/2026", "19/11/2026")
 * - DD/MM/YYYY HH:mm:ss
 * - ISO 8601 strings (e.g. "2026-11-04T10:00:00Z" from SharePoint Created/Modified)
 * - Standard Date constructor strings
 */
export const parseDateSafe = (dateStr: string | null | undefined): Date | null => {
  if (!dateStr || typeof dateStr !== 'string') return null;
  const trimmed = dateStr.trim();
  if (!trimmed || trimmed === '—' || trimmed === '-') return null;

  // DD/MM/YYYY or DD-MM-YYYY (with optional time)
  const dmyMatch = trimmed.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{4})(?:[,\s]+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?/);
  if (dmyMatch) {
    const day = parseInt(dmyMatch[1], 10);
    const month = parseInt(dmyMatch[2], 10) - 1; // 0-indexed month
    const year = parseInt(dmyMatch[3], 10);
    const hours = dmyMatch[4] ? parseInt(dmyMatch[4], 10) : 0;
    const minutes = dmyMatch[5] ? parseInt(dmyMatch[5], 10) : 0;
    const seconds = dmyMatch[6] ? parseInt(dmyMatch[6], 10) : 0;
    const d = new Date(year, month, day, hours, minutes, seconds);
    if (!isNaN(d.getTime())) return d;
  }

  // Fallback to standard Date constructor
  const standardDate = new Date(trimmed);
  if (!isNaN(standardDate.getTime())) {
    return standardDate;
  }

  return null;
};

/**
 * Formats a date string into a friendly "MMM D, YYYY" string (e.g. "Nov 4, 2026").
 * Correctly parses day/month/year strings so the month is never confused with the day.
 */
export const formatDateFriendly = (dateStr: string | null | undefined): string => {
  if (!dateStr) return '—';
  const dateObj = parseDateSafe(dateStr);
  if (!dateObj) return dateStr;
  return dateObj.toLocaleDateString('en-US', {
    year: 'numeric',
    month: 'short',
    day: 'numeric'
  });
};
