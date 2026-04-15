/**
 * Pizza Hut Fiscal Calendar Parser
 * Parses the U.S. Period Calendars to map periods/weeks to actual dates
 */

// Hardcoded fiscal calendar for 2026 based on the Excel file
// PH weeks run Tuesday to Monday
const FISCAL_CALENDAR_2026 = {
  'P1': {
    name: 'Period 1',
    weeks: {
      'W1': { start: '2025-12-30', end: '2026-01-05', weekNum: 1 },
      'W2': { start: '2026-01-06', end: '2026-01-12', weekNum: 2 },
      'W3': { start: '2026-01-13', end: '2026-01-19', weekNum: 3 },
      'W4': { start: '2026-01-20', end: '2026-01-26', weekNum: 4 }
    }
  },
  'P2': {
    name: 'Period 2',
    weeks: {
      'W1': { start: '2026-01-27', end: '2026-02-02', weekNum: 5 },
      'W2': { start: '2026-02-03', end: '2026-02-09', weekNum: 6 },
      'W3': { start: '2026-02-10', end: '2026-02-16', weekNum: 7 },
      'W4': { start: '2026-02-17', end: '2026-02-23', weekNum: 8 }
    }
  },
  'P3': {
    name: 'Period 3',
    weeks: {
      'W1': { start: '2026-02-24', end: '2026-03-02', weekNum: 9 },
      'W2': { start: '2026-03-03', end: '2026-03-09', weekNum: 10 },
      'W3': { start: '2026-03-10', end: '2026-03-16', weekNum: 11 },
      'W4': { start: '2026-03-17', end: '2026-03-23', weekNum: 12 }
    }
  },
  'P4': {
    name: 'Period 4',
    weeks: {
      'W1': { start: '2026-03-24', end: '2026-03-30', weekNum: 13 },
      'W2': { start: '2026-03-31', end: '2026-04-06', weekNum: 14 },
      'W3': { start: '2026-04-07', end: '2026-04-13', weekNum: 15 },
      'W4': { start: '2026-04-14', end: '2026-04-20', weekNum: 16 }
    }
  },
  'P5': {
    name: 'Period 5',
    weeks: {
      'W1': { start: '2026-04-21', end: '2026-04-27', weekNum: 17 },
      'W2': { start: '2026-04-28', end: '2026-05-04', weekNum: 18 },
      'W3': { start: '2026-05-05', end: '2026-05-11', weekNum: 19 },
      'W4': { start: '2026-05-12', end: '2026-05-18', weekNum: 20 }
    }
  },
  'P6': {
    name: 'Period 6',
    weeks: {
      'W1': { start: '2026-05-19', end: '2026-05-25', weekNum: 21 },
      'W2': { start: '2026-05-26', end: '2026-06-01', weekNum: 22 },
      'W3': { start: '2026-06-02', end: '2026-06-08', weekNum: 23 },
      'W4': { start: '2026-06-09', end: '2026-06-15', weekNum: 24 }
    }
  },
  'P7': {
    name: 'Period 7',
    weeks: {
      'W1': { start: '2026-06-16', end: '2026-06-22', weekNum: 25 },
      'W2': { start: '2026-06-23', end: '2026-06-29', weekNum: 26 },
      'W3': { start: '2026-06-30', end: '2026-07-06', weekNum: 27 },
      'W4': { start: '2026-07-07', end: '2026-07-13', weekNum: 28 }
    }
  },
  'P8': {
    name: 'Period 8',
    weeks: {
      'W1': { start: '2026-07-14', end: '2026-07-20', weekNum: 29 },
      'W2': { start: '2026-07-21', end: '2026-07-27', weekNum: 30 },
      'W3': { start: '2026-07-28', end: '2026-08-03', weekNum: 31 },
      'W4': { start: '2026-08-04', end: '2026-08-10', weekNum: 32 }
    }
  },
  'P9': {
    name: 'Period 9',
    weeks: {
      'W1': { start: '2026-08-11', end: '2026-08-17', weekNum: 33 },
      'W2': { start: '2026-08-18', end: '2026-08-24', weekNum: 34 },
      'W3': { start: '2026-08-25', end: '2026-08-31', weekNum: 35 },
      'W4': { start: '2026-09-01', end: '2026-09-07', weekNum: 36 }
    }
  },
  'P10': {
    name: 'Period 10',
    weeks: {
      'W1': { start: '2026-09-08', end: '2026-09-14', weekNum: 37 },
      'W2': { start: '2026-09-15', end: '2026-09-21', weekNum: 38 },
      'W3': { start: '2026-09-22', end: '2026-09-28', weekNum: 39 },
      'W4': { start: '2026-09-29', end: '2026-10-05', weekNum: 40 }
    }
  },
  'P11': {
    name: 'Period 11',
    weeks: {
      'W1': { start: '2026-10-06', end: '2026-10-12', weekNum: 41 },
      'W2': { start: '2026-10-13', end: '2026-10-19', weekNum: 42 },
      'W3': { start: '2026-10-20', end: '2026-10-26', weekNum: 43 },
      'W4': { start: '2026-10-27', end: '2026-11-02', weekNum: 44 }
    }
  },
  'P12': {
    name: 'Period 12',
    weeks: {
      'W1': { start: '2026-11-03', end: '2026-11-09', weekNum: 45 },
      'W2': { start: '2026-11-10', end: '2026-11-16', weekNum: 46 },
      'W3': { start: '2026-11-17', end: '2026-11-23', weekNum: 47 },
      'W4': { start: '2026-11-24', end: '2026-11-30', weekNum: 48 }
    }
  },
  'P13': {
    name: 'Period 13',
    weeks: {
      'W1': { start: '2026-12-01', end: '2026-12-07', weekNum: 49 },
      'W2': { start: '2026-12-08', end: '2026-12-14', weekNum: 50 },
      'W3': { start: '2026-12-15', end: '2026-12-21', weekNum: 51 },
      'W4': { start: '2026-12-22', end: '2026-12-28', weekNum: 52 }
    }
  }
};

/**
 * Get the period and week for a given date
 */
function getPeriodForDate(dateStr) {
  for (const [period, data] of Object.entries(FISCAL_CALENDAR_2026)) {
    for (const [week, dates] of Object.entries(data.weeks)) {
      if (dateStr >= dates.start && dateStr <= dates.end) {
        return { period, week, periodWeek: period + week, ...dates };
      }
    }
  }
  return null;
}

/**
 * Get all dates in a specific period
 */
function getDatesInPeriod(periodStr) {
  const period = FISCAL_CALENDAR_2026[periodStr];
  if (!period) return [];
  
  const dates = [];
  for (const week of Object.values(period.weeks)) {
    let current = new Date(week.start);
    const end = new Date(week.end);
    while (current <= end) {
      dates.push(current.toISOString().split('T')[0]);
      current.setDate(current.getDate() + 1);
    }
  }
  return dates;
}

/**
 * Get all week keys for a period (e.g., ['P4W1', 'P4W2', 'P4W3', 'P4W4'])
 */
function getWeeksInPeriod(periodStr) {
  const period = FISCAL_CALENDAR_2026[periodStr];
  if (!period) return [];
  return Object.keys(period.weeks).map(w => periodStr + w);
}

/**
 * Get the week key (Tuesday-based) for a date
 */
function getWeekKeyForDate(dateStr) {
  const d = new Date(dateStr + 'T12:00:00Z');
  const day = d.getUTCDay();
  // Tuesday = 2, so days from Tuesday = (day + 5) % 7
  const daysFromTue = (day + 5) % 7;
  const tue = new Date(d);
  tue.setUTCDate(d.getUTCDate() - daysFromTue);
  return tue.toISOString().split('T')[0];
}

module.exports = {
  FISCAL_CALENDAR_2026,
  getPeriodForDate,
  getDatesInPeriod,
  getWeeksInPeriod,
  getWeekKeyForDate
};