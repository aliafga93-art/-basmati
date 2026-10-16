import { Rules } from './types';

export const DEFAULT_5_PUNCH_RULES: Rules = {
  PUNCH_MODE: '5_punches',
  REQUIRED_PUNCHES_PER_SHIFT: 5,
  MAX_HOURS_BETWEEN_PUNCHES: 32,
  TIME_PERIODS: {
    "دخول صباحي": { start: "07:30", end: "09:30" },
    "بصمة وسطية": { start: "12:00", end: "15:59" },
    "بصمة مسائية": { start: "18:00", end: "21:59" },
    "بصمة ليلية": { start: "22:00", end: "03:00" },
    "بصمة خروج": { start: "07:30", end: "10:00" },
  }
};

export const DEFAULT_3_PUNCH_RULES: Rules = {
  PUNCH_MODE: '3_punches',
  REQUIRED_PUNCHES_PER_SHIFT: 3,
  MAX_HOURS_BETWEEN_PUNCHES: 32,
  TIME_PERIODS: {
    "دخول صباحي": { start: "07:30", end: "10:00" },
    "إثبات التواجد": { start: "12:00", end: "16:00" },
    "بصمة انصراف": { start: "14:30", end: "18:00" },
  }
};

export const ATTENDANCE_RULES: Rules = DEFAULT_3_PUNCH_RULES;

export const COLORS = {
  primary: '#09090b',    // Zinc 900
  secondary: '#71717a',  // Zinc 500
  accent: '#18181b',     // Zinc 900
  success: '#059669',    // Emerald 600
  danger: '#e11d48',     // Rose 600
  warning: '#d97706',    // Amber 600
  info: '#0284c7',       // Sky 600
  surface: '#ffffff',
  bg: '#fafafa'
};
