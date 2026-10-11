
import { Rules } from './types';

export const ATTENDANCE_RULES: Rules = {
  REQUIRED_PUNCHES_PER_SHIFT: 5,
  MAX_HOURS_BETWEEN_PUNCHES: 32, 
  TIME_PERIODS: {
    "دخول صباحي": { start: "07:30", end: "08:15" },
    "بصمة وسطية": { start: "12:00", end: "15:59" },
    "بصمة مسائية": { start: "18:00", end: "21:59" },
    "بصمة ليلية": { start: "22:00", end: "03:00" },
    "بصمة خروج": { start: "07:30", end: "10:00" }, 
  }
};

export const COLORS = {
  primary: '#3b82f6',    // Blue 500
  secondary: '#1e293b',  // Slate 800
  accent: '#8b5cf6',     // Violet 500
  success: '#10b981',    // Emerald 500
  danger: '#ef4444',     // Red 500
  warning: '#f59e0b',    // Amber 500
  info: '#0ea5e9',       // Sky 500
  surface: '#ffffff',
  bg: '#f8fafc'
};
