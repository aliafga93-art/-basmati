export interface PunchDetail {
  time: string;
  period: string;
  is_duplicate: boolean;
  raw_datetime: Date;
}

export interface ShiftReport {
  employee_id: string | number;
  employee_name: string;
  shift_date: string;
  punches_details: PunchDetail[];
  punches_count: number;
  status: 'مكتملة' | 'ناقصة' | 'نوبة سابقة';
  missing_periods?: string[];
  is_previous_shift?: boolean;       // هل هي بصمة خروج لنوبة سابقة من الشهر السابق؟
  previous_shift_label?: string;     // "بصمة خروج لنوبة سابقة"
  exit_date_str?: string;            // تاريخ تسجيل بصمة الخروج (مثلاً 01/08)
}

export interface EmployeeSummary {
  id: string | number;
  name: string;
  shifts: ShiftReport[];
  totalPunches: number;
  completionRate: number;
  performanceScore: number;
  isEligible: boolean;
  fingerprintDaysCount: number; // عدد أيام البصمات الفعلية للشهر المدقق
  totalPeriodDays: number;       // إجمالي أيام الفترة المدققة
  missingShiftsCount: number;    // عدد النوبات الناقصة للشهر المدقق
  completedShiftsCount: number;  // عدد النوبات المكتملة للشهر المدقق
  auditedMonthLabel?: string;    // اسم أو رقم الشهر المدقق (مثلاً: شهر 8 / أغسطس)
}

export interface TimePeriod {
  start: string; // HH:mm format
  end: string;   // HH:mm format
}

export type PunchRuleMode = '3_punches' | '5_punches';

export interface Rules {
  PUNCH_MODE: PunchRuleMode;
  REQUIRED_PUNCHES_PER_SHIFT: number;
  MAX_HOURS_BETWEEN_PUNCHES: number;
  TIME_PERIODS: {
    [key: string]: TimePeriod;
  };
}
