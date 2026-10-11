
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
  status: 'مكتملة' | 'ناقصة';
}

export interface EmployeeSummary {
  id: string | number;
  name: string;
  shifts: ShiftReport[];
  totalPunches: number;
  completionRate: number;
  performanceScore: number;
  isEligible: boolean;
}

export interface TimePeriod {
  start: string; // HH:mm format
  end: string;   // HH:mm format
}

export interface Rules {
  REQUIRED_PUNCHES_PER_SHIFT: number;
  MAX_HOURS_BETWEEN_PUNCHES: number;
  TIME_PERIODS: {
    [key: string]: TimePeriod;
  };
}
