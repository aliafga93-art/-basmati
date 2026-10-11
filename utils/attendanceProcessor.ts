
import { ShiftReport, PunchDetail, EmployeeSummary, Rules } from '../types';

const getTimeInMinutes = (date: Date) => date.getHours() * 60 + date.getMinutes();

const getRuleMinutes = (timeStr: string) => {
  const [h, m] = timeStr.split(':').map(Number);
  return h * 60 + m;
};

export const classifyPunch = (
  punchDate: Date,
  isFirst: boolean,
  isLast: boolean,
  shiftStartDate: Date,
  rules: Rules
): string | null => {
  const t = getTimeInMinutes(punchDate);
  const periods = rules.TIME_PERIODS;
  const isNextDay = punchDate.getDate() !== shiftStartDate.getDate();

  const exitStart = getRuleMinutes(periods["بصمة خروج"].start);
  const exitEnd = getRuleMinutes(periods["بصمة خروج"].end);
  if (isLast && (isNextDay || t >= exitStart) && t <= exitEnd) {
    return "بصمة خروج";
  }

  const mornStart = getRuleMinutes(periods["دخول صباحي"].start);
  const mornEnd = getRuleMinutes(periods["دخول صباحي"].end);
  if (isFirst && !isNextDay && t >= mornStart && t <= mornEnd) {
    return "دخول صباحي";
  }

  const nightStart = getRuleMinutes(periods["بصمة ليلية"].start);
  const nightEnd = getRuleMinutes(periods["بصمة ليلية"].end);
  if (t >= nightStart || t <= nightEnd) {
    if (t > mornEnd || t < 240) {
      return "بصمة ليلية";
    }
  }

  const midStart = getRuleMinutes(periods["بصمة وسطية"].start);
  const midEnd = getRuleMinutes(periods["بصمة وسطية"].end);
  if (t >= midStart && t <= midEnd) return "بصمة وسطية";

  const eveStart = getRuleMinutes(periods["بصمة مسائية"].start);
  const eveEnd = getRuleMinutes(periods["بصمة مسائية"].end);
  if (t >= eveStart && t <= eveEnd) return "بصمة مسائية";

  return null;
};

const getVal = (row: any, keys: string[]) => {
  const rowKeys = Object.keys(row);
  for (const k of keys) {
    const foundKey = rowKeys.find(rk => 
      rk.replace(/[-_ ]/g, '').toLowerCase() === k.replace(/[-_ ]/g, '').toLowerCase() ||
      rk.toLowerCase().includes(k.toLowerCase())
    );
    if (foundKey) return row[foundKey];
  }
  return null;
};

export const processExcelData = (rawData: any[], rules: Rules): EmployeeSummary[] => {
  if (!rawData || rawData.length === 0) return [];

  // التجميع حصرياً حسب ACNo لضمان بطاقة واحدة لكل موظف
  const employeesMap: Map<string, { name: string, punches: Date[] }> = new Map();
  
  rawData.forEach(row => {
    const idVal = getVal(row, ['ACNo', 'AC-No', 'EnrollID', 'No', 'رقم', 'معرف', 'ID']);
    if (idVal === null || idVal === undefined) return;
    const id = String(idVal).trim();
    
    const nameVal = getVal(row, ['Name', 'الاسم', 'اسم']) || `موظف (${id})`;
    const name = String(nameVal).trim();

    const dateVal = getVal(row, ['Date', 'التاريخ', 'تاريخ']);
    const timePart = String(getVal(row, ['Time', 'الوقت', 'وقت']) || "");
    
    if (!timePart || !timePart.trim()) return;

    let dt: Date;
    if (typeof dateVal === 'number') {
      dt = new Date((dateVal - 25569) * 86400 * 1000);
    } else {
      dt = new Date(dateVal);
    }

    if (isNaN(dt.getTime())) return;

    const times = timePart.split(/[\s,]+/).filter(t => t.includes(':'));
    
    if (!employeesMap.has(id)) {
      employeesMap.set(id, { name, punches: [] });
    }
    
    const empData = employeesMap.get(id)!;
    if (name && empData.name.startsWith('موظف (')) empData.name = name;

    times.forEach(tStr => {
      const d = new Date(dt);
      const parts = tStr.split(':').map(Number);
      if (parts.length >= 2) {
        d.setHours(parts[0] || 0, parts[1] || 0, 0, 0);
        empData.punches.push(d);
      }
    });
  });

  const allSummaries: EmployeeSummary[] = [];

  employeesMap.forEach((data, id) => {
    const sortedPunches = data.punches.sort((a, b) => a.getTime() - b.getTime());
    const shiftsData: Date[][] = [];
    
    if (sortedPunches.length > 0) {
      let currentShift: Date[] = [sortedPunches[0]];
      for (let i = 1; i < sortedPunches.length; i++) {
        const diffHours = (sortedPunches[i].getTime() - sortedPunches[i-1].getTime()) / (1000 * 3600);
        if (diffHours > rules.MAX_HOURS_BETWEEN_PUNCHES) {
          shiftsData.push(currentShift);
          currentShift = [];
        }
        currentShift.push(sortedPunches[i]);
      }
      if (currentShift.length > 0) shiftsData.push(currentShift);
    }

    const shiftsReports: ShiftReport[] = shiftsData.map(punchDates => {
      const shiftStart = punchDates[0];
      const startDateStr = shiftStart.toLocaleDateString('en-CA');
      const details: PunchDetail[] = [];
      const processedPeriods = new Set<string>();

      punchDates.forEach((p, idx) => {
        const isFirst = idx === 0;
        const isLast = idx === punchDates.length - 1;
        const periodType = classifyPunch(p, isFirst, isLast, shiftStart, rules) || "غير مصنفة";
        
        let isDuplicate = processedPeriods.has(periodType) && periodType !== "غير مصنفة";
        if (periodType === "بصمة خروج") isDuplicate = false;

        details.push({
          time: p.toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit' }),
          period: periodType,
          is_duplicate: isDuplicate,
          raw_datetime: p
        });

        if (!isDuplicate && periodType !== "غير مصنفة") {
          processedPeriods.add(periodType);
        }
      });

      const uniqueValidPeriods = new Set(details.filter(d => !d.is_duplicate && d.period !== "غير مصنفة").map(d => d.period));
      
      return {
        employee_id: id,
        employee_name: data.name,
        shift_date: startDateStr,
        punches_details: details,
        punches_count: uniqueValidPeriods.size,
        status: uniqueValidPeriods.size >= rules.REQUIRED_PUNCHES_PER_SHIFT ? 'مكتملة' : 'ناقصة'
      };
    });

    const activeShifts = shiftsReports.filter(s => s.punches_count > 0);
    if (activeShifts.length === 0) return;

    const totalValidUniquePunches = activeShifts.reduce((sum, s) => sum + s.punches_count, 0);
    const completedShiftsCount = activeShifts.filter(s => s.status === 'مكتملة').length;
    
    allSummaries.push({
      id: id,
      name: data.name,
      shifts: activeShifts,
      totalPunches: totalValidUniquePunches,
      completionRate: Math.round((completedShiftsCount / activeShifts.length) * 100),
      performanceScore: Math.round((totalValidUniquePunches / (activeShifts.length * rules.REQUIRED_PUNCHES_PER_SHIFT)) * 100),
      isEligible: completedShiftsCount === activeShifts.length
    });
  });

  return allSummaries.sort((a, b) => {
    const numA = parseInt(String(a.id).replace(/\D/g, '')) || 0;
    const numB = parseInt(String(b.id).replace(/\D/g, '')) || 0;
    return numA - numB;
  });
};
