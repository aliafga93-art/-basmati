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

  // Mode: 3 Punches (أمر تقليل البصمات إلى 3)
  if (rules.PUNCH_MODE === '3_punches') {
    // 1. بصمة الانصراف / الخروج
    if (periods["بصمة انصراف"] || periods["بصمة خروج"]) {
      const exitConfig = periods["بصمة انصراف"] || periods["بصمة خروج"];
      const exitStart = getRuleMinutes(exitConfig.start);
      const exitEnd = getRuleMinutes(exitConfig.end);
      if (isLast && (isNextDay || (t >= exitStart && t <= exitEnd) || t >= 840)) { // 14:00+
        return Object.keys(periods).find(k => k.includes('انصراف') || k.includes('خروج')) || "بصمة انصراف";
      }
    }

    // 2. دخول صباحي
    if (periods["دخول صباحي"]) {
      const mornStart = getRuleMinutes(periods["دخول صباحي"].start);
      const mornEnd = getRuleMinutes(periods["دخول صباحي"].end);
      if (isFirst && !isNextDay && t >= (mornStart - 60) && t <= mornEnd) {
        return "دخول صباحي";
      }
    }

    // 3. إثبات التواجد / وسطية
    if (periods["إثبات التواجد"] || periods["بصمة وسطية"]) {
      const midKey = periods["إثبات التواجد"] ? "إثبات التواجد" : "بصمة وسطية";
      const midStart = getRuleMinutes(periods[midKey].start);
      const midEnd = getRuleMinutes(periods[midKey].end);
      if (t >= midStart && t <= midEnd) {
        return midKey;
      }
    }

    // Fallback classification based on timing
    if (t < 660) return "دخول صباحي";
    if (t >= 660 && t < 900) return periods["إثبات التواجد"] ? "إثبات التواجد" : "بصمة وسطية";
    return periods["بصمة انصراف"] ? "بصمة انصراف" : "بصمة خروج";
  }

  // Mode: 5 Punches (الاعتيادي 5 بصمات)
  if (periods["بصمة خروج"]) {
    const exitStart = getRuleMinutes(periods["بصمة خروج"].start);
    const exitEnd = getRuleMinutes(periods["بصمة خروج"].end);
    if (isLast && (isNextDay || t >= exitStart) && t <= exitEnd) {
      return "بصمة خروج";
    }
  }

  if (periods["دخول صباحي"]) {
    const mornStart = getRuleMinutes(periods["دخول صباحي"].start);
    const mornEnd = getRuleMinutes(periods["دخول صباحي"].end);
    if (isFirst && !isNextDay && t >= mornStart && t <= mornEnd) {
      return "دخول صباحي";
    }
  }

  if (periods["بصمة ليلية"]) {
    const nightStart = getRuleMinutes(periods["بصمة ليلية"].start);
    const nightEnd = getRuleMinutes(periods["بصمة ليلية"].end);
    if (t >= nightStart || t <= nightEnd) {
      if (t > 1200 || t < 240) {
        return "بصمة ليلية";
      }
    }
  }

  if (periods["بصمة وسطية"]) {
    const midStart = getRuleMinutes(periods["بصمة وسطية"].start);
    const midEnd = getRuleMinutes(periods["بصمة وسطية"].end);
    if (t >= midStart && t <= midEnd) return "بصمة وسطية";
  }

  if (periods["بصمة مسائية"]) {
    const eveStart = getRuleMinutes(periods["بصمة مسائية"].start);
    const eveEnd = getRuleMinutes(periods["بصمة مسائية"].end);
    if (t >= eveStart && t <= eveEnd) return "بصمة مسائية";
  }

  // Fallback for morning exit on next day
  if (isLast && (isNextDay || t <= 600)) {
    return "بصمة خروج";
  }

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

  const employeesMap: Map<string, { name: string, punches: Date[] }> = new Map();
  const monthFrequencyMap: Map<string, number> = new Map();
  
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

        // Track month frequency (YYYY-MM)
        const monthKey = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
        monthFrequencyMap.set(monthKey, (monthFrequencyMap.get(monthKey) || 0) + 1);
      }
    });
  });

  // Find primary audited month in dataset (e.g. 2025-08 -> Month 8)
  let targetMonthKey = '';
  let maxCount = -1;
  monthFrequencyMap.forEach((count, key) => {
    if (count > maxCount) {
      maxCount = count;
      targetMonthKey = key;
    }
  });

  const targetMonthNum = targetMonthKey ? parseInt(targetMonthKey.split('-')[1], 10) : 0;
  const auditedMonthLabel = targetMonthNum ? `الشهر ${targetMonthNum}` : '';

  // Calculate unique dates in the target audited month
  const targetMonthDates = new Set<string>();
  employeesMap.forEach(empData => {
    empData.punches.forEach(p => {
      const pMonthKey = `${p.getFullYear()}-${String(p.getMonth() + 1).padStart(2, '0')}`;
      if (!targetMonthKey || pMonthKey === targetMonthKey) {
        targetMonthDates.add(p.toLocaleDateString('en-CA'));
      }
    });
  });

  const totalPeriodDays = Math.max(1, targetMonthDates.size);
  const expectedPeriodNames = Object.keys(rules.TIME_PERIODS);
  const allSummaries: EmployeeSummary[] = [];

  employeesMap.forEach((data, id) => {
    // Sort and deduplicate punches within 2 minutes of each other
    const rawSorted = data.punches.sort((a, b) => a.getTime() - b.getTime());
    const sortedPunches: Date[] = [];
    
    for (const p of rawSorted) {
      if (sortedPunches.length === 0) {
        sortedPunches.push(p);
      } else {
        const lastP = sortedPunches[sortedPunches.length - 1];
        const diffMinutes = Math.abs((p.getTime() - lastP.getTime()) / (1000 * 60));
        // If more than 2 minutes apart, accept as distinct punch
        if (diffMinutes >= 2) {
          sortedPunches.push(p);
        }
      }
    }

    const shiftsData: Date[][] = [];
    const empTargetMonthActiveDates = new Set<string>();
    
    sortedPunches.forEach(p => {
      const pMonthKey = `${p.getFullYear()}-${String(p.getMonth() + 1).padStart(2, '0')}`;
      if (!targetMonthKey || pMonthKey === targetMonthKey) {
        empTargetMonthActiveDates.add(p.toLocaleDateString('en-CA'));
      }
    });

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
      const shiftEnd = punchDates[punchDates.length - 1];
      const startDateStr = shiftStart.toLocaleDateString('en-CA');
      const startMonthKey = `${shiftStart.getFullYear()}-${String(shiftStart.getMonth() + 1).padStart(2, '0')}`;
      const endMonthKey = `${shiftEnd.getFullYear()}-${String(shiftEnd.getMonth() + 1).padStart(2, '0')}`;
      const exitDateStr = shiftEnd.toLocaleDateString('en-CA');

      // Check if this shift is a previous month exit punch:
      // Case 1: Shift started in a previous month (e.g. 2025-07-31) and ends on Day 1 or 2 of target month (2025-08-01)
      // Case 2: Shift is dated before targetMonthKey
      // Case 3: Shift starts on day 1 of target month, but only has early morning/exit punches (e.g. before 10:00 AM) and only 1 or 2 punches without morning entry
      const isPriorMonth = targetMonthKey && startMonthKey < targetMonthKey;
      const isDayOneOrphanExit = targetMonthKey && startMonthKey === targetMonthKey && shiftStart.getDate() === 1 && 
        punchDates.length <= 2 && shiftEnd.getHours() <= 10 && shiftStart.getHours() >= 5;

      const isPreviousMonthShift = Boolean(isPriorMonth || isDayOneOrphanExit);

      const details: PunchDetail[] = [];
      const processedPeriods = new Set<string>();

      punchDates.forEach((p, idx) => {
        const isFirst = idx === 0;
        const isLast = idx === punchDates.length - 1;
        const periodType = classifyPunch(p, isFirst, isLast, shiftStart, rules) || "غير مصنفة";
        
        let isDuplicate = processedPeriods.has(periodType) && periodType !== "غير مصنفة";
        if (periodType === "بصمة خروج" || periodType === "بصمة انصراف") isDuplicate = false;

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
      const missingPeriods = expectedPeriodNames.filter(pName => !uniqueValidPeriods.has(pName));
      const isComplete = uniqueValidPeriods.size >= rules.REQUIRED_PUNCHES_PER_SHIFT;

      if (isPreviousMonthShift) {
        // كما طلب المستخدم تماماً:
        // "لا تظهر اليوم الذي يسبقها مثلا 31/7 بل اكتب بصمة خروج لنوبة سابقة ولا تحتسبها ضمن النوبات الشهر 8"
        return {
          employee_id: id,
          employee_name: data.name,
          shift_date: 'بصمة خروج لنوبة سابقة',
          punches_details: details,
          punches_count: uniqueValidPeriods.size,
          status: 'نوبة سابقة',
          missing_periods: [], // معفاة تماماً ولا تحتسب نواقص
          is_previous_shift: true,
          previous_shift_label: 'بصمة خروج لنوبة سابقة',
          exit_date_str: exitDateStr
        };
      }

      return {
        employee_id: id,
        employee_name: data.name,
        shift_date: startDateStr,
        punches_details: details,
        punches_count: uniqueValidPeriods.size,
        status: isComplete ? 'مكتملة' : 'ناقصة',
        missing_periods: missingPeriods,
        is_previous_shift: false
      };
    });

    // فلترة النوبات الخاصة بالشهر المدقق فقط لاحتساب النسب والنواقص
    // "ولا تحتسبها ضمن النوبات الشهر 8"
    const currentMonthShifts = shiftsReports.filter(s => !s.is_previous_shift && s.punches_count > 0);
    
    // إذا لم تكن هناك أي نوبات للشهر أو كان للموظف فقط بصمة خروج لنوبة سابقة
    const totalValidUniquePunches = currentMonthShifts.reduce((sum, s) => sum + s.punches_count, 0);
    const completedShiftsCount = currentMonthShifts.filter(s => s.status === 'مكتملة').length;
    const missingShiftsCount = currentMonthShifts.filter(s => s.status === 'ناقصة').length;

    const completionRate = currentMonthShifts.length === 0 
      ? 100 
      : Math.round((completedShiftsCount / currentMonthShifts.length) * 100);

    const performanceScore = currentMonthShifts.length === 0 
      ? 100 
      : Math.min(100, Math.round((totalValidUniquePunches / (currentMonthShifts.length * rules.REQUIRED_PUNCHES_PER_SHIFT)) * 100));

    const isEligible = currentMonthShifts.length > 0 ? (completedShiftsCount === currentMonthShifts.length) : true;

    allSummaries.push({
      id: id,
      name: data.name,
      shifts: shiftsReports, // نحتفظ بها للعرض في سجل الموظف ولكنها موسومة بـ is_previous_shift
      totalPunches: totalValidUniquePunches,
      completionRate: completionRate,
      performanceScore: performanceScore,
      isEligible: isEligible,
      fingerprintDaysCount: empTargetMonthActiveDates.size,
      totalPeriodDays: totalPeriodDays,
      missingShiftsCount: missingShiftsCount,
      completedShiftsCount: completedShiftsCount,
      auditedMonthLabel: auditedMonthLabel
    });
  });

  return allSummaries.sort((a, b) => {
    const numA = parseInt(String(a.id).replace(/\D/g, '')) || 0;
    const numB = parseInt(String(b.id).replace(/\D/g, '')) || 0;
    return numA - numB;
  });
};
