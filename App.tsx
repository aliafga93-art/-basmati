import React, { useState, useMemo, useRef, useEffect } from 'react';
import * as XLSX from 'xlsx';
import { 
  FileUp, Search, Users, LayoutDashboard, Fingerprint, 
  RefreshCcw, X, ShieldCheck, CheckCircle2, 
  AlertTriangle, FileSpreadsheet, Printer, 
  Calendar, ArrowUpDown, Clock, Moon, Sun, Check, ArrowRight,
  TrendingUp, Download, Eye, Sparkles, UserCheck, 
  UserX, BarChart3, Award, ExternalLink, HelpCircle,
  FileText, ChevronLeft, ChevronRight, Filter, Info,
  SlidersHorizontal, CheckCircle, AlertCircle
} from 'lucide-react';

import { EmployeeSummary, ShiftReport, Rules, PunchRuleMode } from './types';
import { DEFAULT_3_PUNCH_RULES, DEFAULT_5_PUNCH_RULES } from './constants';
import { processExcelData } from './utils/attendanceProcessor';
import { generateSampleDataset } from './utils/demoData';

export default function App() {
  // Mode selection: 3 punches vs 5 punches
  const [punchMode, setPunchMode] = useState<PunchRuleMode>('3_punches');
  
  // Current active rules
  const activeRules: Rules = useMemo(() => {
    return punchMode === '3_punches' ? DEFAULT_3_PUNCH_RULES : DEFAULT_5_PUNCH_RULES;
  }, [punchMode]);

  // Raw rows storage to allow re-processing when rule changes
  const [rawDataset, setRawDataset] = useState<any[] | null>(null);
  
  // Audited Employees Data
  const [employees, setEmployees] = useState<EmployeeSummary[]>(() => {
    return generateSampleDataset(DEFAULT_3_PUNCH_RULES);
  });

  // Active view tab: Default to high-density audit matrix as requested by user
  const [activeTab, setActiveTab] = useState<'matrix' | 'dashboard' | 'deficiencies' | 'rule_comparison'>('matrix');

  // Drawer / Inspection state
  const [selectedEmployee, setSelectedEmployee] = useState<EmployeeSummary | null>(null);

  // Search & Filter state
  const [searchQuery, setSearchQuery] = useState('');
  const [statusFilter, setStatusFilter] = useState<'all' | 'compliant' | 'deficient' | 'previous_exits'>('all');
  const [sortField, setSortField] = useState<'id' | 'name' | 'days' | 'compliance' | 'completed'>('id');
  const [sortAsc, setSortAsc] = useState(true);

  // Dark mode
  const [isDarkMode, setIsDarkMode] = useState(false);

  // File input ref
  const fileInputRef = useRef<HTMLInputElement>(null);
  const searchInputRef = useRef<HTMLInputElement>(null);

  // Toggle dark mode class on document
  useEffect(() => {
    if (isDarkMode) {
      document.documentElement.classList.add('dark');
    } else {
      document.documentElement.classList.remove('dark');
    }
  }, [isDarkMode]);

  // Keyboard shortcut for search & closing drawer
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if ((e.metaKey || e.ctrlKey) && e.key === 'k') {
        e.preventDefault();
        searchInputRef.current?.focus();
      }
      if (e.key === 'Escape') {
        setSelectedEmployee(null);
      }
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, []);

  // Handle re-processing whenever punchMode changes
  const handlePunchModeChange = (newMode: PunchRuleMode) => {
    setPunchMode(newMode);
    const newRules = newMode === '3_punches' ? DEFAULT_3_PUNCH_RULES : DEFAULT_5_PUNCH_RULES;
    
    if (rawDataset && rawDataset.length > 0) {
      const reprocessed = processExcelData(rawDataset, newRules);
      setEmployees(reprocessed);
      if (selectedEmployee) {
        const updated = reprocessed.find(e => String(e.id) === String(selectedEmployee.id)) || null;
        setSelectedEmployee(updated);
      }
    } else {
      const reprocessed = generateSampleDataset(newRules);
      setEmployees(reprocessed);
      if (selectedEmployee) {
        const updated = reprocessed.find(e => String(e.id) === String(selectedEmployee.id)) || null;
        setSelectedEmployee(updated);
      }
    }
  };

  // Handle Excel File Upload
  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bstr = evt.target?.result;
        const wb = XLSX.read(bstr, { type: 'binary', cellDates: true });
        const wsName = wb.SheetNames[0];
        const ws = wb.Sheets[wsName];
        const jsonData = XLSX.utils.sheet_to_json(ws);

        if (jsonData.length === 0) {
          alert('الملف فارغ أو لا يحتوي على صفوف بيانات صالحة.');
          return;
        }

        setRawDataset(jsonData);
        const processed = processExcelData(jsonData, activeRules);
        setEmployees(processed);
        setSelectedEmployee(null);
      } catch (err) {
        console.error('Error parsing excel:', err);
        alert('حدث خطأ أثناء قراءة ملف الإكسل. يرجى التأكد من الصيغة وتسميات الأعمدة.');
      }
    };
    reader.readAsBinaryString(file);
    e.target.value = '';
  };

  // Reset to Demo Dataset
  const handleLoadDemo = () => {
    setRawDataset(null);
    setEmployees(generateSampleDataset(activeRules));
    setSelectedEmployee(null);
  };

  // Calculate high-level KPIs & statistics
  const kpis = useMemo(() => {
    const totalEmp = employees.length;
    if (totalEmp === 0) {
      return { 
        totalEmp: 0, 
        compliantEmp: 0, 
        deficientEmp: 0,
        totalDays: 0, 
        avgCompliance: 0, 
        totalDeficiencies: 0,
        totalShifts: 0,
        completedShifts: 0,
        exemptExitsCount: 0,
        empWithExemptExits: 0
      };
    }

    const compliantEmp = employees.filter(e => e.isEligible).length;
    const deficientEmp = totalEmp - compliantEmp;
    const maxPeriodDays = Math.max(...employees.map(e => e.fingerprintDaysCount || 0), 1);
    const avgCompliance = Math.round(employees.reduce((acc, e) => acc + e.completionRate, 0) / totalEmp);
    const totalDeficiencies = employees.reduce((acc, e) => acc + e.missingShiftsCount, 0);

    // Count shifts (excluding previous shift exits)
    let totalShifts = 0;
    let completedShifts = 0;
    let exemptExitsCount = 0;
    let empWithExemptExits = 0;

    employees.forEach(emp => {
      let hasPreviousExit = false;
      emp.shifts.forEach(s => {
        if (s.is_previous_shift) {
          exemptExitsCount += 1;
          hasPreviousExit = true;
        } else {
          totalShifts += 1;
          if (s.status === 'مكتملة') completedShifts += 1;
        }
      });
      if (hasPreviousExit) empWithExemptExits += 1;
    });

    return {
      totalEmp,
      compliantEmp,
      deficientEmp,
      totalDays: maxPeriodDays,
      avgCompliance,
      totalDeficiencies,
      totalShifts,
      completedShifts,
      exemptExitsCount,
      empWithExemptExits
    };
  }, [employees]);

  // Comparison data for 3 vs 5 punches
  const comparisonData = useMemo(() => {
    const dataFor3 = rawDataset ? processExcelData(rawDataset, DEFAULT_3_PUNCH_RULES) : generateSampleDataset(DEFAULT_3_PUNCH_RULES);
    const dataFor5 = rawDataset ? processExcelData(rawDataset, DEFAULT_5_PUNCH_RULES) : generateSampleDataset(DEFAULT_5_PUNCH_RULES);

    const comp3Count = dataFor3.filter(e => e.isEligible).length;
    const comp5Count = dataFor5.filter(e => e.isEligible).length;
    const avg3 = Math.round(dataFor3.reduce((acc, e) => acc + e.completionRate, 0) / (dataFor3.length || 1));
    const avg5 = Math.round(dataFor5.reduce((acc, e) => acc + e.completionRate, 0) / (dataFor5.length || 1));

    const benefits = dataFor3.map(emp3 => {
      const emp5 = dataFor5.find(e => String(e.id) === String(emp3.id));
      const rate3 = emp3.completionRate;
      const rate5 = emp5 ? emp5.completionRate : 0;
      const diff = rate3 - rate5;
      const missing3 = emp3.missingShiftsCount;
      const missing5 = emp5 ? emp5.missingShiftsCount : 0;
      return {
        id: emp3.id,
        name: emp3.name,
        days: emp3.fingerprintDaysCount,
        rate3,
        rate5,
        diff,
        savedShifts: Math.max(0, missing5 - missing3),
        benefited: diff > 0 || (missing5 > missing3)
      };
    }).filter(b => b.benefited);

    return {
      comp3Count,
      comp5Count,
      avg3,
      avg5,
      benefits
    };
  }, [rawDataset]);

  // Deficiencies flat list across all employees (excluding previous month exit shifts)
  const allDeficiencies = useMemo(() => {
    const list: {
      empId: string | number;
      empName: string;
      shiftDate: string;
      recordedCount: number;
      missingPeriods: string[];
      punchesText: string;
    }[] = [];

    employees.forEach(emp => {
      emp.shifts.forEach(shift => {
        // نوبة سابقة معفاة ولا تعتبر نقصاً في الشهر الحالي
        if (shift.status === 'ناقصة' && !shift.is_previous_shift) {
          const punchesText = shift.punches_details.map(p => `${p.time} (${p.period})`).join(' ، ') || 'لا توجد بصمات';
          list.push({
            empId: emp.id,
            empName: emp.name,
            shiftDate: shift.shift_date,
            recordedCount: shift.punches_count,
            missingPeriods: shift.missing_periods || ['بصمة غير مكتملة'],
            punchesText
          });
        }
      });
    });

    return list;
  }, [employees]);

  // List of employees with previous shift exits detected (for quick audit verification)
  const previousShiftExits = useMemo(() => {
    const list: { emp: EmployeeSummary; shift: ShiftReport }[] = [];
    employees.forEach(emp => {
      emp.shifts.forEach(s => {
        if (s.is_previous_shift) {
          list.push({ emp, shift: s });
        }
      });
    });
    return list;
  }, [employees]);

  // Top compliant employees (100% adherence)
  const topCompliant = useMemo(() => {
    return employees.filter(e => e.isEligible).slice(0, 6);
  }, [employees]);

  // Priority attention employees (lowest compliance)
  const priorityAttention = useMemo(() => {
    return employees.filter(e => !e.isEligible).sort((a, b) => a.completionRate - b.completionRate).slice(0, 6);
  }, [employees]);

  // Filtered and Sorted employees
  const filteredEmployees = useMemo(() => {
    return employees
      .filter(emp => {
        const matchesSearch = 
          String(emp.id).toLowerCase().includes(searchQuery.toLowerCase()) ||
          emp.name.toLowerCase().includes(searchQuery.toLowerCase());
        
        if (!matchesSearch) return false;

        if (statusFilter === 'compliant') return emp.isEligible;
        if (statusFilter === 'deficient') return !emp.isEligible;
        if (statusFilter === 'previous_exits') {
          return emp.shifts.some(s => s.is_previous_shift);
        }
        return true;
      })
      .sort((a, b) => {
        let valA: any = a.id;
        let valB: any = b.id;

        if (sortField === 'id') {
          valA = parseInt(String(a.id).replace(/\D/g, '')) || 0;
          valB = parseInt(String(b.id).replace(/\D/g, '')) || 0;
        } else if (sortField === 'name') {
          valA = a.name;
          valB = b.name;
        } else if (sortField === 'days') {
          valA = a.fingerprintDaysCount;
          valB = b.fingerprintDaysCount;
        } else if (sortField === 'compliance') {
          valA = a.completionRate;
          valB = b.completionRate;
        } else if (sortField === 'completed') {
          valA = a.completedShiftsCount;
          valB = b.completedShiftsCount;
        }

        if (valA < valB) return sortAsc ? -1 : 1;
        if (valA > valB) return sortAsc ? 1 : -1;
        return 0;
      });
  }, [employees, searchQuery, statusFilter, sortField, sortAsc]);

  // Toggle sort field
  const handleSort = (field: 'id' | 'name' | 'days' | 'compliance' | 'completed') => {
    if (sortField === field) {
      setSortAsc(!sortAsc);
    } else {
      setSortField(field);
      setSortAsc(true);
    }
  };

  // Export multi-sheet Excel using official terms (نوبة / نوبات)
  const handleExportExcel = () => {
    try {
      const wb = XLSX.utils.book_new();

      // Sheet 1: Summary Matrix
      const summaryRows = employees.map((emp, index) => {
        const auditedShifts = emp.shifts.filter(s => !s.is_previous_shift);
        const previousExitsCount = emp.shifts.filter(s => s.is_previous_shift).length;
        return {
          'ت': index + 1,
          'الرقم الوظيفي': emp.id,
          'اسم الموظف': emp.name,
          'أيام البصمات المقيدة': emp.fingerprintDaysCount,
          'إجمالي أيام الفترة': emp.totalPeriodDays,
          'إجمالي نوبات الشهر': auditedShifts.length,
          'النوبات المكتملة': emp.completedShiftsCount,
          'النوبات الناقصة': emp.missingShiftsCount,
          'بصمات خروج نوبة سابقة (معفاة)': previousExitsCount,
          'معدل الالتزام': `${emp.completionRate}%`,
          'حالة التدقيق': emp.isEligible ? 'مطابق تماماً' : 'توجد نواقص',
          'القاعدة المعتمدة': punchMode === '3_punches' ? 'نظام 3 بصمات (أمر مخفض)' : 'نظام 5 بصمات (اعتيادي)'
        };
      });
      const wsSummary = XLSX.utils.json_to_sheet(summaryRows);
      XLSX.utils.book_append_sheet(wb, wsSummary, 'كشف التدقيق العام');

      // Sheet 2: Deficiencies Sheet
      const defRows = allDeficiencies.map((d, index) => ({
        'ت': index + 1,
        'الرقم الوظيفي': d.empId,
        'اسم الموظف': d.empName,
        'تاريخ النوبة': d.shiftDate,
        'عدد البصمات المسجلة': d.recordedCount,
        'البصمات الناقصة المطلوبة': d.missingPeriods.join(' - '),
        'تفاصيل البصمات المسجلة': d.punchesText
      }));
      const wsDef = XLSX.utils.json_to_sheet(defRows);
      XLSX.utils.book_append_sheet(wb, wsDef, 'محضر النواقص التفصيلي');

      // Sheet 3: Previous Shift Exits
      const prevRows = previousShiftExits.map((item, index) => ({
        'ت': index + 1,
        'الرقم الوظيفي': item.emp.id,
        'اسم الموظف': item.emp.name,
        'تاريخ بصمة الخروج': item.shift.exit_date_str || 'أول الشهر',
        'توقيت البصمة': item.shift.punches_details.map(p => p.time).join(' ، '),
        'الوصف الإداري': 'بصمة خروج لنوبة سابقة - مستبعدة من احتساب نوبات ونواقص الشهر الحالي'
      }));
      const wsPrev = XLSX.utils.json_to_sheet(prevRows);
      XLSX.utils.book_append_sheet(wb, wsPrev, 'بصمات خروج نوبات سابقة');

      const fileName = `تدقيق_نوبات_البصمات_${punchMode === '3_punches' ? '3بصمات' : '5بصمات'}_${new Date().toISOString().slice(0, 10)}.xlsx`;
      XLSX.writeFile(wb, fileName);
    } catch (err) {
      console.error('Failed to export:', err);
      alert('تعذر تصدير الملف، يرجى المحاولة مرة أخرى.');
    }
  };

  // Print Audit Sheet
  const handlePrint = () => {
    window.print();
  };

  return (
    <div className="flex flex-col h-screen w-screen bg-[#fafafa] dark:bg-[#09090b] text-zinc-900 dark:text-zinc-100 font-sans text-xs antialiased overflow-hidden select-none">
      
      {/* Hidden File Input */}
      <input 
        type="file" 
        ref={fileInputRef} 
        onChange={handleFileUpload} 
        accept=".xlsx, .xls, .csv" 
        className="hidden" 
      />

      {/* ── TOP EXECUTIVE APP BAR (Stripe / Linear Precision) ── */}
      <header className="h-12 border-b border-zinc-200 dark:border-zinc-800 bg-white/95 dark:bg-zinc-900/95 backdrop-blur-md px-4 flex items-center justify-between shrink-0 z-30">
        
        {/* Brand & System Context */}
        <div className="flex items-center gap-3">
          <div className="w-7 h-7 rounded-md bg-zinc-900 dark:bg-zinc-100 flex items-center justify-center text-white dark:text-zinc-950 shadow-xs">
            <Fingerprint className="w-4 h-4 stroke-[2.2]" />
          </div>
          <div className="flex items-center gap-2">
            <span className="font-extrabold text-[13px] tracking-tight text-zinc-900 dark:text-zinc-50">
              نظام تدقيق النوبات
            </span>
            <span className="text-zinc-300 dark:text-zinc-700">/</span>
            <span className="text-[11px] font-medium text-zinc-500 dark:text-zinc-400">
              رقابة البصمات الإدارية
            </span>
          </div>
        </div>

        {/* Center: DYNAMIC PUNCH RULE TOGGLE (3 vs 5 punches) */}
        <div className="flex items-center p-0.5 bg-zinc-100 dark:bg-zinc-800 rounded-lg border border-zinc-200/80 dark:border-zinc-700/80 text-[11px]">
          <button
            onClick={() => handlePunchModeChange('3_punches')}
            className={`px-3 py-1 rounded-md font-bold transition-all flex items-center gap-1.5 ${
              punchMode === '3_punches'
                ? 'bg-white dark:bg-zinc-900 text-zinc-900 dark:text-zinc-100 shadow-xs border border-zinc-200/50 dark:border-zinc-700'
                : 'text-zinc-500 dark:text-zinc-400 hover:text-zinc-900 dark:hover:text-zinc-200'
            }`}
            title="تطبيق أمر تقليل البصمات إلى 3 أوقات (دخول، إثبات تواجد، انصراف)"
          >
            {punchMode === '3_punches' && <Check className="w-3 h-3 text-emerald-600 dark:text-emerald-400 stroke-[3]" />}
            <span>نظام 3 بصمات (أمر مخفض)</span>
          </button>

          <button
            onClick={() => handlePunchModeChange('5_punches')}
            className={`px-3 py-1 rounded-md font-bold transition-all flex items-center gap-1.5 ${
              punchMode === '5_punches'
                ? 'bg-white dark:bg-zinc-900 text-zinc-900 dark:text-zinc-100 shadow-xs border border-zinc-200/50 dark:border-zinc-700'
                : 'text-zinc-500 dark:text-zinc-400 hover:text-zinc-900 dark:hover:text-zinc-200'
            }`}
            title="تطبيق قاعدة 5 بصمات الاعتيادية لكل نوبة كاملة"
          >
            {punchMode === '5_punches' && <Check className="w-3 h-3 text-emerald-600 dark:text-emerald-400 stroke-[3]" />}
            <span>نظام 5 بصمات (الاعتيادي)</span>
          </button>
        </div>

        {/* Left Utilities (RTL: Actions on left) */}
        <div className="flex items-center gap-2">
          {/* Upload Button */}
          <button
            onClick={() => fileInputRef.current?.click()}
            className="h-7.5 px-3 rounded-md bg-zinc-900 hover:bg-zinc-800 dark:bg-zinc-100 dark:hover:bg-white text-white dark:text-zinc-950 font-bold flex items-center gap-1.5 transition-colors shadow-xs"
            title="استيراد كشف بصمات إكسل جديد"
          >
            <FileUp className="w-3.5 h-3.5" />
            <span>استيراد Excel</span>
          </button>

          {/* Export Excel */}
          <button
            onClick={handleExportExcel}
            className="h-7.5 px-2.5 rounded-md border border-zinc-200 dark:border-zinc-700 bg-white hover:bg-zinc-50 dark:bg-zinc-800 dark:hover:bg-zinc-700 text-zinc-800 dark:text-zinc-200 font-semibold flex items-center gap-1.5 transition-colors"
            title="تصدير مصنف إكسل رسمي متعدد الصفحات"
          >
            <Download className="w-3.5 h-3.5 text-zinc-500" />
            <span>تصدير</span>
          </button>

          {/* Print */}
          <button
            onClick={handlePrint}
            className="h-7.5 w-7.5 rounded-md border border-zinc-200 dark:border-zinc-700 bg-white hover:bg-zinc-50 dark:bg-zinc-800 dark:hover:bg-zinc-700 text-zinc-700 dark:text-zinc-300 flex items-center justify-center transition-colors"
            title="طباعة محضر التدقيق الرسمي"
          >
            <Printer className="w-3.5 h-3.5" />
          </button>

          {/* Reload Demo */}
          <button
            onClick={handleLoadDemo}
            className="h-7.5 w-7.5 rounded-md border border-zinc-200 dark:border-zinc-700 bg-white hover:bg-zinc-50 dark:bg-zinc-800 dark:hover:bg-zinc-700 text-zinc-700 dark:text-zinc-300 flex items-center justify-center transition-colors"
            title="إعادة تحميل السجلات النموذجية"
          >
            <RefreshCcw className="w-3.5 h-3.5" />
          </button>

          {/* Dark Mode */}
          <button
            onClick={() => setIsDarkMode(!isDarkMode)}
            className="h-7.5 w-7.5 rounded-md border border-zinc-200 dark:border-zinc-700 bg-white hover:bg-zinc-50 dark:bg-zinc-800 dark:hover:bg-zinc-700 text-zinc-700 dark:text-zinc-300 flex items-center justify-center transition-colors"
            title="تبديل المظهر النهاري/الليلي"
          >
            {isDarkMode ? <Sun className="w-3.5 h-3.5" /> : <Moon className="w-3.5 h-3.5" />}
          </button>
        </div>
      </header>

      {/* ── SECONDARY METRIC RIBBON & SECTION TABS ── */}
      <div className="bg-white dark:bg-zinc-900 border-b border-zinc-200 dark:border-zinc-800 px-4 py-1.5 shrink-0">
        <div className="flex flex-wrap items-center justify-between gap-3">
          
          {/* Navigation Tabs */}
          <div className="flex items-center gap-1 p-0.5 bg-zinc-100 dark:bg-zinc-800/80 rounded-md border border-zinc-200/60 dark:border-zinc-700/60">
            <button
              onClick={() => setActiveTab('matrix')}
              className={`px-3 py-1 rounded text-xs font-bold transition-all flex items-center gap-1.5 ${
                activeTab === 'matrix'
                  ? 'bg-white dark:bg-zinc-900 text-zinc-900 dark:text-zinc-100 shadow-xs'
                  : 'text-zinc-500 dark:text-zinc-400 hover:text-zinc-900 dark:hover:text-zinc-200'
              }`}
            >
              <LayoutDashboard className="w-3.5 h-3.5 text-zinc-600 dark:text-zinc-400" />
              <span>جدول تدقيق النوبات المدمج</span>
            </button>

            <button
              onClick={() => setActiveTab('dashboard')}
              className={`px-3 py-1 rounded text-xs font-bold transition-all flex items-center gap-1.5 ${
                activeTab === 'dashboard'
                  ? 'bg-white dark:bg-zinc-900 text-zinc-900 dark:text-zinc-100 shadow-xs'
                  : 'text-zinc-500 dark:text-zinc-400 hover:text-zinc-900 dark:hover:text-zinc-200'
              }`}
            >
              <BarChart3 className="w-3.5 h-3.5 text-zinc-600 dark:text-zinc-400" />
              <span>لوحة المؤشرات والتحليل</span>
            </button>

            <button
              onClick={() => setActiveTab('deficiencies')}
              className={`px-3 py-1 rounded text-xs font-bold transition-all flex items-center gap-1.5 ${
                activeTab === 'deficiencies'
                  ? 'bg-white dark:bg-zinc-900 text-zinc-900 dark:text-zinc-100 shadow-xs'
                  : 'text-zinc-500 dark:text-zinc-400 hover:text-zinc-900 dark:hover:text-zinc-200'
              }`}
            >
              <AlertTriangle className="w-3.5 h-3.5 text-amber-500" />
              <span>محضر النواقص التفصيلي</span>
              {kpis.totalDeficiencies > 0 && (
                <span className="px-1.5 py-0.2 rounded text-[10px] bg-amber-100 dark:bg-amber-900/50 text-amber-800 dark:text-amber-300 font-mono font-bold">
                  {kpis.totalDeficiencies}
                </span>
              )}
            </button>

            <button
              onClick={() => setActiveTab('rule_comparison')}
              className={`px-3 py-1 rounded text-xs font-bold transition-all flex items-center gap-1.5 ${
                activeTab === 'rule_comparison'
                  ? 'bg-white dark:bg-zinc-900 text-zinc-900 dark:text-zinc-100 shadow-xs'
                  : 'text-zinc-500 dark:text-zinc-400 hover:text-zinc-900 dark:hover:text-zinc-200'
              }`}
            >
              <TrendingUp className="w-3.5 h-3.5 text-zinc-600 dark:text-zinc-400" />
              <span>أثر قاعدة الـ 3 بصمات</span>
            </button>
          </div>

          {/* Quiet Linear/Stripe inline KPI Ribbon */}
          <div className="flex items-center gap-3 text-xs text-zinc-600 dark:text-zinc-400">
            <div className="flex items-center gap-1.5">
              <span className="text-zinc-400">القوة العاملة:</span>
              <span className="font-mono font-bold text-zinc-900 dark:text-zinc-100">{kpis.totalEmp}</span>
              <span className="text-zinc-400 font-mono text-[10px]">({kpis.compliantEmp} مطابق · {kpis.deficientEmp} نقص)</span>
            </div>

            <span className="text-zinc-300 dark:text-zinc-700">·</span>

            <div className="flex items-center gap-1.5">
              <span className="text-zinc-400">أيام البصمات:</span>
              <span className="font-mono font-bold text-zinc-900 dark:text-zinc-100">{kpis.totalDays}</span>
            </div>

            <span className="text-zinc-300 dark:text-zinc-700">·</span>

            <div className="flex items-center gap-1.5">
              <span className="text-zinc-400">النوبات المكتملة:</span>
              <span className="font-mono font-bold text-emerald-600 dark:text-emerald-400">{kpis.completedShifts} / {kpis.totalShifts}</span>
            </div>

            <span className="text-zinc-300 dark:text-zinc-700">·</span>

            <div className="flex items-center gap-1.5">
              <span className="text-zinc-400">الالتزام العام:</span>
              <span className="font-mono font-bold text-zinc-900 dark:text-zinc-100">{kpis.avgCompliance}%</span>
            </div>

            {kpis.exemptExitsCount > 0 && (
              <>
                <span className="text-zinc-300 dark:text-zinc-700">·</span>
                <div 
                  onClick={() => setStatusFilter(statusFilter === 'previous_exits' ? 'all' : 'previous_exits')}
                  className="cursor-pointer flex items-center gap-1 text-sky-700 dark:text-sky-400 hover:underline"
                  title="نوبات عمل بدأت بالشهر السابق وانتهت أول هذا الشهر، مستبعدة من الاحتساب"
                >
                  <Clock className="w-3 h-3 text-sky-500" />
                  <span className="font-bold">{kpis.exemptExitsCount} نوبة سابقة معفاة</span>
                </div>
              </>
            )}
          </div>

        </div>
      </div>

      {/* ── MAIN CONTENT WORKSPACE ── */}
      <div className="flex-1 flex overflow-hidden relative">

        {/* ════ VIEW 1: HIGH-DENSITY AUDIT MATRIX (جدول تدقيق النوبات المدمج) ════ */}
        {activeTab === 'matrix' && (
          <div className="flex-1 flex flex-col h-full overflow-hidden bg-white dark:bg-zinc-900">
            
            {/* Filter & Search Bar */}
            <div className="h-10 px-4 border-b border-zinc-200 dark:border-zinc-800 flex items-center justify-between gap-4 bg-zinc-50/60 dark:bg-zinc-900/60 shrink-0">
              <div className="flex items-center gap-2 flex-1 max-w-md">
                <Search className="w-3.5 h-3.5 text-zinc-400" />
                <input
                  ref={searchInputRef}
                  type="text"
                  value={searchQuery}
                  onChange={(e) => setSearchQuery(e.target.value)}
                  placeholder="البحث بالاسم أو الرقم الوظيفي... (اضغط ⌘K للتركيز)"
                  className="bg-transparent border-none text-xs text-zinc-900 dark:text-zinc-100 placeholder-zinc-400 focus:outline-none w-full"
                />
                {searchQuery && (
                  <button onClick={() => setSearchQuery('')} className="text-zinc-400 hover:text-zinc-600">
                    <X className="w-3 h-3" />
                  </button>
                )}
              </div>

              {/* Status Filter Buttons */}
              <div className="flex items-center gap-1 text-xs">
                <span className="text-zinc-400 text-[11px] font-medium pl-1">تصفية:</span>
                <button
                  onClick={() => setStatusFilter('all')}
                  className={`px-2 py-0.5 rounded text-[11px] font-bold transition-colors ${
                    statusFilter === 'all'
                      ? 'bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900'
                      : 'text-zinc-600 dark:text-zinc-400 hover:bg-zinc-200/50'
                  }`}
                >
                  الكل ({employees.length})
                </button>
                <button
                  onClick={() => setStatusFilter('compliant')}
                  className={`px-2 py-0.5 rounded text-[11px] font-bold transition-colors ${
                    statusFilter === 'compliant'
                      ? 'bg-emerald-600 text-white'
                      : 'text-zinc-600 dark:text-zinc-400 hover:bg-zinc-200/50'
                  }`}
                >
                  مطابق ({kpis.compliantEmp})
                </button>
                <button
                  onClick={() => setStatusFilter('deficient')}
                  className={`px-2 py-0.5 rounded text-[11px] font-bold transition-colors ${
                    statusFilter === 'deficient'
                      ? 'bg-rose-600 text-white'
                      : 'text-zinc-600 dark:text-zinc-400 hover:bg-zinc-200/50'
                  }`}
                >
                  نواقص ({kpis.deficientEmp})
                </button>
                {kpis.empWithExemptExits > 0 && (
                  <button
                    onClick={() => setStatusFilter('previous_exits')}
                    className={`px-2 py-0.5 rounded text-[11px] font-bold transition-colors ${
                      statusFilter === 'previous_exits'
                        ? 'bg-sky-600 text-white'
                        : 'text-zinc-600 dark:text-zinc-400 hover:bg-zinc-200/50'
                    }`}
                  >
                    نوبة سابقة معفاة ({kpis.empWithExemptExits})
                  </button>
                )}
              </div>
            </div>

            {/* High Density Table Container */}
            <div className="flex-1 overflow-auto custom-scroll">
              <table className="w-full text-right border-collapse">
                <thead className="sticky top-0 bg-zinc-100 dark:bg-zinc-800 text-zinc-600 dark:text-zinc-300 text-[11px] font-bold border-b border-zinc-200 dark:border-zinc-700 z-10">
                  <tr>
                    <th className="py-2 px-3 w-10 text-center">ت</th>
                    <th 
                      onClick={() => handleSort('id')} 
                      className="py-2 px-3 cursor-pointer hover:text-zinc-900 dark:hover:text-zinc-100 transition-colors w-24"
                    >
                      <div className="flex items-center gap-1">
                        <span>الرقم الوظيفي</span>
                        <ArrowUpDown className="w-3 h-3 text-zinc-400" />
                      </div>
                    </th>
                    <th 
                      onClick={() => handleSort('name')} 
                      className="py-2 px-3 cursor-pointer hover:text-zinc-900 dark:hover:text-zinc-100 transition-colors"
                    >
                      <div className="flex items-center gap-1">
                        <span>اسم الموظف</span>
                        <ArrowUpDown className="w-3 h-3 text-zinc-400" />
                      </div>
                    </th>
                    <th 
                      onClick={() => handleSort('days')} 
                      className="py-2 px-3 cursor-pointer hover:text-zinc-900 dark:hover:text-zinc-100 transition-colors w-28"
                      title="عدد أيام البصمات الفعلية التي حضر فيها الموظف"
                    >
                      <div className="flex items-center gap-1">
                        <span>أيام البصمات</span>
                        <ArrowUpDown className="w-3 h-3 text-zinc-400" />
                      </div>
                    </th>
                    <th className="py-2 px-3 w-28 text-center">نوبات الشهر</th>
                    <th 
                      onClick={() => handleSort('completed')} 
                      className="py-2 px-3 cursor-pointer hover:text-zinc-900 dark:hover:text-zinc-100 transition-colors w-28 text-center"
                    >
                      <div className="flex items-center justify-center gap-1">
                        <span>المكتملة</span>
                        <ArrowUpDown className="w-3 h-3 text-zinc-400" />
                      </div>
                    </th>
                    <th className="py-2 px-3 w-24 text-center">النواقص</th>
                    <th className="py-2 px-3 w-32 text-center" title="بصمات خروج لنوبات سابقة معفاة من احتساب الشهر الحالي">
                      نوبة سابقة معفاة
                    </th>
                    <th 
                      onClick={() => handleSort('compliance')} 
                      className="py-2 px-3 cursor-pointer hover:text-zinc-900 dark:hover:text-zinc-100 transition-colors w-36"
                    >
                      <div className="flex items-center gap-1">
                        <span>معدل الالتزام</span>
                        <ArrowUpDown className="w-3 h-3 text-zinc-400" />
                      </div>
                    </th>
                    <th className="py-2 px-3 w-28 text-center">حالة التدقيق</th>
                    <th className="py-2 px-3 w-16 text-center">فحص</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-zinc-200 dark:divide-zinc-800 text-xs">
                  {filteredEmployees.length === 0 ? (
                    <tr>
                      <td colSpan={11} className="py-16 text-center text-zinc-400">
                        لا توجد سجلات تطابق معايير البحث الحالية
                      </td>
                    </tr>
                  ) : (
                    filteredEmployees.map((emp, idx) => {
                      const isSelected = selectedEmployee?.id === emp.id;
                      const auditedShiftsCount = emp.shifts.filter(s => !s.is_previous_shift).length;
                      const previousExitsCount = emp.shifts.filter(s => s.is_previous_shift).length;
                      return (
                        <tr
                          key={emp.id}
                          onClick={() => setSelectedEmployee(emp)}
                          className={`cursor-pointer transition-colors ${
                            isSelected 
                              ? 'bg-zinc-100/90 dark:bg-zinc-800 font-medium' 
                              : 'hover:bg-zinc-50 dark:hover:bg-zinc-800/40'
                          }`}
                        >
                          <td className="py-2 px-3 text-center text-zinc-400 font-mono text-[11px]">
                            {String(idx + 1).padStart(2, '0')}
                          </td>
                          <td className="py-2 px-3 font-mono font-bold text-zinc-800 dark:text-zinc-200">
                            {emp.id}
                          </td>
                          <td className="py-2 px-3 font-bold text-zinc-900 dark:text-zinc-100">
                            {emp.name}
                          </td>

                          {/* أيام البصمات */}
                          <td className="py-2 px-3 font-mono">
                            <span className="font-bold text-zinc-800 dark:text-zinc-200">
                              {emp.fingerprintDaysCount}
                            </span>
                            <span className="text-zinc-400 text-[11px]"> / {emp.totalPeriodDays} يوم</span>
                          </td>

                          {/* نوبات الشهر */}
                          <td className="py-2 px-3 text-center font-mono font-bold text-zinc-700 dark:text-zinc-300">
                            {auditedShiftsCount}
                          </td>

                          {/* النوبات المكتملة */}
                          <td className="py-2 px-3 text-center font-mono font-bold text-emerald-600 dark:text-emerald-400">
                            {emp.completedShiftsCount}
                          </td>

                          {/* نواقص النوبات */}
                          <td className="py-2 px-3 text-center font-mono font-bold">
                            {emp.missingShiftsCount > 0 ? (
                              <span className="text-rose-600 dark:text-rose-400 font-bold">
                                {emp.missingShiftsCount}
                              </span>
                            ) : (
                              <span className="text-zinc-400">0</span>
                            )}
                          </td>

                          {/* بصمات خروج سابقة معفاة */}
                          <td className="py-2 px-3 text-center font-mono">
                            {previousExitsCount > 0 ? (
                              <span className="inline-flex items-center gap-1 text-[11px] font-bold text-sky-700 dark:text-sky-300 bg-sky-50 dark:bg-sky-950/60 px-2 py-0.5 rounded border border-sky-200 dark:border-sky-800/60">
                                <Clock className="w-3 h-3 text-sky-500" />
                                {previousExitsCount} معفاة
                              </span>
                            ) : (
                              <span className="text-zinc-400 text-[11px]">-</span>
                            )}
                          </td>

                          {/* معدل الالتزام */}
                          <td className="py-2 px-3">
                            <div className="flex items-center gap-2">
                              <div className="w-16 h-1.5 bg-zinc-200 dark:bg-zinc-700 rounded-full overflow-hidden shrink-0">
                                <div 
                                  className={`h-full rounded-full ${
                                    emp.completionRate === 100 
                                      ? 'bg-emerald-500' 
                                      : emp.completionRate >= 80 
                                        ? 'bg-amber-500' 
                                        : 'bg-rose-500'
                                  }`}
                                  style={{ width: `${emp.completionRate}%` }}
                                />
                              </div>
                              <span className="font-mono font-bold text-zinc-900 dark:text-zinc-100 text-[11px]">
                                {emp.completionRate}%
                              </span>
                            </div>
                          </td>

                          {/* حالة التدقيق */}
                          <td className="py-2 px-3 text-center">
                            {emp.isEligible ? (
                              <span className="inline-flex items-center gap-1 text-[10px] font-bold text-emerald-700 dark:text-emerald-300 bg-emerald-50 dark:bg-emerald-950/50 px-2 py-0.5 rounded border border-emerald-200 dark:border-emerald-800">
                                <Check className="w-2.5 h-2.5 stroke-[3]" />
                                مطابق
                              </span>
                            ) : (
                              <span className="inline-flex items-center gap-1 text-[10px] font-bold text-amber-700 dark:text-amber-300 bg-amber-50 dark:bg-amber-950/50 px-2 py-0.5 rounded border border-amber-200 dark:border-amber-800">
                                <AlertTriangle className="w-2.5 h-2.5 stroke-[2.5]" />
                                نقص {emp.missingShiftsCount}
                              </span>
                            )}
                          </td>

                          {/* زر الفحص */}
                          <td className="py-2 px-3 text-center">
                            <button
                              onClick={(e) => {
                                e.stopPropagation();
                                setSelectedEmployee(emp);
                              }}
                              className="w-6 h-6 rounded flex items-center justify-center text-zinc-500 hover:text-zinc-900 dark:hover:text-zinc-100 hover:bg-zinc-200 dark:hover:bg-zinc-700 transition-colors mx-auto"
                              title="فتح سجل تفاصيل نوبات الموظف"
                            >
                              <Eye className="w-3.5 h-3.5" />
                            </button>
                          </td>
                        </tr>
                      );
                    })
                  )}
                </tbody>
              </table>
            </div>

            {/* Table Footer Summary Bar */}
            <div className="h-8 px-4 border-t border-zinc-200 dark:border-zinc-800 bg-zinc-50 dark:bg-zinc-900 flex items-center justify-between text-[11px] text-zinc-500 shrink-0">
              <div>
                عرض <b className="font-mono text-zinc-700 dark:text-zinc-300">{filteredEmployees.length}</b> من أصل <b className="font-mono text-zinc-700 dark:text-zinc-300">{employees.length}</b> موظفاً
              </div>
              <div className="flex items-center gap-3">
                <span>القاعدة النشطة: <b className="text-zinc-700 dark:text-zinc-300">{punchMode === '3_punches' ? '3 بصمات لكل نوبة' : '5 بصمات لكل نوبة'}</b></span>
                <span>·</span>
                <span className="text-sky-600 dark:text-sky-400">بصمات خروج النوبات السابقة مستبعدة تلقائياً من الاحتساب</span>
              </div>
            </div>

          </div>
        )}

        {/* ════ VIEW 2: EXECUTIVE ANALYTICS DASHBOARD (لوحة المؤشرات والتحليل) ════ */}
        {activeTab === 'dashboard' && (
          <div className="flex-1 flex flex-col h-full overflow-auto custom-scroll bg-[#fafafa] dark:bg-[#09090b] p-6">
            <div className="max-w-6xl mx-auto w-full space-y-6">
              
              {/* Executive Header Card */}
              <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 p-5 rounded-lg border border-zinc-200 dark:border-zinc-800 bg-white dark:bg-zinc-900 shadow-xs">
                <div>
                  <div className="flex items-center gap-2">
                    <h1 className="text-sm font-extrabold text-zinc-900 dark:text-zinc-100">
                      لوحة قيادة وتحليل نوبات العمل الشهرية
                    </h1>
                    <span className="px-2 py-0.5 rounded text-[10px] font-bold bg-emerald-50 text-emerald-700 dark:bg-emerald-950/40 dark:text-emerald-400 border border-emerald-200 dark:border-emerald-800">
                      محضر معتمد
                    </span>
                  </div>
                  <p className="text-zinc-500 text-xs mt-1 leading-relaxed">
                    تدقيق شامل لانتظام الكوادر وبصمات الدخول، التواجد، والانصراف وفق {punchMode === '3_punches' ? 'أمر تخفيض البصمات (3 بصمات)' : 'نظام البصمات الكامل (5 بصمات)'}.
                  </p>
                </div>

                <div className="flex items-center gap-2">
                  <button
                    onClick={() => setActiveTab('matrix')}
                    className="h-8 px-3 rounded-md bg-zinc-900 hover:bg-zinc-800 dark:bg-zinc-100 dark:hover:bg-white text-white dark:text-zinc-950 font-bold flex items-center gap-1.5 transition-colors shadow-xs"
                  >
                    <span>فتح جدول التدقيق المدمج</span>
                    <ChevronLeft className="w-3.5 h-3.5" />
                  </button>
                  <button
                    onClick={handleExportExcel}
                    className="h-8 px-3 rounded-md border border-zinc-200 dark:border-zinc-700 bg-white hover:bg-zinc-50 dark:bg-zinc-800 text-zinc-800 dark:text-zinc-200 font-semibold flex items-center gap-1.5 transition-colors"
                  >
                    <Download className="w-3.5 h-3.5 text-zinc-500" />
                    <span>تصدير Excel</span>
                  </button>
                </div>
              </div>

              {/* 4 PRIMARY METRIC CARDS */}
              <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
                
                {/* 1. القوة العاملة */}
                <div className="p-4 rounded-lg border border-zinc-200 dark:border-zinc-800 bg-white dark:bg-zinc-900 shadow-xs flex flex-col justify-between">
                  <div className="flex items-center justify-between">
                    <span className="text-xs font-bold text-zinc-500">القوة العاملة المدققة</span>
                    <div className="w-7 h-7 rounded-md bg-zinc-100 dark:bg-zinc-800 flex items-center justify-center text-zinc-600 dark:text-zinc-300">
                      <Users className="w-4 h-4" />
                    </div>
                  </div>
                  <div className="mt-3">
                    <div className="text-2xl font-black font-mono text-zinc-900 dark:text-zinc-100">
                      {kpis.totalEmp} <span className="text-xs font-sans font-normal text-zinc-500">موظف</span>
                    </div>
                    <div className="mt-2 flex items-center justify-between text-[11px] text-zinc-500 border-t border-zinc-100 dark:border-zinc-800 pt-2 font-mono">
                      <span>مطابقون: <b className="text-emerald-600 dark:text-emerald-400">{kpis.compliantEmp}</b></span>
                      <span>نواقص: <b className="text-amber-600 dark:text-amber-400">{kpis.deficientEmp}</b></span>
                    </div>
                  </div>
                </div>

                {/* 2. معدل التزام النوبات */}
                <div className="p-4 rounded-lg border border-zinc-200 dark:border-zinc-800 bg-white dark:bg-zinc-900 shadow-xs flex flex-col justify-between">
                  <div className="flex items-center justify-between">
                    <span className="text-xs font-bold text-zinc-500">معدل التزام النوبات</span>
                    <div className="w-7 h-7 rounded-md bg-emerald-50 dark:bg-emerald-950/40 flex items-center justify-center text-emerald-600 dark:text-emerald-400">
                      <Award className="w-4 h-4" />
                    </div>
                  </div>
                  <div className="mt-3">
                    <div className="flex items-baseline gap-2">
                      <div className="text-2xl font-black font-mono text-zinc-900 dark:text-zinc-100">
                        {kpis.avgCompliance}%
                      </div>
                      <span className="text-[10px] font-bold text-emerald-600 bg-emerald-50 dark:bg-emerald-950 px-1.5 py-0.2 rounded">
                        {kpis.avgCompliance >= 90 ? 'مستوى ممتاز' : kpis.avgCompliance >= 75 ? 'مستوى جيد' : 'يحتاج متابعة'}
                      </span>
                    </div>
                    <div className="mt-2 w-full h-1.5 bg-zinc-100 dark:bg-zinc-800 rounded-full overflow-hidden">
                      <div 
                        className="h-full bg-emerald-500 rounded-full transition-all duration-500" 
                        style={{ width: `${kpis.avgCompliance}%` }}
                      />
                    </div>
                  </div>
                </div>

                {/* 3. النوبات المكتملة */}
                <div className="p-4 rounded-lg border border-zinc-200 dark:border-zinc-800 bg-white dark:bg-zinc-900 shadow-xs flex flex-col justify-between">
                  <div className="flex items-center justify-between">
                    <span className="text-xs font-bold text-zinc-500">النوبات المكتملة</span>
                    <div className="w-7 h-7 rounded-md bg-zinc-100 dark:bg-zinc-800 flex items-center justify-center text-zinc-600 dark:text-zinc-300">
                      <CheckCircle2 className="w-4 h-4 text-emerald-600" />
                    </div>
                  </div>
                  <div className="mt-3">
                    <div className="text-2xl font-black font-mono text-zinc-900 dark:text-zinc-100">
                      {kpis.completedShifts} <span className="text-xs font-sans font-normal text-zinc-500">من أصل {kpis.totalShifts} نوبة</span>
                    </div>
                    <div className="mt-2 flex items-center justify-between text-[11px] text-zinc-500 border-t border-zinc-100 dark:border-zinc-800 pt-2 font-mono">
                      <span>النواقص: <b className="text-rose-600">{kpis.totalDeficiencies} نوبة</b></span>
                      <span>معفاة: <b className="text-zinc-600 dark:text-zinc-400">{kpis.exemptExitsCount}</b></span>
                    </div>
                  </div>
                </div>

                {/* 4. أيام البصمات */}
                <div className="p-4 rounded-lg border border-zinc-200 dark:border-zinc-800 bg-white dark:bg-zinc-900 shadow-xs flex flex-col justify-between">
                  <div className="flex items-center justify-between">
                    <span className="text-xs font-bold text-zinc-500">أيام البصمات المقيدة</span>
                    <div className="w-7 h-7 rounded-md bg-zinc-100 dark:bg-zinc-800 flex items-center justify-center text-zinc-600 dark:text-zinc-300">
                      <Calendar className="w-4 h-4" />
                    </div>
                  </div>
                  <div className="mt-3">
                    <div className="text-2xl font-black font-mono text-zinc-900 dark:text-zinc-100">
                      {kpis.totalDays} <span className="text-xs font-sans font-normal text-zinc-500">يوم عمل مدقق</span>
                    </div>
                    <div className="mt-2 flex items-center justify-between text-[11px] text-zinc-500 border-t border-zinc-100 dark:border-zinc-800 pt-2">
                      <span>القاعدة: <b>{punchMode === '3_punches' ? '3 بصمات' : '5 بصمات'}</b></span>
                      <span className="font-mono text-emerald-600 font-bold">نشطة</span>
                    </div>
                  </div>
                </div>

              </div>

              {/* SPECIAL AUDIT ALERT: PREVIOUS SHIFT EXITS (معالجة النوبات السابقة) */}
              {previousShiftExits.length > 0 && (
                <div className="p-4 rounded-lg border border-sky-200 dark:border-sky-900/60 bg-sky-50/50 dark:bg-sky-950/20 shadow-xs">
                  <div className="flex items-center justify-between mb-2">
                    <div className="flex items-center gap-2">
                      <div className="w-6 h-6 rounded bg-sky-100 dark:bg-sky-900/60 text-sky-700 dark:text-sky-300 flex items-center justify-center">
                        <Clock className="w-3.5 h-3.5" />
                      </div>
                      <span className="font-bold text-xs text-sky-900 dark:text-sky-200">
                        حالات بصمة خروج لنوبة سابقة (معفاة تلقائياً من احتساب الشهر الحالي)
                      </span>
                    </div>
                    <span className="text-[11px] font-mono font-bold text-sky-700 dark:text-sky-300 bg-sky-100/70 dark:bg-sky-900/40 px-2 py-0.5 rounded">
                      {previousShiftExits.length} حالة موثقة
                    </span>
                  </div>
                  <p className="text-[11px] text-sky-800/80 dark:text-sky-300/80 leading-relaxed mb-3">
                    وفقاً للضوابط الإدارية المعتمدة، عند رصد بصمة خروج في مستهل الشهر ناتجة عن نوبة عمل بدأت في الشهر السابق، يقوم النظام بوسمها كـ <b>«بصمة خروج لنوبة سابقة»</b> واستبعادها من احتساب نوبات ونواقص الشهر الحالي لضمان عدم إلحاق أي غبن إداري بالموظف.
                  </p>
                  <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-2">
                    {previousShiftExits.map((item, idx) => (
                      <div 
                        key={idx}
                        onClick={() => setSelectedEmployee(item.emp)}
                        className="cursor-pointer p-2.5 rounded bg-white dark:bg-zinc-900 border border-sky-200 dark:border-sky-800/50 flex items-center justify-between text-xs hover:border-sky-400 transition-colors"
                      >
                        <div className="flex items-center gap-2">
                          <span className="font-mono font-bold text-zinc-700 dark:text-zinc-300">{item.emp.id}</span>
                          <span className="font-bold text-zinc-900 dark:text-zinc-100 truncate">{item.emp.name}</span>
                        </div>
                        <span className="text-[10px] font-bold text-sky-700 dark:text-sky-300 bg-sky-50 dark:bg-sky-950 px-1.5 py-0.5 rounded">
                          فحص السجل
                        </span>
                      </div>
                    ))}
                  </div>
                </div>
              )}

              {/* TWO COLUMN INSIGHTS: TOP COMPLIANT VS PRIORITY DEFICIENCIES */}
              <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
                
                {/* 1. الموظفون الأكثر انضباطاً (100% التزام) */}
                <div className="p-4 rounded-lg border border-zinc-200 dark:border-zinc-800 bg-white dark:bg-zinc-900 shadow-xs flex flex-col">
                  <div className="flex items-center justify-between border-b border-zinc-100 dark:border-zinc-800 pb-3 mb-3">
                    <div className="flex items-center gap-2">
                      <div className="w-6 h-6 rounded bg-emerald-50 dark:bg-emerald-950/60 text-emerald-600 flex items-center justify-center">
                        <CheckCircle2 className="w-3.5 h-3.5" />
                      </div>
                      <span className="font-bold text-xs text-zinc-900 dark:text-zinc-100">
                        عينة الموظفين المكتملة نوباتهم (التزام 100%)
                      </span>
                    </div>
                    <span className="text-[11px] font-mono text-zinc-400">
                      {kpis.compliantEmp} من أصل {kpis.totalEmp}
                    </span>
                  </div>

                  <div className="space-y-2 flex-1">
                    {topCompliant.length === 0 ? (
                      <p className="text-center text-zinc-400 py-6 text-xs">لا يوجد موظفون بنسبة 100% حالياً</p>
                    ) : (
                      topCompliant.map((emp) => (
                        <div 
                          key={emp.id}
                          onClick={() => setSelectedEmployee(emp)}
                          className="cursor-pointer p-2.5 rounded bg-zinc-50 dark:bg-zinc-800/40 hover:bg-zinc-100 dark:hover:bg-zinc-800 border border-zinc-200/60 dark:border-zinc-700/60 flex items-center justify-between text-xs transition-colors"
                        >
                          <div className="flex items-center gap-2.5">
                            <span className="w-7 h-7 rounded bg-white dark:bg-zinc-900 text-zinc-700 dark:text-zinc-300 font-mono font-bold text-[11px] flex items-center justify-center border border-zinc-200 dark:border-zinc-700">
                              {emp.id}
                            </span>
                            <div>
                              <div className="font-bold text-zinc-900 dark:text-zinc-100">{emp.name}</div>
                              <div className="text-[10px] text-zinc-400 font-mono">
                                أيام البصمات: {emp.fingerprintDaysCount} • النوبات المكتملة: {emp.completedShiftsCount}
                              </div>
                            </div>
                          </div>
                          <div className="flex items-center gap-2 font-mono">
                            <span className="px-2 py-0.5 rounded text-[10px] font-bold bg-emerald-100 text-emerald-800 dark:bg-emerald-950 dark:text-emerald-300">
                              100% مكتمل
                            </span>
                          </div>
                        </div>
                      ))
                    )}
                  </div>
                </div>

                {/* 2. حالات النواقص ذات الأولوية */}
                <div className="p-4 rounded-lg border border-zinc-200 dark:border-zinc-800 bg-white dark:bg-zinc-900 shadow-xs flex flex-col">
                  <div className="flex items-center justify-between border-b border-zinc-100 dark:border-zinc-800 pb-3 mb-3">
                    <div className="flex items-center gap-2">
                      <div className="w-6 h-6 rounded bg-amber-50 dark:bg-amber-950/60 text-amber-600 flex items-center justify-center">
                        <AlertTriangle className="w-3.5 h-3.5" />
                      </div>
                      <span className="font-bold text-xs text-zinc-900 dark:text-zinc-100">
                        أبرز حالات النواقص في النوبات (أولوية المراجعة)
                      </span>
                    </div>
                    <button
                      onClick={() => setActiveTab('deficiencies')}
                      className="text-[11px] font-bold text-zinc-600 dark:text-zinc-400 hover:text-zinc-900 dark:hover:text-zinc-200 flex items-center gap-1"
                    >
                      <span>عرض المحضر الكامل</span>
                      <ChevronLeft className="w-3 h-3" />
                    </button>
                  </div>

                  <div className="space-y-2 flex-1">
                    {priorityAttention.length === 0 ? (
                      <div className="text-center py-6 text-emerald-600 font-bold text-xs flex flex-col items-center gap-1">
                        <CheckCircle2 className="w-6 h-6 text-emerald-500" />
                        <span>جميع النوبات مستوفاة بالكامل وفق القاعدة الحالية!</span>
                      </div>
                    ) : (
                      priorityAttention.map((emp) => (
                        <div 
                          key={emp.id}
                          onClick={() => setSelectedEmployee(emp)}
                          className="cursor-pointer p-2.5 rounded bg-zinc-50 dark:bg-zinc-800/40 hover:bg-zinc-100 dark:hover:bg-zinc-800 border border-zinc-200/60 dark:border-zinc-700/60 flex items-center justify-between text-xs transition-colors"
                        >
                          <div className="flex items-center gap-2.5">
                            <span className="w-7 h-7 rounded bg-white dark:bg-zinc-900 text-zinc-700 dark:text-zinc-300 font-mono font-bold text-[11px] flex items-center justify-center border border-zinc-200 dark:border-zinc-700">
                              {emp.id}
                            </span>
                            <div>
                              <div className="font-bold text-zinc-900 dark:text-zinc-100">{emp.name}</div>
                              <div className="text-[10px] text-zinc-400 font-mono">
                                أيام البصمات: {emp.fingerprintDaysCount} • نقص {emp.missingShiftsCount} نوبة
                              </div>
                            </div>
                          </div>
                          <div className="flex items-center gap-2 font-mono">
                            <span className="px-2 py-0.5 rounded text-[10px] font-bold bg-amber-100 text-amber-800 dark:bg-amber-950 dark:text-amber-300">
                              الالتزام {emp.completionRate}%
                            </span>
                          </div>
                        </div>
                      ))
                    )}
                  </div>
                </div>

              </div>

            </div>
          </div>
        )}

        {/* ════ VIEW 3: DEFICIENCIES LOG (محضر النواقص التفصيلي) ════ */}
        {activeTab === 'deficiencies' && (
          <div className="flex-1 flex flex-col h-full overflow-hidden bg-white dark:bg-zinc-900">
            <div className="h-11 px-4 border-b border-zinc-200 dark:border-zinc-800 flex items-center justify-between bg-zinc-50/60 dark:bg-zinc-900/60 shrink-0">
              <div className="flex items-center gap-2">
                <AlertTriangle className="w-4 h-4 text-amber-600" />
                <span className="font-bold text-xs text-zinc-900 dark:text-zinc-100">
                  محضر النواقص التفصيلي للنوبات
                </span>
                <span className="text-[11px] text-zinc-500 font-mono">
                  (إجمالي {allDeficiencies.length} نوبة غير مكتملة)
                </span>
              </div>
              <div className="flex items-center gap-2">
                <button
                  onClick={handleExportExcel}
                  className="h-7 px-2.5 rounded border border-zinc-200 dark:border-zinc-700 bg-white dark:bg-zinc-800 text-xs font-bold flex items-center gap-1.5 hover:bg-zinc-50 transition-colors"
                >
                  <Download className="w-3.5 h-3.5 text-zinc-500" />
                  <span>تصدير محضر النواقص</span>
                </button>
              </div>
            </div>

            <div className="flex-1 overflow-auto custom-scroll">
              <table className="w-full text-right border-collapse">
                <thead className="sticky top-0 bg-zinc-100 dark:bg-zinc-800 text-zinc-600 dark:text-zinc-300 text-[11px] font-bold border-b border-zinc-200 dark:border-zinc-700 z-10">
                  <tr>
                    <th className="py-2.5 px-3 w-10 text-center">ت</th>
                    <th className="py-2.5 px-3 w-28">الرقم الوظيفي</th>
                    <th className="py-2.5 px-3 w-48">اسم الموظف</th>
                    <th className="py-2.5 px-3 w-32">تاريخ النوبة</th>
                    <th className="py-2.5 px-3 w-28 text-center">البصمات المسجلة</th>
                    <th className="py-2.5 px-3 w-48">البصمات الناقصة المطلوبة</th>
                    <th className="py-2.5 px-3">تفاصيل البصمات المقيدة للنوبة</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-zinc-200 dark:divide-zinc-800 text-xs">
                  {allDeficiencies.length === 0 ? (
                    <tr>
                      <td colSpan={7} className="py-20 text-center text-emerald-600 font-bold">
                        <CheckCircle2 className="w-8 h-8 mx-auto mb-2 text-emerald-500" />
                        ممتاز! لا توجد أي نوبات ناقصة مسجلة في هذا الشهر وفق القاعدة الحالية.
                      </td>
                    </tr>
                  ) : (
                    allDeficiencies.map((d, idx) => (
                      <tr 
                        key={idx} 
                        onClick={() => {
                          const emp = employees.find(e => String(e.id) === String(d.empId));
                          if (emp) setSelectedEmployee(emp);
                        }}
                        className="hover:bg-zinc-50 dark:hover:bg-zinc-800/40 cursor-pointer transition-colors"
                      >
                        <td className="py-2 px-3 text-center text-zinc-400 font-mono text-[11px]">
                          {idx + 1}
                        </td>
                        <td className="py-2 px-3 font-mono font-bold text-zinc-800 dark:text-zinc-200">
                          {d.empId}
                        </td>
                        <td className="py-2 px-3 font-bold text-zinc-900 dark:text-zinc-100">
                          {d.empName}
                        </td>
                        <td className="py-2 px-3 font-mono text-zinc-700 dark:text-zinc-300">
                          {d.shiftDate}
                        </td>
                        <td className="py-2 px-3 text-center font-mono font-bold text-rose-600">
                          {d.recordedCount} من {punchMode === '3_punches' ? 3 : 5}
                        </td>
                        <td className="py-2 px-3">
                          <span className="text-rose-600 dark:text-rose-400 font-bold">
                            {d.missingPeriods.join(' ، ')}
                          </span>
                        </td>
                        <td className="py-2 px-3 text-zinc-500 dark:text-zinc-400 font-mono text-[11px]">
                          {d.punchesText}
                        </td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* ════ VIEW 4: RULE COMPARISON (دراسة أثر الـ 3 بصمات) ════ */}
        {activeTab === 'rule_comparison' && (
          <div className="flex-1 flex flex-col h-full overflow-auto custom-scroll bg-[#fafafa] dark:bg-[#09090b] p-6">
            <div className="max-w-5xl mx-auto w-full space-y-6">
              
              <div className="p-5 rounded-lg border border-zinc-200 dark:border-zinc-800 bg-white dark:bg-zinc-900 shadow-xs">
                <div className="flex items-center gap-2 mb-1">
                  <TrendingUp className="w-4 h-4 text-emerald-600" />
                  <h2 className="text-sm font-extrabold text-zinc-900 dark:text-zinc-100">
                    تقرير الأثر الإداري لقاعدة تقليل البصمات (3 بصمات بدلاً من 5)
                  </h2>
                </div>
                <p className="text-xs text-zinc-500 leading-relaxed">
                  دراسة إحصائية مقارنة توضح التغير في نسب التزام الكادر الوظيفي وعدد الموظفين الذين ارتفعت نسبة التزامهم وانتقلت حالتهم إلى مطابق تماماً.
                </p>
              </div>

              {/* Comparative Stats Cards */}
              <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                
                <div className="p-4 rounded-lg border border-zinc-200 dark:border-zinc-800 bg-white dark:bg-zinc-900 shadow-xs">
                  <span className="text-xs font-bold text-zinc-500">الموظفون المطابقون تماماً</span>
                  <div className="mt-3 flex items-baseline gap-2">
                    <span className="text-2xl font-black font-mono text-emerald-600">
                      {comparisonData.comp3Count}
                    </span>
                    <span className="text-xs text-zinc-400 font-mono">
                      (مقابل {comparisonData.comp5Count} بنظام 5 بصمات)
                    </span>
                  </div>
                  <div className="mt-2 text-[11px] text-emerald-600 font-bold">
                    +{comparisonData.comp3Count - comparisonData.comp5Count} موظفاً دخلوا دائرة المطابقة
                  </div>
                </div>

                <div className="p-4 rounded-lg border border-zinc-200 dark:border-zinc-800 bg-white dark:bg-zinc-900 shadow-xs">
                  <span className="text-xs font-bold text-zinc-500">متوسط الالتزام العام للكادر</span>
                  <div className="mt-3 flex items-baseline gap-2">
                    <span className="text-2xl font-black font-mono text-zinc-900 dark:text-zinc-100">
                      {comparisonData.avg3}%
                    </span>
                    <span className="text-xs text-zinc-400 font-mono">
                      (مقابل {comparisonData.avg5}% بنظام 5 بصمات)
                    </span>
                  </div>
                  <div className="mt-2 text-[11px] text-emerald-600 font-bold">
                    تحسن بنسبة +{comparisonData.avg3 - comparisonData.avg5}% في معدل الأداء
                  </div>
                </div>

                <div className="p-4 rounded-lg border border-zinc-200 dark:border-zinc-800 bg-white dark:bg-zinc-900 shadow-xs">
                  <span className="text-xs font-bold text-zinc-500">المستفيدون من أمر التخفيض</span>
                  <div className="mt-3 flex items-baseline gap-2">
                    <span className="text-2xl font-black font-mono text-zinc-900 dark:text-zinc-100">
                      {comparisonData.benefits.length}
                    </span>
                    <span className="text-xs text-zinc-500">موظفاً</span>
                  </div>
                  <div className="mt-2 text-[11px] text-zinc-500">
                    استفادوا بارتفاع تقييم نوباتهم أو إسقاط نواقص غير مبررة
                  </div>
                </div>

              </div>

              {/* Benefited Employees Table */}
              <div className="rounded-lg border border-zinc-200 dark:border-zinc-800 bg-white dark:bg-zinc-900 shadow-xs overflow-hidden">
                <div className="p-3 border-b border-zinc-200 dark:border-zinc-800 font-bold text-xs text-zinc-900 dark:text-zinc-100">
                  قائمة الموظفين المستفيدين من أمر الـ 3 بصمات
                </div>
                <div className="overflow-auto max-h-96 custom-scroll">
                  <table className="w-full text-right border-collapse">
                    <thead className="sticky top-0 bg-zinc-100 dark:bg-zinc-800 text-zinc-600 dark:text-zinc-300 text-[11px] font-bold border-b border-zinc-200 dark:border-zinc-700">
                      <tr>
                        <th className="py-2 px-3">الرقم الوظيفي</th>
                        <th className="py-2 px-3">اسم الموظف</th>
                        <th className="py-2 px-3 text-center">أيام البصمات</th>
                        <th className="py-2 px-3 text-center">التزام (5 بصمات)</th>
                        <th className="py-2 px-3 text-center">التزام (3 بصمات)</th>
                        <th className="py-2 px-3 text-center">فارق النسبة</th>
                        <th className="py-2 px-3 text-center">نوبات تم إنقاذها</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-zinc-200 dark:divide-zinc-800 text-xs">
                      {comparisonData.benefits.map((b) => (
                        <tr key={b.id} className="hover:bg-zinc-50 dark:hover:bg-zinc-800/40">
                          <td className="py-2 px-3 font-mono font-bold">{b.id}</td>
                          <td className="py-2 px-3 font-bold">{b.name}</td>
                          <td className="py-2 px-3 text-center font-mono">{b.days}</td>
                          <td className="py-2 px-3 text-center font-mono text-zinc-500">{b.rate5}%</td>
                          <td className="py-2 px-3 text-center font-mono font-bold text-emerald-600">{b.rate3}%</td>
                          <td className="py-2 px-3 text-center font-mono font-bold text-emerald-600">+{b.diff}%</td>
                          <td className="py-2 px-3 text-center font-mono font-bold text-emerald-700 dark:text-emerald-300">
                            {b.savedShifts} نوبة
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>

            </div>
          </div>
        )}

        {/* ════ SLIDE-OVER INSPECTION DRAWER: DETAILED EMPLOYEE SHIFTS ════ */}
        {selectedEmployee && (
          <div className="w-full max-w-xl h-full border-r border-zinc-200 dark:border-zinc-800 bg-white dark:bg-zinc-900 flex flex-col shadow-2xl z-40 animate-in slide-in-from-right duration-200 shrink-0">
            
            {/* Drawer Header */}
            <div className="h-14 px-4 border-b border-zinc-200 dark:border-zinc-800 flex items-center justify-between bg-zinc-50/80 dark:bg-zinc-900/80 shrink-0">
              <div className="flex items-center gap-2.5">
                <span className="w-8 h-8 rounded bg-zinc-900 dark:bg-zinc-100 text-white dark:text-zinc-900 font-mono font-extrabold text-xs flex items-center justify-center">
                  {selectedEmployee.id}
                </span>
                <div>
                  <div className="font-extrabold text-xs text-zinc-900 dark:text-zinc-100">
                    {selectedEmployee.name}
                  </div>
                  <div className="text-[10px] text-zinc-400 font-mono">
                    الرقم الوظيفي: {selectedEmployee.id}
                  </div>
                </div>
              </div>

              <div className="flex items-center gap-1.5">
                <button
                  onClick={() => setSelectedEmployee(null)}
                  className="w-7 h-7 rounded flex items-center justify-center text-zinc-400 hover:text-zinc-700 dark:hover:text-zinc-200 hover:bg-zinc-200/50 dark:hover:bg-zinc-800 transition-colors"
                  title="إغلاق (Esc)"
                >
                  <X className="w-4 h-4" />
                </button>
              </div>
            </div>

            {/* Quick Metrics Bar inside Drawer */}
            <div className="p-3 bg-zinc-100/60 dark:bg-zinc-800/60 border-b border-zinc-200 dark:border-zinc-800 grid grid-cols-4 gap-2 text-center text-xs shrink-0 font-mono">
              <div className="bg-white dark:bg-zinc-900 p-2 rounded border border-zinc-200/60 dark:border-zinc-700/60">
                <div className="text-[10px] text-zinc-400 font-sans">أيام البصمات</div>
                <div className="font-bold text-zinc-900 dark:text-zinc-100 mt-0.5">{selectedEmployee.fingerprintDaysCount}</div>
              </div>

              <div className="bg-white dark:bg-zinc-900 p-2 rounded border border-zinc-200/60 dark:border-zinc-700/60">
                <div className="text-[10px] text-zinc-400 font-sans">نوبات مكتملة</div>
                <div className="font-bold text-emerald-600 dark:text-emerald-400 mt-0.5">{selectedEmployee.completedShiftsCount}</div>
              </div>

              <div className="bg-white dark:bg-zinc-900 p-2 rounded border border-zinc-200/60 dark:border-zinc-700/60">
                <div className="text-[10px] text-zinc-400 font-sans">نواقص النوبات</div>
                <div className="font-bold text-rose-600 dark:text-rose-400 mt-0.5">{selectedEmployee.missingShiftsCount}</div>
              </div>

              <div className="bg-white dark:bg-zinc-900 p-2 rounded border border-zinc-200/60 dark:border-zinc-700/60">
                <div className="text-[10px] text-zinc-400 font-sans">معدل الالتزام</div>
                <div className="font-bold text-zinc-900 dark:text-zinc-100 mt-0.5">{selectedEmployee.completionRate}%</div>
              </div>
            </div>

            {/* Scrollable Day-by-Day Shifts Breakdown */}
            <div className="flex-1 overflow-auto custom-scroll p-4 space-y-3">
              <div className="text-[11px] font-bold text-zinc-500 mb-2">
                سجل نوبات وبصمات الموظف التفصيلي ({selectedEmployee.shifts.length} سجل مقيد):
              </div>

              {selectedEmployee.shifts.map((shift, sIdx) => {
                const isPrevExit = shift.is_previous_shift;
                return (
                  <div
                    key={sIdx}
                    className={`p-3 rounded-lg border transition-all ${
                      isPrevExit
                        ? 'bg-sky-50/50 dark:bg-sky-950/20 border-sky-200 dark:border-sky-800/60'
                        : shift.status === 'مكتملة'
                          ? 'bg-white dark:bg-zinc-900 border-zinc-200 dark:border-zinc-800'
                          : 'bg-rose-50/30 dark:bg-rose-950/20 border-rose-200 dark:border-rose-900/60'
                    }`}
                  >
                    {/* Shift Header Row */}
                    <div className="flex items-center justify-between pb-2 mb-2 border-b border-zinc-100 dark:border-zinc-800">
                      <div className="flex items-center gap-2">
                        <Calendar className="w-3.5 h-3.5 text-zinc-400" />
                        <span className="font-bold font-mono text-zinc-900 dark:text-zinc-100">
                          {isPrevExit ? 'بصمة خروج لنوبة سابقة' : shift.shift_date}
                        </span>
                        {shift.exit_date_str && (
                          <span className="text-[10px] text-zinc-400 font-mono">
                            (تاريخ الخروج: {shift.exit_date_str})
                          </span>
                        )}
                      </div>

                      {/* Status Tag */}
                      {isPrevExit ? (
                        <span className="px-2 py-0.5 rounded text-[10px] font-bold bg-sky-100 text-sky-800 dark:bg-sky-950 dark:text-sky-300 border border-sky-300 dark:border-sky-800">
                          نوبة سابقة معفاة من احتساب الشهر
                        </span>
                      ) : shift.status === 'مكتملة' ? (
                        <span className="px-2 py-0.5 rounded text-[10px] font-bold bg-emerald-100 text-emerald-800 dark:bg-emerald-950 dark:text-emerald-300">
                          مكتملة ({shift.punches_count} بصمات)
                        </span>
                      ) : (
                        <span className="px-2 py-0.5 rounded text-[10px] font-bold bg-rose-100 text-rose-800 dark:bg-rose-950 dark:text-rose-300">
                          ناقصة (سُجلت {shift.punches_count} فقط)
                        </span>
                      )}
                    </div>

                    {/* Explanatory Note for Previous Shift */}
                    {isPrevExit && (
                      <div className="p-2 mb-2 rounded bg-sky-100/60 dark:bg-sky-900/30 text-[11px] text-sky-900 dark:text-sky-200 flex items-start gap-1.5">
                        <Info className="w-3.5 h-3.5 text-sky-600 shrink-0 mt-0.5" />
                        <span>
                          هذه البصمة تمثل خروجاً لنوبة عمل بدأت في الشهر السابق (مثلاً 31/7 وانتهت 1/8). النظام يتعامل معها بصفة <b>«بصمة خروج لنوبة سابقة»</b> دون إظهار اليوم السابق في جدول الشهر ودون احتسابها ضمن نوبات أو نواقص هذا الشهر.
                        </span>
                      </div>
                    )}

                    {/* Punches Timeline Grid */}
                    <div className="grid grid-cols-2 sm:grid-cols-3 gap-1.5 font-mono text-[11px]">
                      {shift.punches_details.map((p, pIdx) => (
                        <div 
                          key={pIdx}
                          className="p-1.5 rounded bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200/50 dark:border-zinc-700/50 flex flex-col"
                        >
                          <span className="text-[10px] font-sans text-zinc-500 font-bold truncate">
                            {p.period}
                          </span>
                          <span className="font-bold text-zinc-800 dark:text-zinc-200">
                            {p.time}
                          </span>
                        </div>
                      ))}
                    </div>

                    {/* Missing punches alert if any */}
                    {!isPrevExit && shift.missing_periods && shift.missing_periods.length > 0 && (
                      <div className="mt-2 text-[10px] text-rose-600 dark:text-rose-400 font-bold flex items-center gap-1">
                        <AlertCircle className="w-3 h-3 shrink-0" />
                        <span>البصمات الناقصة: {shift.missing_periods.join(' ، ')}</span>
                      </div>
                    )}

                  </div>
                );
              })}
            </div>

            {/* Drawer Footer Actions */}
            <div className="p-3 border-t border-zinc-200 dark:border-zinc-800 bg-zinc-50 dark:bg-zinc-900 flex items-center justify-between shrink-0">
              <button
                onClick={handlePrint}
                className="px-3 py-1.5 rounded border border-zinc-300 dark:border-zinc-700 bg-white dark:bg-zinc-800 text-zinc-800 dark:text-zinc-200 text-xs font-bold flex items-center gap-1.5 hover:bg-zinc-100 transition-colors shadow-xs"
              >
                <Printer className="w-3.5 h-3.5" />
                <span>طباعة إشعار تدقيق الموظف</span>
              </button>

              <button
                onClick={() => setSelectedEmployee(null)}
                className="px-4 py-1.5 rounded bg-zinc-900 dark:bg-zinc-100 text-white dark:text-zinc-900 text-xs font-bold transition-colors shadow-xs"
              >
                إغلاق
              </button>
            </div>

          </div>
        )}

      </div>

    </div>
  );
}
