import React, { useState, useMemo, useRef, useEffect } from 'react';
import * as XLSX from 'xlsx';
import { 
  FileUp, Search, Users, LayoutDashboard, Fingerprint, 
  RefreshCcw, X, ShieldCheck, CheckCircle2, 
  AlertTriangle, FileSpreadsheet, ShieldAlert,
  Clock, ListChecks, ChevronLeft, ChevronRight,
  Settings, SunMedium, MoonStar, Download,
  Printer, BarChart3, Filter, Check, ArrowUpDown,
  Calendar, Info, ExternalLink, ArrowRight
} from 'lucide-react';
import { processExcelData } from './utils/attendanceProcessor';
import { generateSampleDataset } from './utils/demoData';
import { generateBulkAuditSummaryAI, generateDetailedAuditSummaryAI } from './services/geminiService';
import { EmployeeSummary, Rules, ShiftReport } from './types';
import { ATTENDANCE_RULES } from './constants';

type TabView = 'roster' | 'deficiencies' | 'analytics';
type FilterStatus = 'all' | 'compliant' | 'deficient';
type SortField = 'id_asc' | 'id_desc' | 'rate_desc' | 'rate_asc' | 'days_desc' | 'name_asc';
type TableDensity = 'compact' | 'comfortable';

const PAGE_SIZE = 25;

export const App: React.FC = () => {
  const [employees, setEmployees] = useState<EmployeeSummary[]>([]);
  const [rawData, setRawData] = useState<any[]>([]);
  const [fileName, setFileName] = useState<string | null>(null);
  const [rules, setRules] = useState<Rules>(ATTENDANCE_RULES);
  
  // Navigation & View State
  const [activeTab, setActiveTab] = useState<TabView>('roster');
  const [searchTerm, setSearchTerm] = useState('');
  const [filterStatus, setFilterStatus] = useState<FilterStatus>('all');
  const [sortField, setSortField] = useState<SortField>('id_asc');
  const [tableDensity, setTableDensity] = useState<TableDensity>('comfortable');
  const [currentPage, setCurrentPage] = useState(1);
  
  // Inspection & Modals
  const [inspectEmployee, setInspectEmployee] = useState<EmployeeSummary | null>(null);
  const [isSettingsOpen, setIsSettingsOpen] = useState(false);
  const [isProcessing, setIsProcessing] = useState(false);
  const [administrativeOrderNo, setAdministrativeOrderNo] = useState('24 / لسنة 2026');

  // Dark Mode State
  const [isDarkMode, setIsDarkMode] = useState<boolean>(() => {
    return window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches;
  });

  const fileInputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    if (isDarkMode) {
      document.documentElement.classList.add('dark');
    } else {
      document.documentElement.classList.remove('dark');
    }
  }, [isDarkMode]);

  // Load demo data by default on first mount if empty so user immediately sees high-end interface
  useEffect(() => {
    if (employees.length === 0) {
      const demo = generateSampleDataset(rules);
      setEmployees(demo);
      setFileName('سجلات_التدقيق_المركزي.xlsx');
    }
  }, []);

  // System-level Metrics
  const systemMetrics = useMemo(() => {
    if (employees.length === 0) return null;
    const totalStaff = employees.length;
    const compliantCount = employees.filter(e => e.isEligible).length;
    const deficientCount = totalStaff - compliantCount;
    const totalShiftsDays = employees.reduce((sum, e) => sum + e.shifts.length, 0);
    const totalPunchesCount = employees.reduce((sum, e) => sum + e.totalPunches, 0);
    const complianceRate = totalStaff > 0 ? Math.round((compliantCount / totalStaff) * 100) : 0;
    
    // Average punches per shift day
    const avgPunchesPerDay = totalShiftsDays > 0 ? (totalPunchesCount / totalShiftsDays).toFixed(1) : '0';

    return {
      totalStaff,
      compliantCount,
      deficientCount,
      totalShiftsDays,
      totalPunchesCount,
      complianceRate,
      avgPunchesPerDay
    };
  }, [employees]);

  // Filtered & Sorted Staff Roster
  const filteredEmployees = useMemo(() => {
    let list = [...employees];

    if (filterStatus === 'compliant') {
      list = list.filter(e => e.isEligible);
    } else if (filterStatus === 'deficient') {
      list = list.filter(e => !e.isEligible);
    }

    const q = searchTerm.trim().toLowerCase();
    if (q) {
      list = list.filter(e => 
        e.name.toLowerCase().includes(q) || 
        String(e.id).includes(q)
      );
    }

    list.sort((a, b) => {
      switch (sortField) {
        case 'id_desc': {
          const numA = parseInt(String(a.id).replace(/\D/g, '')) || 0;
          const numB = parseInt(String(b.id).replace(/\D/g, '')) || 0;
          return numB - numA;
        }
        case 'rate_desc': return b.completionRate - a.completionRate;
        case 'rate_asc': return a.completionRate - b.completionRate;
        case 'days_desc': return b.shifts.length - a.shifts.length;
        case 'name_asc': return a.name.localeCompare(b.name, 'ar');
        case 'id_asc':
        default: {
          const numA = parseInt(String(a.id).replace(/\D/g, '')) || 0;
          const numB = parseInt(String(b.id).replace(/\D/g, '')) || 0;
          return numA - numB;
        }
      }
    });

    return list;
  }, [employees, filterStatus, searchTerm, sortField]);

  const paginatedEmployees = useMemo(() => {
    const start = (currentPage - 1) * PAGE_SIZE;
    return filteredEmployees.slice(start, start + PAGE_SIZE);
  }, [filteredEmployees, currentPage]);

  const totalPages = Math.ceil(filteredEmployees.length / PAGE_SIZE) || 1;

  useEffect(() => {
    setCurrentPage(1);
  }, [searchTerm, filterStatus, sortField]);

  // Change punch policy (3 vs 5 punches)
  const setPunchPolicy = (count: number) => {
    const updatedRules: Rules = {
      ...rules,
      REQUIRED_PUNCHES_PER_SHIFT: count
    };
    setRules(updatedRules);

    if (rawData.length > 0) {
      const processed = processExcelData(rawData, updatedRules);
      setEmployees(processed);
    } else {
      const sample = generateSampleDataset(updatedRules);
      setEmployees(sample);
    }

    // Keep inspected employee synced
    if (inspectEmployee) {
      const updatedInspected = (rawData.length > 0 
        ? processExcelData(rawData, updatedRules) 
        : generateSampleDataset(updatedRules)
      ).find(e => String(e.id) === String(inspectEmployee.id));
      if (updatedInspected) setInspectEmployee(updatedInspected);
    }
  };

  // File Upload Handler
  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setFileName(file.name);
    setIsProcessing(true);

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const buffer = new Uint8Array(evt.target?.result as ArrayBuffer);
        const workbook = XLSX.read(buffer, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const json = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheetName]);
        
        setRawData(json);
        const processed = processExcelData(json, rules);
        setEmployees(processed);
        setActiveTab('roster');
        setFilterStatus('all');
        setInspectEmployee(null);
      } catch (err) {
        console.error("Failed to parse attendance file:", err);
      } finally {
        setIsProcessing(false);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  // Load Corporate Demo Dataset
  const handleLoadDemo = () => {
    setIsProcessing(true);
    setFileName('سجلات_التدقيق_المركزي.xlsx');
    setTimeout(() => {
      const demo = generateSampleDataset(rules);
      setEmployees(demo);
      setRawData([]);
      setActiveTab('roster');
      setFilterStatus('all');
      setIsProcessing(false);
    }, 150);
  };

  // Export Executive Multi-Sheet Excel Workbook
  const handleExportWorkbook = () => {
    if (employees.length === 0) return;

    // Sheet 1: Master Summary
    const masterSummary = employees.map(emp => ({
      'الرقم الوظيفي': emp.id,
      'اسم الموظف': emp.name,
      'عدد أيام الحضور': emp.shifts.length,
      'إجمالي حركات البصمة': emp.totalPunches,
      'أيام مكتملة (مستوفية)': emp.shifts.filter(s => s.status === 'مكتملة').length,
      'أيام النواقص': emp.shifts.filter(s => s.status === 'ناقصة').length,
      'نسبة الالتزام': `${emp.completionRate}%`,
      'الضابطة المعتمدة': `${rules.REQUIRED_PUNCHES_PER_SHIFT} بصمات`,
      'القرار الإداري': emp.isEligible ? 'مستوفٍ للشروط' : 'به نواقص غير معتمدة'
    }));

    // Sheet 2: Deficiencies Minutes
    const deficienciesLog: any[] = [];
    employees.forEach(emp => {
      emp.shifts.filter(s => s.status === 'ناقصة').forEach(shift => {
        const recorded = shift.punches_details.filter(p => !p.is_duplicate).map(p => p.period);
        const allFive = ["دخول صباحي", "بصمة وسطية", "بصمة مسائية", "بصمة ليلية", "بصمة خروج"];
        const missing = allFive.filter(p => !recorded.includes(p));

        deficienciesLog.push({
          'التاريخ': shift.shift_date,
          'الرقم الوظيفي': emp.id,
          'اسم الموظف': emp.name,
          'عدد البصمات المسجلة': shift.punches_count,
          'الحد الإلزامي المطلوب': rules.REQUIRED_PUNCHES_PER_SHIFT,
          'البصمات المفقودة': missing.join(' - '),
          'أوقات البصمات المسجلة': shift.punches_details.map(p => `${p.period}: ${p.time}`).join(' | ')
        });
      });
    });

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(masterSummary), "الملخص العام للمطابقة");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(deficienciesLog), "محضر النواقص التفصيلي");
    XLSX.writeFile(wb, `محضر_تدقيق_البصمات_المركزي_${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  return (
    <div className="flex flex-col h-screen w-screen bg-slate-100 dark:bg-slate-950 text-slate-900 dark:text-slate-100 font-sans text-xs select-none overflow-hidden">
      
      {/* =========================================================================
          1. EXECUTIVE TOP COMMAND BAR (Authoritative Institutional Header)
      ========================================================================= */}
      <header className="h-14 px-4 lg:px-6 border-b border-slate-200 dark:border-slate-800 bg-white dark:bg-slate-900 flex items-center justify-between gap-4 shrink-0 z-30">
        
        {/* Brand & Administrative Identity */}
        <div className="flex items-center gap-3 shrink-0">
          <div className="w-9 h-9 rounded-md bg-slate-900 dark:bg-slate-800 text-white flex items-center justify-center font-bold tracking-tight shadow-xs border border-slate-700/50">
            <ShieldCheck size={20} className="text-blue-400" />
          </div>
          <div>
            <div className="flex items-center gap-2">
              <span className="font-bold text-sm tracking-tight text-slate-900 dark:text-white">
                بصمتي برو
              </span>
              <span className="text-[10px] font-mono text-slate-400 dark:text-slate-500">
                ·
              </span>
              <span className="text-[11px] font-semibold text-slate-600 dark:text-slate-300">
                منظومة التدقيق المركزي للبصمات
              </span>
            </div>
            <p className="text-[10px] text-slate-500 dark:text-slate-400">
              التدقيق والمطابقة الإدارية وفق الضوابط المعتمدة
            </p>
          </div>
        </div>

        {/* Center Workspace Tabs */}
        <nav className="hidden md:flex items-center gap-1 bg-slate-100 dark:bg-slate-800/80 p-1 rounded-md border border-slate-200 dark:border-slate-800 text-xs font-medium">
          <button
            onClick={() => setActiveTab('roster')}
            className={`px-3.5 py-1.5 rounded transition-all flex items-center gap-2 ${
              activeTab === 'roster'
                ? 'bg-white dark:bg-slate-700 text-slate-900 dark:text-white shadow-xs font-semibold'
                : 'text-slate-600 dark:text-slate-400 hover:text-slate-900 dark:hover:text-white'
            }`}
          >
            <Users size={14} />
            <span>سجل الكادر والمطابقة</span>
          </button>
          
          <button
            onClick={() => setActiveTab('deficiencies')}
            className={`px-3.5 py-1.5 rounded transition-all flex items-center gap-2 ${
              activeTab === 'deficiencies'
                ? 'bg-white dark:bg-slate-700 text-slate-900 dark:text-white shadow-xs font-semibold'
                : 'text-slate-600 dark:text-slate-400 hover:text-slate-900 dark:hover:text-white'
            }`}
          >
            <ListChecks size={14} />
            <span>محضر النواقص والمخالفات</span>
            {systemMetrics && systemMetrics.deficientCount > 0 && (
              <span className="font-mono text-[10px] px-1.5 py-0.2 rounded bg-rose-100 dark:bg-rose-950/70 text-rose-700 dark:text-rose-300 font-semibold">
                {systemMetrics.deficientCount}
              </span>
            )}
          </button>

          <button
            onClick={() => setActiveTab('analytics')}
            className={`px-3.5 py-1.5 rounded transition-all flex items-center gap-2 ${
              activeTab === 'analytics'
                ? 'bg-white dark:bg-slate-700 text-slate-900 dark:text-white shadow-xs font-semibold'
                : 'text-slate-600 dark:text-slate-400 hover:text-slate-900 dark:hover:text-white'
            }`}
          >
            <BarChart3 size={14} />
            <span>التحليل الزمني والورديات</span>
          </button>
        </nav>

        {/* Right Tools & Policy Control */}
        <div className="flex items-center gap-2 shrink-0">
          
          {/* Punch Rule Selector Segment (User Request 2) */}
          <div className="flex items-center bg-slate-100 dark:bg-slate-800/80 p-1 rounded-md border border-slate-200 dark:border-slate-800 text-[11px]">
            <span className="px-2 text-slate-500 dark:text-slate-400 font-medium hidden lg:inline">
              الضابطة المقررة:
            </span>
            <button
              onClick={() => setPunchPolicy(3)}
              className={`px-2.5 py-1 rounded transition-colors font-semibold ${
                rules.REQUIRED_PUNCHES_PER_SHIFT === 3
                  ? 'bg-white dark:bg-slate-700 text-slate-900 dark:text-white shadow-xs font-bold'
                  : 'text-slate-600 dark:text-slate-400 hover:text-slate-900 dark:hover:text-white'
              }`}
              title="تطبيق قرار تقليل البصمات إلى 3 بصمات لكل وردية"
            >
              قرار 3 بصمات
            </button>
            <button
              onClick={() => setPunchPolicy(5)}
              className={`px-2.5 py-1 rounded transition-colors font-semibold ${
                rules.REQUIRED_PUNCHES_PER_SHIFT === 5
                  ? 'bg-white dark:bg-slate-700 text-slate-900 dark:text-white shadow-xs font-bold'
                  : 'text-slate-600 dark:text-slate-400 hover:text-slate-900 dark:hover:text-white'
              }`}
              title="تطبيق قاعدة الـ 5 بصمات الاعتيادية"
            >
              قاعدة 5 بصمات
            </button>
          </div>

          {/* Export Official Multi-Sheet Excel */}
          <button
            onClick={handleExportWorkbook}
            disabled={employees.length === 0}
            className="hidden sm:flex items-center gap-1.5 px-3 py-1.5 rounded-md border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 hover:bg-slate-50 dark:hover:bg-slate-750 text-slate-800 dark:text-slate-200 font-semibold text-xs shadow-xs transition-colors disabled:opacity-50"
            title="تصدير المصنف الإداري المكتمل"
          >
            <Download size={14} className="text-slate-600 dark:text-slate-300" />
            <span>تصدير Excel</span>
          </button>

          {/* Settings Trigger */}
          <button
            onClick={() => setIsSettingsOpen(true)}
            className="p-1.5 rounded-md border border-slate-200 dark:border-slate-800 bg-white dark:bg-slate-800 hover:bg-slate-50 dark:hover:bg-slate-700 text-slate-600 dark:text-slate-300 transition-colors"
            title="إعدادات النوافذ الزمنية للبصمات"
          >
            <Settings size={15} />
          </button>

          {/* Theme Toggle */}
          <button
            onClick={() => setIsDarkMode(!isDarkMode)}
            className="p-1.5 rounded-md border border-slate-200 dark:border-slate-800 bg-white dark:bg-slate-800 hover:bg-slate-50 dark:hover:bg-slate-700 text-slate-600 dark:text-slate-300 transition-colors"
            title={isDarkMode ? 'الوضع الفاتح' : 'الوضع الداكن'}
          >
            {isDarkMode ? <SunMedium size={15} className="text-amber-400" /> : <MoonStar size={15} />}
          </button>

          {/* Primary Upload CTA */}
          <button
            onClick={() => fileInputRef.current?.click()}
            disabled={isProcessing}
            className="flex items-center gap-1.5 px-3 py-1.5 rounded-md bg-slate-900 dark:bg-slate-800 hover:bg-slate-800 dark:hover:bg-slate-700 text-white font-semibold text-xs shadow-xs transition-colors border border-slate-700"
          >
            {isProcessing ? <RefreshCcw size={14} className="animate-spin" /> : <FileUp size={14} />}
            <span>استيراد ملف</span>
          </button>
          <input
            type="file"
            ref={fileInputRef}
            className="hidden"
            accept=".xlsx, .xls"
            onChange={handleFileUpload}
          />

        </div>
      </header>

      {/* =========================================================================
          2. EXECUTIVE SYSTEM KPI STRIP (Continuous Hairline Metrics Bar)
      ========================================================================= */}
      {systemMetrics && (
        <section className="bg-white dark:bg-slate-900 border-b border-slate-200 dark:border-slate-800 px-4 lg:px-6 py-2.5 shrink-0">
          <div className="grid grid-cols-2 md:grid-cols-5 divide-x divide-x-reverse divide-slate-100 dark:divide-slate-800 gap-y-2">
            
            {/* KPI 1: Total Staff */}
            <div className="px-3">
              <span className="text-[11px] text-slate-500 dark:text-slate-400 block font-medium">
                إجمالي الكادر المسجل
              </span>
              <div className="flex items-baseline gap-1.5 mt-0.5">
                <span className="text-lg font-bold font-mono tabular-nums text-slate-900 dark:text-white">
                  {systemMetrics.totalStaff}
                </span>
                <span className="text-[10px] text-slate-400">موظف</span>
              </div>
            </div>

            {/* KPI 2: Total Shift Days (Requested by User) */}
            <div className="px-3">
              <div className="flex items-center gap-1">
                <span className="text-[11px] text-slate-500 dark:text-slate-400 block font-medium">
                  أيام النوبات الإجمالية
                </span>
                <span className="text-[9px] font-mono px-1 rounded bg-blue-50 dark:bg-blue-950/60 text-blue-700 dark:text-blue-300 border border-blue-200/50 dark:border-blue-800/50">
                  محسوبة
                </span>
              </div>
              <div className="flex items-baseline gap-1.5 mt-0.5">
                <span className="text-lg font-bold font-mono tabular-nums text-slate-900 dark:text-white">
                  {systemMetrics.totalShiftsDays}
                </span>
                <span className="text-[10px] text-slate-400">يوم نوبة</span>
              </div>
            </div>

            {/* KPI 3: Total Punch Actions */}
            <div className="px-3">
              <span className="text-[11px] text-slate-500 dark:text-slate-400 block font-medium">
                حركات البصمة المسجلة
              </span>
              <div className="flex items-baseline gap-1.5 mt-0.5">
                <span className="text-lg font-bold font-mono tabular-nums text-slate-900 dark:text-white">
                  {systemMetrics.totalPunchesCount}
                </span>
                <span className="text-[10px] text-slate-400">بصمة</span>
                <span className="text-[10px] font-mono text-slate-400 mr-1">
                  (معدل {systemMetrics.avgPunchesPerDay} / نوبة)
                </span>
              </div>
            </div>

            {/* KPI 4: Compliant Staff */}
            <div className="px-3">
              <div className="flex items-center justify-between">
                <span className="text-[11px] text-slate-500 dark:text-slate-400 block font-medium">
                  سجلات مكتملة (مستوفية)
                </span>
                <span className="text-[10px] font-mono font-semibold text-emerald-600 dark:text-emerald-400">
                  {systemMetrics.complianceRate}%
                </span>
              </div>
              <div className="flex items-baseline gap-1.5 mt-0.5">
                <span className="text-lg font-bold font-mono tabular-nums text-emerald-600 dark:text-emerald-400">
                  {systemMetrics.compliantCount}
                </span>
                <span className="text-[10px] text-slate-400">موظف منضبط</span>
              </div>
            </div>

            {/* KPI 5: Deficient Records */}
            <div className="px-3">
              <div className="flex items-center justify-between">
                <span className="text-[11px] text-slate-500 dark:text-slate-400 block font-medium">
                  سجلات بها نواقص
                </span>
                {systemMetrics.deficientCount > 0 && (
                  <span className="text-[10px] font-mono font-semibold text-rose-600 dark:text-rose-400">
                    {100 - systemMetrics.complianceRate}%
                  </span>
                )}
              </div>
              <div className="flex items-baseline gap-1.5 mt-0.5">
                <span className="text-lg font-bold font-mono tabular-nums text-rose-600 dark:text-rose-400">
                  {systemMetrics.deficientCount}
                </span>
                <span className="text-[10px] text-slate-400">موظف للمراجعة</span>
              </div>
            </div>

          </div>
        </section>
      )}

      {/* =========================================================================
          3. MAIN WORKSPACE VIEWPORT (Tab Routing & High Density Tables)
      ========================================================================= */}
      <main className="flex-1 overflow-hidden flex flex-row min-w-0 relative">
        
        {/* Left Primary Data Content */}
        <div className="flex-1 overflow-y-auto custom-scroll flex flex-col min-w-0 p-4 lg:p-6 space-y-4">
          
          {/* TAB 1: MASTER STAFF ATTENDANCE ROSTER */}
          {activeTab === 'roster' && (
            <div className="flex-1 flex flex-col space-y-3">
              
              {/* Context Toolbar: Directive Header & Controls */}
              <div className="flex flex-col md:flex-row items-stretch md:items-center justify-between gap-3 bg-white dark:bg-slate-900 p-3 rounded-lg border border-slate-200 dark:border-slate-800 shadow-xs">
                
                {/* Search & Order Reference */}
                <div className="flex items-center gap-2 flex-1 max-w-lg">
                  <div className="relative flex-1">
                    <Search className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400" size={14} />
                    <input
                      type="text"
                      placeholder="بحث بالاسم أو الرقم الوظيفي..."
                      value={searchTerm}
                      onChange={(e) => setSearchTerm(e.target.value)}
                      className="w-full bg-slate-50 dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded pr-9 pl-7 py-1.5 text-xs text-slate-900 dark:text-white focus:outline-none focus:border-slate-500 transition-colors placeholder:text-slate-400"
                    />
                    {searchTerm && (
                      <button
                        onClick={() => setSearchTerm('')}
                        className="absolute left-2.5 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600"
                      >
                        <X size={12} />
                      </button>
                    )}
                  </div>

                  <div className="hidden sm:flex items-center gap-1.5 px-2.5 py-1.5 rounded bg-slate-50 dark:bg-slate-800 border border-slate-200 dark:border-slate-700 text-[11px] text-slate-600 dark:text-slate-300">
                    <span className="text-slate-400">الأمر:</span>
                    <input 
                      type="text"
                      value={administrativeOrderNo}
                      onChange={(e) => setAdministrativeOrderNo(e.target.value)}
                      className="bg-transparent border-none p-0 w-24 font-mono font-semibold focus:outline-none text-slate-800 dark:text-slate-200 text-[11px]"
                      title="رقم الأمر الإداري / التعميم"
                    />
                  </div>
                </div>

                {/* Filter & Sort Controls */}
                <div className="flex items-center gap-2 flex-wrap">
                  
                  {/* Status Filters */}
                  <div className="flex items-center bg-slate-100 dark:bg-slate-800 p-0.5 rounded border border-slate-200 dark:border-slate-700 text-xs">
                    <button
                      onClick={() => setFilterStatus('all')}
                      className={`px-2.5 py-1 rounded transition-colors font-medium ${
                        filterStatus === 'all'
                          ? 'bg-white dark:bg-slate-700 text-slate-900 dark:text-white shadow-xs font-semibold'
                          : 'text-slate-600 dark:text-slate-400 hover:text-slate-900'
                      }`}
                    >
                      الكل ({employees.length})
                    </button>
                    <button
                      onClick={() => setFilterStatus('compliant')}
                      className={`px-2.5 py-1 rounded transition-colors font-medium ${
                        filterStatus === 'compliant'
                          ? 'bg-white dark:bg-slate-700 text-emerald-700 dark:text-emerald-300 shadow-xs font-semibold'
                          : 'text-slate-600 dark:text-slate-400 hover:text-slate-900'
                      }`}
                    >
                      مستوفٍ ({employees.filter(e => e.isEligible).length})
                    </button>
                    <button
                      onClick={() => setFilterStatus('deficient')}
                      className={`px-2.5 py-1 rounded transition-colors font-medium ${
                        filterStatus === 'deficient'
                          ? 'bg-white dark:bg-slate-700 text-rose-700 dark:text-rose-300 shadow-xs font-semibold'
                          : 'text-slate-600 dark:text-slate-400 hover:text-slate-900'
                      }`}
                    >
                      به نواقص ({employees.filter(e => !e.isEligible).length})
                    </button>
                  </div>

                  {/* Sort Field */}
                  <select
                    value={sortField}
                    onChange={(e) => setSortField(e.target.value as SortField)}
                    className="bg-slate-50 dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded px-2.5 py-1.5 text-xs text-slate-700 dark:text-slate-300 focus:outline-none cursor-pointer"
                  >
                    <option value="id_asc">الرقم الوظيفي (تصاعدي)</option>
                    <option value="id_desc">الرقم الوظيفي (تنازلي)</option>
                    <option value="rate_desc">الأعلى التزاماً</option>
                    <option value="rate_asc">الأقل التزاماً</option>
                    <option value="days_desc">الأكثر أيام عمل</option>
                    <option value="name_asc">الاسم (أبجدياً)</option>
                  </select>

                  {/* Density Switch */}
                  <button
                    onClick={() => setTableDensity(tableDensity === 'comfortable' ? 'compact' : 'comfortable')}
                    className="p-1.5 rounded border border-slate-200 dark:border-slate-700 bg-slate-50 dark:bg-slate-800 hover:bg-slate-100 text-slate-600 dark:text-slate-300"
                    title={tableDensity === 'comfortable' ? 'التبديل إلى العرض المكثف' : 'التبديل إلى العرض المريح'}
                  >
                    <ArrowUpDown size={14} />
                  </button>

                  {/* Quick Reload Sample */}
                  <button
                    onClick={handleLoadDemo}
                    className="px-2.5 py-1.5 rounded border border-slate-200 dark:border-slate-700 bg-slate-50 dark:bg-slate-800 hover:bg-slate-100 text-slate-600 dark:text-slate-300 font-medium text-xs transition-colors"
                    title="إعادة تحميل العينة المؤسسية"
                  >
                    عينة نموذجية
                  </button>

                </div>
              </div>

              {/* High-Density Data Grid */}
              <div className="bg-white dark:bg-slate-900 rounded-lg border border-slate-200 dark:border-slate-800 overflow-hidden shadow-xs flex-1 flex flex-col">
                <div className="overflow-x-auto flex-1 custom-scroll">
                  <table className="w-full text-right border-collapse text-xs">
                    <thead>
                      <tr className="bg-slate-50 dark:bg-slate-850/80 border-b border-slate-200 dark:border-slate-800 text-slate-500 dark:text-slate-400 font-semibold text-[11px] select-none sticky top-0 z-10">
                        <th className="py-2.5 px-4 w-24">الرقم الوظيفي</th>
                        <th className="py-2.5 px-4 min-w-[180px]">اسم الموظف</th>
                        <th className="py-2.5 px-4 text-center w-28">أيام الحضور</th>
                        <th className="py-2.5 px-4 text-center w-28">إجمالي البصمات</th>
                        <th className="py-2.5 px-4 text-center w-28">أيام النواقص</th>
                        <th className="py-2.5 px-4 text-center w-40">نسبة الالتزام</th>
                        <th className="py-2.5 px-4 text-center w-32">القرار الإداري</th>
                        <th className="py-2.5 px-4 text-left w-24">الإجراء</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100 dark:divide-slate-800/80 font-normal">
                      {paginatedEmployees.length === 0 ? (
                        <tr>
                          <td colSpan={8} className="py-16 text-center text-slate-400 dark:text-slate-500">
                            لا توجد سجلات تطابق معايير البحث والفلترة المحددة.
                          </td>
                        </tr>
                      ) : (
                        paginatedEmployees.map((emp) => {
                          const incompleteDays = emp.shifts.filter(s => s.status === 'ناقصة').length;
                          const isSelected = inspectEmployee?.id === emp.id;

                          return (
                            <tr
                              key={emp.id}
                              onClick={() => setInspectEmployee(emp)}
                              className={`transition-colors cursor-pointer ${
                                isSelected
                                  ? 'bg-blue-50/60 dark:bg-slate-800/90 font-medium'
                                  : 'hover:bg-slate-50/80 dark:hover:bg-slate-850/50'
                              } ${tableDensity === 'compact' ? 'h-9' : 'h-12'}`}
                            >
                              {/* ACNo ID */}
                              <td className="py-2 px-4 font-mono font-semibold text-slate-600 dark:text-slate-400 tabular-nums">
                                #{emp.id}
                              </td>

                              {/* Name */}
                              <td className="py-2 px-4">
                                <div className="flex items-center gap-2">
                                  <span className="font-semibold text-slate-900 dark:text-slate-100">
                                    {emp.name}
                                  </span>
                                </div>
                              </td>

                              {/* Days count (User Request 2) */}
                              <td className="py-2 px-4 text-center font-mono tabular-nums text-slate-700 dark:text-slate-300">
                                <span className="font-semibold">{emp.shifts.length}</span>
                                <span className="text-[10px] text-slate-400 mr-1">يوم</span>
                              </td>

                              {/* Total punches */}
                              <td className="py-2 px-4 text-center font-mono tabular-nums font-semibold text-slate-800 dark:text-slate-200">
                                {emp.totalPunches}
                              </td>

                              {/* Deficient days */}
                              <td className="py-2 px-4 text-center font-mono tabular-nums">
                                {incompleteDays > 0 ? (
                                  <span className="text-rose-600 dark:text-rose-400 font-semibold">
                                    {incompleteDays} يوم
                                  </span>
                                ) : (
                                  <span className="text-slate-300 dark:text-slate-600">—</span>
                                )}
                              </td>

                              {/* Compliance meter */}
                              <td className="py-2 px-4 text-center">
                                <div className="flex items-center justify-center gap-2">
                                  <div className="w-20 h-1.5 bg-slate-100 dark:bg-slate-800 rounded-full overflow-hidden shrink-0">
                                    <div
                                      className={`h-full rounded-full transition-all ${
                                        emp.isEligible ? 'bg-emerald-500' : 'bg-amber-500'
                                      }`}
                                      style={{ width: `${Math.min(100, emp.completionRate)}%` }}
                                    />
                                  </div>
                                  <span className="font-mono font-semibold tabular-nums text-slate-700 dark:text-slate-300 text-[11px] w-9 text-left">
                                    {emp.completionRate}%
                                  </span>
                                </div>
                              </td>

                              {/* Administrative Status */}
                              <td className="py-2 px-4 text-center">
                                <span className={`inline-block px-2 py-0.5 rounded text-[11px] font-semibold border ${
                                  emp.isEligible
                                    ? 'bg-emerald-50 text-emerald-700 border-emerald-200/80 dark:bg-emerald-950/40 dark:text-emerald-300 dark:border-emerald-800/40'
                                    : 'bg-rose-50 text-rose-700 border-rose-200/80 dark:bg-rose-950/40 dark:text-rose-300 dark:border-rose-800/40'
                                }`}>
                                  {emp.isEligible ? 'مستوفٍ للشروط' : 'به نواقص'}
                                </span>
                              </td>

                              {/* Action */}
                              <td className="py-2 px-4 text-left">
                                <button
                                  onClick={(e) => {
                                    e.stopPropagation();
                                    setInspectEmployee(emp);
                                  }}
                                  className="px-2.5 py-1 text-xs rounded border border-slate-200 dark:border-slate-700 hover:bg-slate-100 dark:hover:bg-slate-800 text-slate-700 dark:text-slate-300 font-medium transition-colors"
                                >
                                  فحص السجل
                                </button>
                              </td>
                            </tr>
                          );
                        })
                      )}
                    </tbody>
                  </table>
                </div>

                {/* Pagination Strip */}
                <div className="h-11 px-4 border-t border-slate-200 dark:border-slate-800 bg-slate-50/70 dark:bg-slate-900/70 flex items-center justify-between text-xs text-slate-500 dark:text-slate-400 shrink-0">
                  <div>
                    عرض <strong className="font-mono text-slate-800 dark:text-white">{filteredEmployees.length}</strong> موظف إجمالاً 
                    (صفحة <strong className="font-mono text-slate-800 dark:text-white">{currentPage}</strong> من <strong className="font-mono text-slate-800 dark:text-white">{totalPages}</strong>)
                  </div>
                  <div className="flex items-center gap-1.5">
                    <button
                      onClick={() => setCurrentPage(p => Math.max(1, p - 1))}
                      disabled={currentPage === 1}
                      className="px-2.5 py-1 rounded border border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-800 disabled:opacity-40 hover:bg-slate-100 transition-colors"
                    >
                      السابق
                    </button>
                    <button
                      onClick={() => setCurrentPage(p => Math.min(totalPages, p + 1))}
                      disabled={currentPage === totalPages}
                      className="px-2.5 py-1 rounded border border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-800 disabled:opacity-40 hover:bg-slate-100 transition-colors"
                    >
                      التالي
                    </button>
                  </div>
                </div>

              </div>

            </div>
          )}

          {/* TAB 2: CENTRAL DEFICIENCY MINUTES & IRREGULARITIES */}
          {activeTab === 'deficiencies' && (
            <div className="flex-1 flex flex-col space-y-4">
              
              {/* Memorandum Notice Header */}
              <div className="p-4 bg-white dark:bg-slate-900 rounded-lg border border-slate-200 dark:border-slate-800 shadow-xs flex flex-col sm:flex-row items-start sm:items-center justify-between gap-3">
                <div>
                  <div className="flex items-center gap-2">
                    <ShieldAlert size={18} className="text-rose-600 dark:text-rose-400" />
                    <h2 className="font-bold text-sm text-slate-900 dark:text-white">
                      محضر كشف النواقص والمخالفات الإدارية
                    </h2>
                  </div>
                  <p className="text-slate-500 dark:text-slate-400 text-xs mt-1">
                    حصر شامل لحالات عدم اكتمال البصمات وفق الضابطة المقررة ({rules.REQUIRED_PUNCHES_PER_SHIFT} بصمات / وردية) - رقم التعميم: {administrativeOrderNo}
                  </p>
                </div>
                
                <button
                  onClick={handleExportWorkbook}
                  className="px-3 py-1.5 rounded-md border border-slate-300 dark:border-slate-700 bg-slate-50 dark:bg-slate-800 hover:bg-slate-100 text-slate-800 dark:text-slate-200 font-semibold text-xs flex items-center gap-1.5 transition-colors"
                >
                  <Download size={14} />
                  <span>تصدير كشف النواقص (Excel)</span>
                </button>
              </div>

              {/* Deficient Employees Roster */}
              <div className="grid grid-cols-1 md:grid-cols-2 gap-3 flex-1 overflow-y-auto">
                {employees.filter(e => !e.isEligible).length === 0 ? (
                  <div className="col-span-2 py-16 text-center bg-white dark:bg-slate-900 rounded-lg border border-slate-200 dark:border-slate-800 text-slate-400">
                    <CheckCircle2 size={32} className="mx-auto text-emerald-500 mb-2 opacity-80" />
                    <p className="font-semibold text-slate-700 dark:text-slate-300 text-sm">
                      كافة موظفي الكادر منضبطون ومستوفون لجميع البصمات المقررة.
                    </p>
                    <p className="text-xs text-slate-400 mt-1">
                      لا توجد أي نواقص مسجلة وفق قاعدة {rules.REQUIRED_PUNCHES_PER_SHIFT} بصمات.
                    </p>
                  </div>
                ) : (
                  employees.filter(e => !e.isEligible).map((emp) => {
                    const deficientShifts = emp.shifts.filter(s => s.status === 'ناقصة');
                    return (
                      <div
                        key={emp.id}
                        className="bg-white dark:bg-slate-900 rounded-lg border border-slate-200 dark:border-slate-800 p-4 shadow-xs flex flex-col justify-between"
                      >
                        <div>
                          <div className="flex items-center justify-between pb-2.5 mb-3 border-b border-slate-100 dark:border-slate-800">
                            <div>
                              <div className="flex items-center gap-2">
                                <h3 className="font-bold text-xs text-slate-900 dark:text-white">
                                  {emp.name}
                                </h3>
                                <span className="font-mono text-slate-400 text-[10px]">
                                  #{emp.id}
                                </span>
                              </div>
                              <span className="text-[11px] text-slate-500 dark:text-slate-400">
                                إجمالي {deficientShifts.length} يوم عمل بها نواقص
                              </span>
                            </div>

                            <button
                              onClick={() => setInspectEmployee(emp)}
                              className="px-2.5 py-1 text-xs rounded border border-slate-200 dark:border-slate-700 hover:bg-slate-100 dark:hover:bg-slate-800 text-slate-700 dark:text-slate-300 font-medium transition-colors"
                            >
                              مراجعة السجل
                            </button>
                          </div>

                          {/* List of Shifts with Missing Slots */}
                          <div className="space-y-2">
                            {deficientShifts.map((shift, idx) => {
                              const found = shift.punches_details.filter(p => !p.is_duplicate).map(p => p.period);
                              const allPossible = ["دخول صباحي", "بصمة وسطية", "بصمة مسائية", "بصمة ليلية", "بصمة خروج"];
                              const missing = allPossible.filter(p => !found.includes(p));

                              return (
                                <div
                                  key={idx}
                                  className="p-2.5 rounded bg-slate-50 dark:bg-slate-800/60 border border-slate-100 dark:border-slate-800 flex flex-col sm:flex-row sm:items-center justify-between gap-1 text-[11px]"
                                >
                                  <div className="flex items-center gap-2">
                                    <span className="font-mono font-semibold text-slate-700 dark:text-slate-300">
                                      {shift.shift_date}
                                    </span>
                                    <span className="text-slate-400">
                                      ({shift.punches_count} من {rules.REQUIRED_PUNCHES_PER_SHIFT} بصمات)
                                    </span>
                                  </div>

                                  <div className="flex flex-wrap gap-1">
                                    {missing.map((m) => (
                                      <span
                                        key={m}
                                        className="px-1.5 py-0.2 rounded text-[10px] bg-rose-50 dark:bg-rose-950/60 text-rose-700 dark:text-rose-300 border border-rose-200/70 dark:border-rose-900/50 font-medium"
                                      >
                                        نقص {m}
                                      </span>
                                    ))}
                                  </div>
                                </div>
                              );
                            })}
                          </div>
                        </div>

                        <div className="mt-3 pt-2 border-t border-slate-100 dark:border-slate-800 flex items-center justify-between text-[11px] text-slate-500">
                          <span>نسبة الالتزام: <strong className="font-mono text-slate-700 dark:text-slate-300">{emp.completionRate}%</strong></span>
                          <span>معدل الأداء: <strong className="font-mono text-slate-700 dark:text-slate-300">{emp.performanceScore}%</strong></span>
                        </div>
                      </div>
                    );
                  })
                )}
              </div>

            </div>
          )}

          {/* TAB 3: SHIFT & TIMING DISTRIBUTION (Enterprise Analytics) */}
          {activeTab === 'analytics' && (
            <div className="space-y-4">
              
              <div className="p-4 bg-white dark:bg-slate-900 rounded-lg border border-slate-200 dark:border-slate-800 shadow-xs">
                <h2 className="font-bold text-sm text-slate-900 dark:text-white">
                  التحليل الإحصائي لتوزيع الورديات والبصمات
                </h2>
                <p className="text-slate-500 dark:text-slate-400 text-xs mt-0.5">
                  بيانات تفصيلية توضح مقارنة تطبيق ضابطة الـ 3 بصمات مقابل الـ 5 بصمات وتوزيع حركات الحضور
                </p>
              </div>

              {/* Policy Impact Comparison Card */}
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                
                {/* 3 vs 5 Impact */}
                <div className="bg-white dark:bg-slate-900 rounded-lg border border-slate-200 dark:border-slate-800 p-4 shadow-xs">
                  <h3 className="font-bold text-xs text-slate-800 dark:text-slate-200 mb-3 pb-2 border-b border-slate-100 dark:border-slate-800">
                    أثر التبديل بين قرار (3 بصمات) و (5 بصمات)
                  </h3>
                  
                  <div className="space-y-3 text-xs">
                    <div className="flex items-center justify-between p-2.5 rounded bg-slate-50 dark:bg-slate-800/60 border border-slate-100 dark:border-slate-800">
                      <div>
                        <span className="font-semibold block text-slate-800 dark:text-slate-200">
                          نظام 3 بصمات (أمر التخفيض)
                        </span>
                        <span className="text-[11px] text-slate-500">
                          يتطلب: دخول صباحي + بصمة وسطية + بصمة خروج
                        </span>
                      </div>
                      <span className="font-mono font-bold text-blue-600 dark:text-blue-400 text-sm">
                        {employees.filter(e => {
                          // Check if compliant under 3
                          return e.shifts.every(s => s.punches_count >= 3);
                        }).length} موظف مستوفٍ
                      </span>
                    </div>

                    <div className="flex items-center justify-between p-2.5 rounded bg-slate-50 dark:bg-slate-800/60 border border-slate-100 dark:border-slate-800">
                      <div>
                        <span className="font-semibold block text-slate-800 dark:text-slate-200">
                          نظام 5 بصمات (الاعتيادي الكامل)
                        </span>
                        <span className="text-[11px] text-slate-500">
                          يتطلب كامل النوبات (صباحي + وسطي + مسائي + ليلي + خروج)
                        </span>
                      </div>
                      <span className="font-mono font-bold text-blue-600 dark:text-blue-400 text-sm">
                        {employees.filter(e => {
                          return e.shifts.every(s => s.punches_count >= 5);
                        }).length} موظف مستوفٍ
                      </span>
                    </div>

                    <p className="text-[11px] text-slate-500 leading-relaxed bg-slate-50 dark:bg-slate-800/40 p-2.5 rounded border border-slate-200/60 dark:border-slate-700/60">
                      <Info size={13} className="inline ml-1 text-blue-500" />
                      يمكنك التحويل الفوري بين الضابطتين بنقرة زر في الشريط العلوي لتحديث نسب الانضباط وقرار الاستحقاق فوراً.
                    </p>
                  </div>
                </div>

                {/* Period Windows Configuration Summary */}
                <div className="bg-white dark:bg-slate-900 rounded-lg border border-slate-200 dark:border-slate-800 p-4 shadow-xs">
                  <div className="flex items-center justify-between mb-3 pb-2 border-b border-slate-100 dark:border-slate-800">
                    <h3 className="font-bold text-xs text-slate-800 dark:text-slate-200">
                      النوافذ الزمنية المعتمدة للفترات
                    </h3>
                    <button
                      onClick={() => setIsSettingsOpen(true)}
                      className="text-xs text-blue-600 hover:underline font-semibold"
                    >
                      تعديل المواعيد
                    </button>
                  </div>

                  <div className="space-y-2 text-xs">
                    {Object.entries(rules.TIME_PERIODS).map(([periodName, times]) => (
                      <div
                        key={periodName}
                        className="flex items-center justify-between p-2 rounded bg-slate-50 dark:bg-slate-800/60 border border-slate-100 dark:border-slate-800"
                      >
                        <span className="font-medium text-slate-700 dark:text-slate-300">{periodName}</span>
                        <span className="font-mono font-semibold tabular-nums text-slate-800 dark:text-slate-200 text-[11px]">
                          من {times.start} إلى {times.end}
                        </span>
                      </div>
                    ))}
                  </div>
                </div>

              </div>

            </div>
          )}

        </div>

        {/* =========================================================================
            4. SLIDE-OVER INSPECTION DRAWER (Executive Employee Dossier)
        ========================================================================= */}
        {inspectEmployee && (
          <aside className="w-full sm:w-[460px] lg:w-[480px] bg-white dark:bg-slate-900 border-r border-slate-200 dark:border-slate-800 shadow-xl flex flex-col z-20 shrink-0 h-full overflow-hidden">
            
            {/* Drawer Header */}
            <div className="p-4 border-b border-slate-200 dark:border-slate-800 bg-slate-50/70 dark:bg-slate-850/80 flex items-center justify-between shrink-0">
              <div className="flex items-center gap-2.5">
                <div className="w-8 h-8 rounded bg-slate-900 dark:bg-slate-800 text-white flex items-center justify-center font-bold text-xs border border-slate-700">
                  {inspectEmployee.name.charAt(0)}
                </div>
                <div>
                  <h3 className="font-bold text-sm text-slate-900 dark:text-white leading-tight">
                    {inspectEmployee.name}
                  </h3>
                  <div className="flex items-center gap-2 text-[11px] font-mono text-slate-500 mt-0.5">
                    <span>الرقم: #{inspectEmployee.id}</span>
                    <span>·</span>
                    <span>الأيام: {inspectEmployee.shifts.length} يوم</span>
                  </div>
                </div>
              </div>

              <div className="flex items-center gap-1.5">
                <button
                  onClick={() => window.print()}
                  className="p-1.5 rounded border border-slate-200 dark:border-slate-700 text-slate-600 dark:text-slate-300 hover:bg-slate-100 transition-colors"
                  title="طباعة محضر الموظف"
                >
                  <Printer size={15} />
                </button>
                <button
                  onClick={() => setInspectEmployee(null)}
                  className="p-1.5 rounded border border-slate-200 dark:border-slate-700 text-slate-500 hover:bg-slate-100 transition-colors"
                  title="إغلاق الملف"
                >
                  <X size={15} />
                </button>
              </div>
            </div>

            {/* Drawer Body */}
            <div className="flex-1 overflow-y-auto custom-scroll p-4 space-y-4">
              
              {/* Executive Determination Card */}
              <div className={`p-3.5 rounded-lg border ${
                inspectEmployee.isEligible
                  ? 'bg-emerald-50/60 dark:bg-emerald-950/30 border-emerald-200 dark:border-emerald-800/60'
                  : 'bg-rose-50/60 dark:bg-rose-950/30 border-rose-200 dark:border-rose-800/60'
              }`}>
                <div className="flex items-center justify-between">
                  <span className="font-semibold text-xs text-slate-800 dark:text-slate-200">
                    القرار الرقابي المعتمد:
                  </span>
                  <span className={`px-2 py-0.5 rounded text-[11px] font-bold border ${
                    inspectEmployee.isEligible
                      ? 'bg-emerald-100 dark:bg-emerald-900 text-emerald-800 dark:text-emerald-200 border-emerald-300'
                      : 'bg-rose-100 dark:bg-rose-900 text-rose-800 dark:text-rose-200 border-rose-300'
                  }`}>
                    {inspectEmployee.isEligible ? 'مستوفٍ للشروط ومستحق' : 'توجد نواقص تستوجب المراجعة'}
                  </span>
                </div>
                <div className="grid grid-cols-3 gap-2 text-center mt-3 pt-2.5 border-t border-slate-200/60 dark:border-slate-800/60 text-xs">
                  <div>
                    <span className="text-[10px] text-slate-400 block">نسبة الالتزام</span>
                    <span className="font-bold font-mono text-slate-800 dark:text-slate-100 text-sm">
                      {inspectEmployee.completionRate}%
                    </span>
                  </div>
                  <div>
                    <span className="text-[10px] text-slate-400 block">إجمالي البصمات</span>
                    <span className="font-bold font-mono text-slate-800 dark:text-slate-100 text-sm">
                      {inspectEmployee.totalPunches}
                    </span>
                  </div>
                  <div>
                    <span className="text-[10px] text-slate-400 block">الضابطة</span>
                    <span className="font-bold font-mono text-slate-800 dark:text-slate-100 text-sm">
                      {rules.REQUIRED_PUNCHES_PER_SHIFT} بصمات
                    </span>
                  </div>
                </div>
              </div>

              {/* Day-by-Day Shift Ledger with Visual 24h Timeline */}
              <div>
                <div className="flex items-center justify-between mb-2">
                  <h4 className="font-bold text-xs text-slate-800 dark:text-slate-200">
                    سجل النوبات والورديات بالتواريخ
                  </h4>
                  <span className="text-[10px] text-slate-400 font-mono">
                    {inspectEmployee.shifts.length} يوم عمل
                  </span>
                </div>

                <div className="space-y-3">
                  {inspectEmployee.shifts.map((shift, idx) => (
                    <div
                      key={idx}
                      className={`p-3 rounded-lg border text-xs ${
                        shift.status === 'مكتملة'
                          ? 'bg-slate-50/70 dark:bg-slate-850/60 border-slate-200 dark:border-slate-800'
                          : 'bg-rose-50/30 dark:bg-rose-950/20 border-rose-200 dark:border-rose-900/60'
                      }`}
                    >
                      {/* Shift Date & Status */}
                      <div className="flex items-center justify-between mb-2">
                        <div className="flex items-center gap-1.5">
                          <Calendar size={13} className="text-slate-400" />
                          <span className="font-mono font-bold text-slate-900 dark:text-white">
                            {shift.shift_date}
                          </span>
                        </div>
                        <span className={`px-2 py-0.5 rounded text-[10px] font-semibold border ${
                          shift.status === 'مكتملة'
                            ? 'bg-emerald-50 text-emerald-700 border-emerald-200 dark:bg-emerald-950/40 dark:text-emerald-300'
                            : 'bg-rose-50 text-rose-700 border-rose-200 dark:bg-rose-950/40 dark:text-rose-300'
                        }`}>
                          {shift.status} ({shift.punches_count} من {rules.REQUIRED_PUNCHES_PER_SHIFT})
                        </span>
                      </div>

                      {/* Visual 24-Hour Timeline Bar (Senior Design Feature) */}
                      <div className="my-2.5 px-1">
                        <div className="text-[9px] text-slate-400 flex justify-between font-mono mb-1">
                          <span>07:00</span>
                          <span>12:00</span>
                          <span>18:00</span>
                          <span>22:00</span>
                          <span>08:00</span>
                        </div>
                        <div className="h-2 w-full bg-slate-200 dark:bg-slate-800 rounded-full relative overflow-hidden flex">
                          {/* Segment markers */}
                          <div className="w-[20%] h-full border-r border-slate-300 dark:border-slate-700 bg-blue-100/40 dark:bg-blue-950/40" title="دخول صباحي" />
                          <div className="w-[25%] h-full border-r border-slate-300 dark:border-slate-700 bg-blue-100/40 dark:bg-blue-950/40" title="بصمة وسطية" />
                          <div className="w-[20%] h-full border-r border-slate-300 dark:border-slate-700 bg-blue-100/40 dark:bg-blue-950/40" title="بصمة مسائية" />
                          <div className="w-[20%] h-full border-r border-slate-300 dark:border-slate-700 bg-blue-100/40 dark:bg-blue-950/40" title="بصمة ليلية" />
                          <div className="w-[15%] h-full bg-blue-100/40 dark:bg-blue-950/40" title="بصمة خروج" />
                        </div>
                      </div>

                      {/* 5-Slot Breakdown */}
                      <div className="grid grid-cols-5 gap-1 text-center mt-2">
                        {["دخول صباحي", "بصمة وسطية", "بصمة مسائية", "بصمة ليلية", "بصمة خروج"].map((pName) => {
                          const p = shift.punches_details.find(d => d.period === pName && !d.is_duplicate);
                          return (
                            <div
                              key={pName}
                              className={`p-1.5 rounded border text-[10px] ${
                                p
                                  ? 'bg-white dark:bg-slate-800 border-slate-200 dark:border-slate-700 font-semibold'
                                  : 'bg-slate-100/40 dark:bg-slate-900/40 border-transparent text-slate-400'
                              }`}
                            >
                              <span className="text-[9px] text-slate-400 block truncate">{pName}</span>
                              <span className="font-mono tabular-nums text-slate-700 dark:text-slate-300">
                                {p ? p.time : '—'}
                              </span>
                            </div>
                          );
                        })}
                      </div>

                    </div>
                  ))}
                </div>
              </div>

            </div>

            {/* Drawer Footer with Official Print Action */}
            <div className="p-3 border-t border-slate-200 dark:border-slate-800 bg-slate-50 dark:bg-slate-850/80 shrink-0 flex items-center justify-between">
              <span className="text-[11px] text-slate-500">
                الأمر المرجعي: {administrativeOrderNo}
              </span>
              <button
                onClick={() => window.print()}
                className="px-3 py-1.5 rounded bg-slate-900 dark:bg-slate-800 hover:bg-slate-800 dark:hover:bg-slate-700 text-white font-semibold text-xs flex items-center gap-1.5 transition-colors border border-slate-700"
              >
                <Printer size={13} />
                <span>طباعة إشعار الموظف</span>
              </button>
            </div>

          </aside>
        )}

      </main>

      {/* =========================================================================
          5. SETTINGS MODAL (Shift Period Time Windows & Directives)
      ========================================================================= */}
      {isSettingsOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/50 backdrop-blur-xs">
          <div className="bg-white dark:bg-slate-900 w-full max-w-lg rounded-xl overflow-hidden shadow-2xl border border-slate-200 dark:border-slate-800">
            
            <div className="px-5 py-3.5 border-b border-slate-200 dark:border-slate-800 bg-slate-50 dark:bg-slate-850 flex items-center justify-between">
              <div className="flex items-center gap-2">
                <Settings size={16} className="text-slate-700 dark:text-slate-300" />
                <h3 className="font-bold text-sm text-slate-900 dark:text-white">
                  إعدادات وضوابط الورديات وفترات البصمات
                </h3>
              </div>
              <button
                onClick={() => setIsSettingsOpen(false)}
                className="p-1 rounded text-slate-400 hover:text-slate-600 transition-colors"
              >
                <X size={16} />
              </button>
            </div>

            <div className="p-5 space-y-4 text-xs">
              
              {/* Punch count toggle */}
              <div>
                <label className="block font-bold text-slate-700 dark:text-slate-300 mb-1.5">
                  العدد الإلزامي للبصمات في كل وردية
                </label>
                <div className="grid grid-cols-2 gap-2">
                  <button
                    type="button"
                    onClick={() => setPunchPolicy(3)}
                    className={`py-2 px-3 rounded-lg border font-bold text-xs transition-colors ${
                      rules.REQUIRED_PUNCHES_PER_SHIFT === 3
                        ? 'bg-blue-50 dark:bg-blue-950/60 border-blue-500 text-blue-700 dark:text-blue-300'
                        : 'bg-slate-50 dark:bg-slate-800 border-slate-200 dark:border-slate-700 text-slate-600'
                    }`}
                  >
                    3 بصمات (أمر التخفيض)
                  </button>
                  <button
                    type="button"
                    onClick={() => setPunchPolicy(5)}
                    className={`py-2 px-3 rounded-lg border font-bold text-xs transition-colors ${
                      rules.REQUIRED_PUNCHES_PER_SHIFT === 5
                        ? 'bg-blue-50 dark:bg-blue-950/60 border-blue-500 text-blue-700 dark:text-blue-300'
                        : 'bg-slate-50 dark:bg-slate-800 border-slate-200 dark:border-slate-700 text-slate-600'
                    }`}
                  >
                    5 بصمات (الاعتيادي)
                  </button>
                </div>
              </div>

              {/* Time Windows */}
              <div>
                <label className="block font-bold text-slate-700 dark:text-slate-300 mb-2">
                  أوقات النوافذ الزمنية لفترات البصمة
                </label>
                <div className="space-y-2">
                  {Object.entries(rules.TIME_PERIODS).map(([periodName, times]) => (
                    <div
                      key={periodName}
                      className="flex items-center justify-between p-2 rounded bg-slate-50 dark:bg-slate-800/60 border border-slate-200 dark:border-slate-700"
                    >
                      <span className="font-semibold text-slate-800 dark:text-slate-200">{periodName}</span>
                      <div className="flex items-center gap-1.5 font-mono">
                        <input
                          type="time"
                          value={times.start}
                          onChange={(e) => {
                            setRules(prev => ({
                              ...prev,
                              TIME_PERIODS: {
                                ...prev.TIME_PERIODS,
                                [periodName]: { ...prev.TIME_PERIODS[periodName], start: e.target.value }
                              }
                            }));
                          }}
                          className="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-700 rounded px-1.5 py-0.5 text-xs text-slate-800 dark:text-slate-200"
                        />
                        <span className="text-slate-400">إلى</span>
                        <input
                          type="time"
                          value={times.end}
                          onChange={(e) => {
                            setRules(prev => ({
                              ...prev,
                              TIME_PERIODS: {
                                ...prev.TIME_PERIODS,
                                [periodName]: { ...prev.TIME_PERIODS[periodName], end: e.target.value }
                              }
                            }));
                          }}
                          className="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-700 rounded px-1.5 py-0.5 text-xs text-slate-800 dark:text-slate-200"
                        />
                      </div>
                    </div>
                  ))}
                </div>
              </div>

            </div>

            <div className="px-5 py-3 border-t border-slate-200 dark:border-slate-800 bg-slate-50 dark:bg-slate-850 flex justify-end gap-2">
              <button
                onClick={() => {
                  setRules(ATTENDANCE_RULES);
                }}
                className="px-3 py-1.5 rounded border border-slate-200 dark:border-slate-700 text-slate-600 dark:text-slate-400 hover:bg-slate-100 text-xs font-semibold"
              >
                استعادة الافتراضي
              </button>
              <button
                onClick={() => setIsSettingsOpen(false)}
                className="px-4 py-1.5 rounded bg-slate-900 dark:bg-slate-800 hover:bg-slate-800 text-white font-semibold text-xs shadow-xs"
              >
                إغلاق وتطبيق
              </button>
            </div>

          </div>
        </div>
      )}

      {/* =========================================================================
          6. PRINT-ONLY OFFICIAL AUDIT SLIP (Appears only during window.print())
      ========================================================================= */}
      <div className="print-only hidden p-8 text-black bg-white">
        {inspectEmployee ? (
          <div>
            <div className="border-b-2 border-black pb-4 mb-6 text-center">
              <h1 className="text-xl font-bold">محضر تدقيق البصمات والمطابقة الإدارية</h1>
              <p className="text-sm mt-1">الرقم المرجعي للأمر الإداري: {administrativeOrderNo}</p>
              <p className="text-xs text-gray-600 mt-0.5">تاريخ الطباعة: {new Date().toLocaleDateString('ar')}</p>
            </div>

            <div className="grid grid-cols-2 gap-4 mb-6 border p-4 text-sm">
              <div><strong>اسم الموظف:</strong> {inspectEmployee.name}</div>
              <div><strong>الرقم الوظيفي:</strong> #{inspectEmployee.id}</div>
              <div><strong>إجمالي أيام الحضور:</strong> {inspectEmployee.shifts.length} يوم</div>
              <div><strong>إجمالي حركات البصمة:</strong> {inspectEmployee.totalPunches} بصمة</div>
              <div><strong>نسبة الالتزام المقررة:</strong> {inspectEmployee.completionRate}%</div>
              <div><strong>القرار الإداري:</strong> {inspectEmployee.isEligible ? 'مستوفٍ للشروط ومستحق' : 'توجد نواقص غير معتمدة'}</div>
            </div>

            <table className="w-full text-right border-collapse text-xs border mb-8">
              <thead>
                <tr className="border-b bg-gray-100">
                  <th className="p-2 border">تاريخ النوبة</th>
                  <th className="p-2 border">حالة اليوم</th>
                  <th className="p-2 border">عدد البصمات</th>
                  <th className="p-2 border">تفاصيل البصمات المسجلة</th>
                </tr>
              </thead>
              <tbody>
                {inspectEmployee.shifts.map((s, idx) => (
                  <tr key={idx} className="border-b">
                    <td className="p-2 border font-mono">{s.shift_date}</td>
                    <td className="p-2 border">{s.status}</td>
                    <td className="p-2 border font-mono">{s.punches_count} من {rules.REQUIRED_PUNCHES_PER_SHIFT}</td>
                    <td className="p-2 border">
                      {s.punches_details.map(d => `${d.period}: ${d.time}`).join(' | ')}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>

            <div className="flex justify-between pt-12 text-sm">
              <div className="text-center">
                <p>توقيع مسؤول الرقابة والتفتيش</p>
                <p className="mt-8">......................................</p>
              </div>
              <div className="text-center">
                <p>مصادقة مدير الموارد البشرية</p>
                <p className="mt-8">......................................</p>
              </div>
            </div>
          </div>
        ) : null}
      </div>

    </div>
  );
};

export default App;
