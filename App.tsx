import React, { useState, useMemo, useRef } from 'react';
import * as XLSX from 'xlsx';
import { 
  FileUp, Search, Users, LayoutDashboard, Fingerprint, 
  RefreshCcw, X, ShieldCheck, CheckCircle2, 
  AlertTriangle, FileSpreadsheet, ShieldAlert,
  Clock, ListChecks, ChevronLeft, ChevronRight,
  Settings, SunMedium, MoonStar, Download,
  Check, ArrowUpDown, Eye, Filter, Sparkles
} from 'lucide-react';
import { processExcelData } from './utils/attendanceProcessor';
import { generateSampleDataset } from './utils/demoData';
import { generateBulkAuditSummaryAI, generateDetailedAuditSummaryAI } from './services/geminiService';
import { EmployeeSummary, Rules } from './types';
import { ATTENDANCE_RULES } from './constants';

type ViewType = 'dashboard' | 'audit';
type FilterType = 'all' | 'eligible' | 'deficient';
type SortType = 'id_asc' | 'id_desc' | 'completion_asc' | 'completion_desc' | 'name_asc';
type ViewMode = 'table' | 'grid';

const ITEMS_PER_PAGE = 20;

export const App: React.FC = () => {
  const [employees, setEmployees] = useState<EmployeeSummary[]>([]);
  const [rawData, setRawData] = useState<any[]>([]);
  const [fileName, setFileName] = useState<string | null>(null);
  const [searchTerm, setSearchTerm] = useState('');
  const [filterType, setFilterType] = useState<FilterType>('all');
  const [sortBy, setSortBy] = useState<SortType>('id_asc');
  const [viewMode, setViewMode] = useState<ViewMode>('table');
  const [currentPage, setCurrentPage] = useState(1);
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [currentView, setCurrentView] = useState<ViewType>('dashboard');
  const [bulkAiReport, setBulkAiReport] = useState<string | null>(null);
  const [rules, setRules] = useState<Rules>(ATTENDANCE_RULES);
  const [isSettingsOpen, setIsSettingsOpen] = useState(false);
  const [selectedEmployee, setSelectedEmployee] = useState<EmployeeSummary | null>(null);
  const [isDarkMode, setIsDarkMode] = useState<boolean>(() => {
    return window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches;
  });

  const fileInputRef = useRef<HTMLInputElement>(null);

  React.useEffect(() => {
    if (isDarkMode) {
      document.documentElement.classList.add('dark');
    } else {
      document.documentElement.classList.remove('dark');
    }
  }, [isDarkMode]);

  // High-level statistics
  const stats = useMemo(() => {
    if (employees.length === 0) return null;
    const total = employees.length;
    const eligible = employees.filter(e => e.isEligible).length;
    const deficient = total - eligible;
    const totalPunches = employees.reduce((sum, e) => sum + e.totalPunches, 0);
    const totalDays = employees.reduce((sum, e) => sum + e.shifts.length, 0);
    const compliancePercent = total > 0 ? Math.round((eligible / total) * 100) : 0;
    return { total, eligible, deficient, totalPunches, totalDays, compliancePercent };
  }, [employees]);

  // Filtering and sorting
  const filteredEmployees = useMemo(() => {
    let result = [...employees];
    if (filterType === 'eligible') result = result.filter(e => e.isEligible);
    else if (filterType === 'deficient') result = result.filter(e => !e.isEligible);

    const term = searchTerm.trim().toLowerCase();
    if (term) {
      result = result.filter(e => 
        e.name.toLowerCase().includes(term) || 
        String(e.id).includes(term)
      );
    }

    result.sort((a, b) => {
      switch (sortBy) {
        case 'completion_asc': return a.completionRate - b.completionRate;
        case 'completion_desc': return b.completionRate - a.completionRate;
        case 'name_asc': return a.name.localeCompare(b.name, 'ar');
        case 'id_desc': return (parseInt(String(b.id).replace(/\D/g, '')) || 0) - (parseInt(String(a.id).replace(/\D/g, '')) || 0);
        case 'id_asc':
        default: return (parseInt(String(a.id).replace(/\D/g, '')) || 0) - (parseInt(String(b.id).replace(/\D/g, '')) || 0);
      }
    });

    return result;
  }, [employees, searchTerm, filterType, sortBy]);

  const paginatedEmployees = useMemo(() => {
    const startIndex = (currentPage - 1) * ITEMS_PER_PAGE;
    return filteredEmployees.slice(startIndex, startIndex + ITEMS_PER_PAGE);
  }, [filteredEmployees, currentPage]);

  const totalPages = Math.ceil(filteredEmployees.length / ITEMS_PER_PAGE);

  React.useEffect(() => {
    setCurrentPage(1);
  }, [searchTerm, filterType, sortBy]);

  const handleFileUpload = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;
    setFileName(file.name);
    setIsAnalyzing(true);
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const json = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
        setRawData(json);
        const processed = processExcelData(json, rules);
        setEmployees(processed);
        setCurrentView('dashboard');
        setFilterType('all');
      } catch (err) {
        console.error("Excel read error:", err);
      } finally {
        setIsAnalyzing(false);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const handleLoadDemoData = () => {
    setIsAnalyzing(true);
    setFileName('سجلات_تجريبية_مؤسسية.xlsx');
    setTimeout(() => {
      const demoEmployees = generateSampleDataset(rules);
      setEmployees(demoEmployees);
      // Create artificial raw data representation for rule toggles
      setRawData([]);
      setCurrentView('dashboard');
      setFilterType('all');
      setIsAnalyzing(false);
    }, 200);
  };

  const handleQuickPunchesChange = (newCount: number) => {
    const newRules = { ...rules, REQUIRED_PUNCHES_PER_SHIFT: newCount };
    setRules(newRules);
    if (rawData.length > 0) {
      const processed = processExcelData(rawData, newRules);
      setEmployees(processed);
    } else if (employees.length > 0) {
      // Re-evaluate demo or existing data
      const processed = generateSampleDataset(newRules);
      setEmployees(processed);
    }
  };

  const handleSaveRules = (newRules: Rules) => {
    setRules(newRules);
    setIsSettingsOpen(false);
    if (rawData.length > 0) {
      const processed = processExcelData(rawData, newRules);
      setEmployees(processed);
    } else if (employees.length > 0) {
      const processed = generateSampleDataset(newRules);
      setEmployees(processed);
    }
  };

  const handleExportToExcel = () => {
    if (employees.length === 0) return;
    const summaryData = employees.map(emp => ({
      'الرقم الوظيفي': emp.id,
      'اسم الموظف': emp.name,
      'عدد أيام الحضور': emp.shifts.length,
      'إجمالي البصمات': emp.totalPunches,
      'أيام مكتملة': emp.shifts.filter(s => s.status === 'مكتملة').length,
      'أيام ناقصة': emp.shifts.filter(s => s.status === 'ناقصة').length,
      'نسبة الالتزام': `${emp.completionRate}%`,
      'الحالة الإدارية': emp.isEligible ? 'مستوفٍ للشروط' : 'به نواقص'
    }));

    const deficiencyLogs: any[] = [];
    employees.forEach(emp => {
      emp.shifts.filter(s => s.status === 'ناقصة').forEach(shift => {
        const found = shift.punches_details.filter(p => !p.is_duplicate).map(p => p.period);
        const allPossible = ["دخول صباحي", "بصمة وسطية", "بصمة مسائية", "بصمة ليلية", "بصمة خروج"];
        const missing = allPossible.filter(p => !found.includes(p));
        deficiencyLogs.push({
          'التاريخ': shift.shift_date,
          'الرقم الوظيفي': emp.id,
          'اسم الموظف': emp.name,
          'البصمات الناقصة': missing.join(' - ')
        });
      });
    });

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(summaryData), "الملخص العام");
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(deficiencyLogs), "سجل النواقص");
    XLSX.writeFile(workbook, `Basmati_Pro_Audit_${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  const runBulkAuditAnalysis = async () => {
    if (employees.length === 0) return;
    setCurrentView('audit');
    setIsAnalyzing(true);
    try {
      const res = await generateBulkAuditSummaryAI(employees);
      setBulkAiReport(res.text);
    } catch (e) {
      setBulkAiReport("تعذر إكمال عملية التدقيق الشامل في الوقت الحالي.");
    } finally {
      setIsAnalyzing(false);
    }
  };

  return (
    <div className="flex flex-col h-screen w-screen bg-slate-50 dark:bg-slate-950 text-slate-900 dark:text-slate-100 font-sans text-xs select-none">
      
      {/* 1. Senior Executive Top Navigation Bar */}
      <header className="h-14 px-4 sm:px-6 border-b border-slate-200 dark:border-slate-800 bg-white dark:bg-slate-900 flex items-center justify-between gap-4 shrink-0 z-40">
        
        {/* Brand & System Wordmark */}
        <div className="flex items-center gap-3">
          <div className="w-8 h-8 rounded-lg bg-slate-900 dark:bg-slate-800 text-white flex items-center justify-center font-bold shadow-xs">
            <ShieldCheck size={18} className="text-blue-400" />
          </div>
          <div>
            <div className="flex items-center gap-1.5">
              <span className="font-bold text-sm tracking-tight text-slate-900 dark:text-white">بصمتي</span>
              <span className="text-[10px] font-semibold tracking-wider px-1.5 py-0.2 bg-blue-50 dark:bg-blue-950/60 text-blue-700 dark:text-blue-400 border border-blue-200/60 dark:border-blue-800/60 rounded">
                برو
              </span>
            </div>
            <p className="text-[10px] text-slate-500 dark:text-slate-400 hidden sm:block">منظومة تدقيق الحضور ومطابقة الورديات</p>
          </div>
        </div>

        {/* View Tabs */}
        {employees.length > 0 && (
          <nav className="flex items-center gap-1 bg-slate-100 dark:bg-slate-800/80 p-0.5 rounded-lg border border-slate-200/80 dark:border-slate-700/80">
            <button
              onClick={() => setCurrentView('dashboard')}
              className={`px-3 py-1.5 rounded-md font-semibold text-xs flex items-center gap-1.5 transition-colors ${
                currentView === 'dashboard' 
                  ? 'bg-white dark:bg-slate-700 text-slate-900 dark:text-white shadow-xs' 
                  : 'text-slate-600 dark:text-slate-400 hover:text-slate-900 dark:hover:text-white'
              }`}
            >
              <LayoutDashboard size={14} />
              <span>لوحة السجلات</span>
            </button>
            <button
              onClick={runBulkAuditAnalysis}
              className={`px-3 py-1.5 rounded-md font-semibold text-xs flex items-center gap-1.5 transition-colors ${
                currentView === 'audit' 
                  ? 'bg-white dark:bg-slate-700 text-slate-900 dark:text-white shadow-xs' 
                  : 'text-slate-600 dark:text-slate-400 hover:text-slate-900 dark:hover:text-white'
              }`}
            >
              <ListChecks size={14} />
              <span>تقرير النواقص</span>
            </button>
          </nav>
        )}

        {/* Actions & Utilities */}
        <div className="flex items-center gap-2">
          
          {/* Punch Rule Selector Segment */}
          <div className="flex items-center bg-slate-100 dark:bg-slate-800/80 p-0.5 rounded-lg border border-slate-200/80 dark:border-slate-700/80 text-[11px]">
            <span className="px-2 text-slate-500 dark:text-slate-400 font-medium hidden md:inline">
              الاشتراط:
            </span>
            <button
              onClick={() => handleQuickPunchesChange(3)}
              className={`px-2 py-1 rounded-md font-semibold transition-colors ${
                rules.REQUIRED_PUNCHES_PER_SHIFT === 3
                  ? 'bg-white dark:bg-slate-700 text-slate-900 dark:text-white shadow-xs font-bold'
                  : 'text-slate-600 dark:text-slate-400 hover:text-slate-900 dark:hover:text-white'
              }`}
              title="تطبيق قاعدة 3 بصمات لكل وردية"
            >
              3 بصمات
            </button>
            <button
              onClick={() => handleQuickPunchesChange(5)}
              className={`px-2 py-1 rounded-md font-semibold transition-colors ${
                rules.REQUIRED_PUNCHES_PER_SHIFT === 5
                  ? 'bg-white dark:bg-slate-700 text-slate-900 dark:text-white shadow-xs font-bold'
                  : 'text-slate-600 dark:text-slate-400 hover:text-slate-900 dark:hover:text-white'
              }`}
              title="تطبيق قاعدة 5 بصمات لكل وردية"
            >
              5 بصمات
            </button>
          </div>

          {/* Settings Trigger */}
          <button
            onClick={() => setIsSettingsOpen(true)}
            className="p-1.5 rounded-lg border border-slate-200 dark:border-slate-800 hover:bg-slate-100 dark:hover:bg-slate-800 text-slate-600 dark:text-slate-300 transition-colors"
            title="إعدادات مواعيد الفترات"
          >
            <Settings size={15} />
          </button>

          {/* Dark / Light Toggle */}
          <button
            onClick={() => setIsDarkMode(!isDarkMode)}
            className="p-1.5 rounded-lg border border-slate-200 dark:border-slate-800 hover:bg-slate-100 dark:hover:bg-slate-800 text-slate-600 dark:text-slate-300 transition-colors"
            title={isDarkMode ? 'التبديل إلى الوضع الفاتح' : 'التبديل إلى الوضع الداكن'}
          >
            {isDarkMode ? <SunMedium size={15} className="text-amber-400" /> : <MoonStar size={15} className="text-slate-600" />}
          </button>

          {/* Export Excel (if employees loaded) */}
          {employees.length > 0 && (
            <button
              onClick={handleExportToExcel}
              className="hidden sm:flex items-center gap-1.5 px-3 py-1.5 rounded-lg border border-emerald-300 dark:border-emerald-800/80 bg-emerald-50/70 hover:bg-emerald-100/70 dark:bg-emerald-950/40 dark:hover:bg-emerald-950/70 text-emerald-800 dark:text-emerald-300 font-semibold text-xs transition-colors"
            >
              <Download size={14} />
              <span>تصدير Excel</span>
            </button>
          )}

          {/* Primary Upload CTA */}
          <button
            onClick={() => fileInputRef.current?.click()}
            className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-slate-900 dark:bg-blue-600 hover:bg-slate-800 dark:hover:bg-blue-500 text-white font-semibold text-xs shadow-xs transition-colors cursor-pointer"
          >
            <FileUp size={14} />
            <span>{employees.length === 0 ? 'استيراد ملف' : 'تغيير الملف'}</span>
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

      {/* Main Workspace Body */}
      <main className="flex-1 overflow-y-auto custom-scroll p-4 sm:p-6 flex flex-col min-w-0">
        
        {/* CASE A: No Data Loaded -> Executive Data Intake Center */}
        {employees.length === 0 ? (
          <div className="max-w-4xl w-full mx-auto my-auto py-8">
            <div className="border border-slate-200 dark:border-slate-800 rounded-xl bg-white dark:bg-slate-900 shadow-xs overflow-hidden">
              
              {/* Header Banner */}
              <div className="p-6 sm:p-8 border-b border-slate-200 dark:border-slate-800 bg-slate-50/50 dark:bg-slate-900/50">
                <div className="max-w-2xl">
                  <div className="inline-flex items-center gap-2 px-2.5 py-1 rounded bg-blue-50 dark:bg-blue-950/50 border border-blue-200/60 dark:border-blue-900/50 text-blue-700 dark:text-blue-300 font-semibold text-[11px] mb-3">
                    <ShieldCheck size={13} />
                    <span>منظومة التدقيق والمطابقة المؤسسية المعتمدة</span>
                  </div>
                  <h1 className="text-xl sm:text-2xl font-bold tracking-tight text-slate-900 dark:text-white">
                    تدقيق سجلات الحضور والبصمات
                  </h1>
                  <p className="text-slate-600 dark:text-slate-400 text-xs sm:text-sm mt-1.5 leading-relaxed">
                    منصة تدقيق احترافية مصممة لمراجعة نوبات الكادر، كشف النواقص في فترات الدخول والخروج والوسطية، واحتساب نسب الالتزام بدقة متناهية.
                  </p>
                </div>
              </div>

              {/* Ingestion Area */}
              <div className="p-6 sm:p-8 space-y-6">
                
                {/* Drag and Drop Box */}
                <div 
                  onClick={() => fileInputRef.current?.click()}
                  className="border-2 border-dashed border-slate-200 dark:border-slate-700 hover:border-blue-500 dark:hover:border-blue-400 bg-slate-50/60 dark:bg-slate-800/30 hover:bg-blue-50/30 dark:hover:bg-blue-950/20 rounded-xl p-8 text-center cursor-pointer transition-all flex flex-col items-center justify-center gap-3"
                >
                  <div className="w-12 h-12 rounded-xl bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 flex items-center justify-center text-slate-700 dark:text-slate-300 shadow-xs">
                    {isAnalyzing ? <RefreshCcw size={22} className="animate-spin text-blue-600" /> : <FileSpreadsheet size={22} className="text-blue-600" />}
                  </div>
                  <div>
                    <p className="font-bold text-sm text-slate-900 dark:text-white">
                      انقر لاختيار ملف Excel أو اسحب الملف وأفلته هنا
                    </p>
                    <p className="text-slate-500 dark:text-slate-400 text-xs mt-1">
                      يدعم ملفات أجهزة البصمة بصيغ (.xlsx, .xls) لكافة الموديلات والأنظمة
                    </p>
                  </div>
                  <div className="flex items-center gap-2 mt-2">
                    <span className="px-3 py-1.5 bg-slate-900 dark:bg-blue-600 text-white rounded-lg text-xs font-semibold shadow-xs">
                      تصفح الملفات
                    </span>
                  </div>
                </div>

                {/* Instant Demo Data Option */}
                <div className="flex flex-col sm:flex-row items-center justify-between p-4 rounded-xl border border-slate-200 dark:border-slate-800 bg-slate-50/50 dark:bg-slate-800/40 gap-4">
                  <div className="flex items-center gap-3">
                    <div className="p-2 rounded-lg bg-emerald-50 dark:bg-emerald-950/40 text-emerald-700 dark:text-emerald-300 border border-emerald-200/60 dark:border-emerald-800/60 shrink-0">
                      <Sparkles size={16} />
                    </div>
                    <div>
                      <p className="font-bold text-xs text-slate-800 dark:text-slate-200">
                        هل ترغب في تجربة النظام مباشرة الآن؟
                      </p>
                      <p className="text-[11px] text-slate-500 dark:text-slate-400 mt-0.5">
                        يمكنك تحميل عينة سجلات مطابقة واقعية (6 موظفين مع حالات منضبطة ونواقص) بنقرة واحدة.
                      </p>
                    </div>
                  </div>
                  <button
                    onClick={handleLoadDemoData}
                    disabled={isAnalyzing}
                    className="w-full sm:w-auto px-4 py-2 bg-white dark:bg-slate-800 hover:bg-slate-100 dark:hover:bg-slate-700 text-slate-800 dark:text-slate-200 font-semibold text-xs rounded-lg border border-slate-300 dark:border-slate-700 shadow-xs transition-colors shrink-0"
                  >
                    تحميل عينة تجريبية
                  </button>
                </div>

                {/* Data Schema Specs */}
                <div className="pt-2">
                  <p className="font-bold text-[11px] uppercase tracking-wider text-slate-400 mb-2.5">
                    الأعمدة المعيارية المدعومة تلقائياً في الملف:
                  </p>
                  <div className="grid grid-cols-2 sm:grid-cols-4 gap-2 text-center text-xs">
                    <div className="p-2.5 rounded-lg border border-slate-200/80 dark:border-slate-800 bg-white dark:bg-slate-900">
                      <span className="font-mono font-bold text-blue-600 dark:text-blue-400 block text-[11px]">ACNo / رقم</span>
                      <span className="text-[10px] text-slate-500 mt-0.5 block">الرقم الوظيفي الفريد</span>
                    </div>
                    <div className="p-2.5 rounded-lg border border-slate-200/80 dark:border-slate-800 bg-white dark:bg-slate-900">
                      <span className="font-mono font-bold text-blue-600 dark:text-blue-400 block text-[11px]">Name / الاسم</span>
                      <span className="text-[10px] text-slate-500 mt-0.5 block">الاسم الثلاثي أو الوظيفي</span>
                    </div>
                    <div className="p-2.5 rounded-lg border border-slate-200/80 dark:border-slate-800 bg-white dark:bg-slate-900">
                      <span className="font-mono font-bold text-blue-600 dark:text-blue-400 block text-[11px]">Date / التاريخ</span>
                      <span className="text-[10px] text-slate-500 mt-0.5 block">تاريخ يوم البصمة</span>
                    </div>
                    <div className="p-2.5 rounded-lg border border-slate-200/80 dark:border-slate-800 bg-white dark:bg-slate-900">
                      <span className="font-mono font-bold text-blue-600 dark:text-blue-400 block text-[11px]">Time / الوقت</span>
                      <span className="text-[10px] text-slate-500 mt-0.5 block">حركات البصمة المسجلة</span>
                    </div>
                  </div>
                </div>

              </div>
            </div>
          </div>
        ) : (
          /* CASE B: Dataset Loaded -> Executive Workspace */
          <div className="space-y-4 max-w-7xl w-full mx-auto">
            
            {/* View Switch: Dashboard vs Audit */}
            {currentView === 'dashboard' ? (
              <>
                {/* 1. KPI Metric Strip */}
                {stats && (
                  <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-5 gap-3">
                    
                    {/* Metric 1: Total Staff */}
                    <div 
                      onClick={() => setFilterType('all')}
                      className={`p-3.5 rounded-xl bg-white dark:bg-slate-900 border transition-all cursor-pointer ${
                        filterType === 'all'
                          ? 'border-blue-500/80 ring-1 ring-blue-500/30'
                          : 'border-slate-200 dark:border-slate-800 hover:border-slate-300 dark:hover:border-slate-700'
                      }`}
                    >
                      <div className="flex items-center justify-between text-slate-500 dark:text-slate-400 mb-1">
                        <span className="text-[11px] font-medium">إجمالي الكادر</span>
                        <Users size={15} />
                      </div>
                      <div className="flex items-baseline gap-1.5">
                        <span className="text-2xl font-bold tracking-tight text-slate-900 dark:text-white tabular-nums">
                          {stats.total}
                        </span>
                        <span className="text-[11px] text-slate-500">موظف</span>
                      </div>
                    </div>

                    {/* Metric 2: Eligible / Compliant */}
                    <div 
                      onClick={() => setFilterType('eligible')}
                      className={`p-3.5 rounded-xl bg-white dark:bg-slate-900 border transition-all cursor-pointer ${
                        filterType === 'eligible'
                          ? 'border-emerald-500/80 ring-1 ring-emerald-500/30'
                          : 'border-slate-200 dark:border-slate-800 hover:border-slate-300 dark:hover:border-slate-700'
                      }`}
                    >
                      <div className="flex items-center justify-between text-emerald-600 dark:text-emerald-400 mb-1">
                        <span className="text-[11px] font-medium">سجلات مكتملة</span>
                        <CheckCircle2 size={15} />
                      </div>
                      <div className="flex items-baseline gap-1.5">
                        <span className="text-2xl font-bold tracking-tight text-emerald-600 dark:text-emerald-400 tabular-nums">
                          {stats.eligible}
                        </span>
                        <span className="text-[11px] text-slate-500">
                          ({stats.total > 0 ? Math.round((stats.eligible / stats.total) * 100) : 0}%)
                        </span>
                      </div>
                    </div>

                    {/* Metric 3: Deficient */}
                    <div 
                      onClick={() => setFilterType('deficient')}
                      className={`p-3.5 rounded-xl bg-white dark:bg-slate-900 border transition-all cursor-pointer ${
                        filterType === 'deficient'
                          ? 'border-rose-500/80 ring-1 ring-rose-500/30'
                          : 'border-slate-200 dark:border-slate-800 hover:border-slate-300 dark:hover:border-slate-700'
                      }`}
                    >
                      <div className="flex items-center justify-between text-rose-600 dark:text-rose-400 mb-1">
                        <span className="text-[11px] font-medium">سجلات بها نواقص</span>
                        <ShieldAlert size={15} />
                      </div>
                      <div className="flex items-baseline gap-1.5">
                        <span className="text-2xl font-bold tracking-tight text-rose-600 dark:text-rose-400 tabular-nums">
                          {stats.deficient}
                        </span>
                        <span className="text-[11px] text-slate-500">
                          ({stats.total > 0 ? Math.round((stats.deficient / stats.total) * 100) : 0}%)
                        </span>
                      </div>
                    </div>

                    {/* Metric 4: Total Shift Days */}
                    <div className="p-3.5 rounded-xl bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800">
                      <div className="flex items-center justify-between text-slate-500 dark:text-slate-400 mb-1">
                        <span className="text-[11px] font-medium">أيام العمل المسجلة</span>
                        <Clock size={15} />
                      </div>
                      <div className="flex items-baseline gap-1.5">
                        <span className="text-2xl font-bold tracking-tight text-slate-900 dark:text-white tabular-nums">
                          {stats.totalDays}
                        </span>
                        <span className="text-[11px] text-slate-500">يوم نوبة</span>
                      </div>
                    </div>

                    {/* Metric 5: Total Punch Actions */}
                    <div className="p-3.5 rounded-xl bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 col-span-2 sm:col-span-1">
                      <div className="flex items-center justify-between text-slate-500 dark:text-slate-400 mb-1">
                        <span className="text-[11px] font-medium">إجمالي حركات البصمة</span>
                        <Fingerprint size={15} />
                      </div>
                      <div className="flex items-baseline gap-1.5">
                        <span className="text-2xl font-bold tracking-tight text-slate-900 dark:text-white tabular-nums">
                          {stats.totalPunches}
                        </span>
                        <span className="text-[11px] text-slate-500">حركة</span>
                      </div>
                    </div>

                  </div>
                )}

                {/* 2. Active Rule Context Notice */}
                <div className="flex flex-wrap items-center justify-between px-4 py-2.5 rounded-lg border border-slate-200 dark:border-slate-800 bg-white dark:bg-slate-900 text-xs">
                  <div className="flex items-center gap-2">
                    <span className="w-2 h-2 rounded-full bg-blue-600 dark:bg-blue-400" />
                    <span className="font-medium text-slate-600 dark:text-slate-300">
                      القاعدة المعمول بها حالياً:
                    </span>
                    <strong className="text-slate-900 dark:text-white font-bold">
                      {rules.REQUIRED_PUNCHES_PER_SHIFT} بصمات / وردية
                    </strong>
                    <span className="text-slate-400">·</span>
                    <span className="text-slate-500 dark:text-slate-400 text-[11px]">
                      {rules.REQUIRED_PUNCHES_PER_SHIFT === 3 ? 'نظام الـ 3 بصمات (أمر التخفيض)' : 'نظام الـ 5 بصمات الاعتيادي'}
                    </span>
                  </div>
                  
                  {fileName && (
                    <span className="text-[11px] font-mono text-slate-500 dark:text-slate-400">
                      الملف النشط: {fileName}
                    </span>
                  )}
                </div>

                {/* 3. Toolbar & Filters */}
                <div className="flex flex-col md:flex-row items-stretch md:items-center justify-between gap-3 pt-1">
                  
                  {/* Search input */}
                  <div className="relative flex-1 max-w-md">
                    <Search className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400" size={14} />
                    <input
                      type="text"
                      placeholder="بحث باسم الموظف أو الرقم الوظيفي..."
                      value={searchTerm}
                      onChange={(e) => setSearchTerm(e.target.value)}
                      className="w-full bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-lg pr-9 pl-8 py-2 text-xs focus:outline-none focus:border-blue-500 dark:focus:border-blue-400 transition-colors placeholder:text-slate-400 text-slate-900 dark:text-white"
                    />
                    {searchTerm && (
                      <button
                        onClick={() => setSearchTerm('')}
                        className="absolute left-2.5 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600"
                      >
                        <X size={13} />
                      </button>
                    )}
                  </div>

                  {/* Filter & View Mode Controls */}
                  <div className="flex items-center gap-2 overflow-x-auto pb-1 md:pb-0">
                    
                    {/* Status Segmented Buttons */}
                    <div className="flex items-center bg-slate-100 dark:bg-slate-800 p-0.5 rounded-lg border border-slate-200 dark:border-slate-700/60 text-xs">
                      <button
                        onClick={() => setFilterType('all')}
                        className={`px-3 py-1.5 rounded-md font-semibold transition-colors ${
                          filterType === 'all'
                            ? 'bg-white dark:bg-slate-700 text-slate-900 dark:text-white shadow-xs'
                            : 'text-slate-600 dark:text-slate-400 hover:text-slate-900'
                        }`}
                      >
                        الكل ({employees.length})
                      </button>
                      <button
                        onClick={() => setFilterType('eligible')}
                        className={`px-3 py-1.5 rounded-md font-semibold transition-colors ${
                          filterType === 'eligible'
                            ? 'bg-white dark:bg-slate-700 text-emerald-700 dark:text-emerald-300 shadow-xs'
                            : 'text-slate-600 dark:text-slate-400 hover:text-slate-900'
                        }`}
                      >
                        منضبط ({employees.filter(e => e.isEligible).length})
                      </button>
                      <button
                        onClick={() => setFilterType('deficient')}
                        className={`px-3 py-1.5 rounded-md font-semibold transition-colors ${
                          filterType === 'deficient'
                            ? 'bg-white dark:bg-slate-700 text-rose-700 dark:text-rose-300 shadow-xs'
                            : 'text-slate-600 dark:text-slate-400 hover:text-slate-900'
                        }`}
                      >
                        نواقص ({employees.filter(e => !e.isEligible).length})
                      </button>
                    </div>

                    {/* Sort Selector */}
                    <div className="relative">
                      <select
                        value={sortBy}
                        onChange={(e) => setSortBy(e.target.value as SortType)}
                        className="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-lg pr-3 pl-8 py-1.5 text-xs font-medium text-slate-700 dark:text-slate-300 focus:outline-none focus:border-blue-500 cursor-pointer"
                      >
                        <option value="id_asc">الرقم الوظيفي (تصاعدي)</option>
                        <option value="id_desc">الرقم الوظيفي (تنازلي)</option>
                        <option value="completion_desc">الأعلى التزاماً</option>
                        <option value="completion_asc">الأقل التزاماً</option>
                        <option value="name_asc">الاسم (أبجدياً)</option>
                      </select>
                    </div>

                    {/* View Mode Toggle */}
                    <div className="flex items-center bg-slate-100 dark:bg-slate-800 p-0.5 rounded-lg border border-slate-200 dark:border-slate-700/60">
                      <button
                        onClick={() => setViewMode('table')}
                        className={`p-1.5 rounded-md transition-colors ${
                          viewMode === 'table'
                            ? 'bg-white dark:bg-slate-700 text-slate-900 dark:text-white shadow-xs'
                            : 'text-slate-500 hover:text-slate-900 dark:text-slate-400'
                        }`}
                        title="عرض الجدول الإداري التفصيلي"
                      >
                        <FileSpreadsheet size={15} />
                      </button>
                      <button
                        onClick={() => setViewMode('grid')}
                        className={`p-1.5 rounded-md transition-colors ${
                          viewMode === 'grid'
                            ? 'bg-white dark:bg-slate-700 text-slate-900 dark:text-white shadow-xs'
                            : 'text-slate-500 hover:text-slate-900 dark:text-slate-400'
                        }`}
                        title="عرض بطاقات الموظفين"
                      >
                        <LayoutDashboard size={15} />
                      </button>
                    </div>

                  </div>
                </div>

                {/* 4. Main Data Content: Table or Grid */}
                {viewMode === 'table' ? (
                  <div className="border border-slate-200 dark:border-slate-800 rounded-xl bg-white dark:bg-slate-900 overflow-hidden shadow-xs">
                    <div className="overflow-x-auto">
                      <table className="w-full text-right border-collapse text-xs">
                        <thead>
                          <tr className="bg-slate-50 dark:bg-slate-800/60 border-b border-slate-200 dark:border-slate-800 text-slate-500 dark:text-slate-400 font-semibold text-[11px]">
                            <th className="py-2.5 px-4 w-24">الرقم الوظيفي</th>
                            <th className="py-2.5 px-4">اسم الموظف</th>
                            <th className="py-2.5 px-4 text-center w-24">أيام العمل</th>
                            <th className="py-2.5 px-4 text-center w-24">إجمالي البصمات</th>
                            <th className="py-2.5 px-4 text-center w-24">أيام النواقص</th>
                            <th className="py-2.5 px-4 text-center w-36">نسبة الالتزام</th>
                            <th className="py-2.5 px-4 text-center w-28">الحالة</th>
                            <th className="py-2.5 px-4 text-left w-20">إجراء</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-100 dark:divide-slate-800/60 font-normal">
                          {paginatedEmployees.length === 0 ? (
                            <tr>
                              <td colSpan={8} className="py-12 text-center text-slate-400">
                                لا توجد سجلات تطابق معايير البحث والفلترة.
                              </td>
                            </tr>
                          ) : (
                            paginatedEmployees.map((emp) => {
                              const incompleteDays = emp.shifts.filter(s => s.status === 'ناقصة').length;
                              return (
                                <tr 
                                  key={emp.id}
                                  className="hover:bg-slate-50/80 dark:hover:bg-slate-800/40 transition-colors"
                                >
                                  <td className="py-2.5 px-4 font-mono font-semibold text-slate-600 dark:text-slate-400 tabular-nums">
                                    #{emp.id}
                                  </td>
                                  <td className="py-2.5 px-4">
                                    <div className="flex items-center gap-2.5">
                                      <div className="w-6 h-6 rounded bg-slate-100 dark:bg-slate-800 text-slate-700 dark:text-slate-300 flex items-center justify-center font-bold text-[11px] shrink-0">
                                        {emp.name.charAt(0)}
                                      </div>
                                      <span className="font-semibold text-slate-800 dark:text-slate-200">
                                        {emp.name}
                                      </span>
                                    </div>
                                  </td>
                                  <td className="py-2.5 px-4 text-center tabular-nums text-slate-600 dark:text-slate-400">
                                    {emp.shifts.length}
                                  </td>
                                  <td className="py-2.5 px-4 text-center font-semibold text-slate-700 dark:text-slate-300 tabular-nums">
                                    {emp.totalPunches}
                                  </td>
                                  <td className="py-2.5 px-4 text-center tabular-nums">
                                    {incompleteDays > 0 ? (
                                      <span className="text-rose-600 dark:text-rose-400 font-semibold">
                                        {incompleteDays} أيام
                                      </span>
                                    ) : (
                                      <span className="text-slate-400">—</span>
                                    )}
                                  </td>
                                  <td className="py-2.5 px-4 text-center">
                                    <div className="flex items-center justify-center gap-2">
                                      <div className="w-16 h-1.5 bg-slate-100 dark:bg-slate-800 rounded-full overflow-hidden">
                                        <div 
                                          className={`h-full rounded-full ${emp.isEligible ? 'bg-emerald-500' : 'bg-amber-500'}`}
                                          style={{ width: `${Math.min(100, emp.completionRate)}%` }}
                                        />
                                      </div>
                                      <span className="font-mono font-semibold text-slate-700 dark:text-slate-300 tabular-nums text-[11px]">
                                        {emp.completionRate}%
                                      </span>
                                    </div>
                                  </td>
                                  <td className="py-2.5 px-4 text-center">
                                    <span className={`inline-block px-2 py-0.5 rounded text-[10px] font-semibold border ${
                                      emp.isEligible 
                                        ? 'bg-emerald-50 text-emerald-700 border-emerald-200/80 dark:bg-emerald-950/40 dark:text-emerald-300 dark:border-emerald-800/40' 
                                        : 'bg-rose-50 text-rose-700 border-rose-200/80 dark:bg-rose-950/40 dark:text-rose-300 dark:border-rose-800/40'
                                    }`}>
                                      {emp.isEligible ? 'مستوفٍ' : 'به نواقص'}
                                    </span>
                                  </td>
                                  <td className="py-2.5 px-4 text-left">
                                    <button
                                      onClick={() => setSelectedEmployee(emp)}
                                      className="px-2 py-1 text-xs rounded border border-slate-200 dark:border-slate-700 hover:bg-slate-100 dark:hover:bg-slate-800 text-slate-700 dark:text-slate-300 font-medium transition-colors"
                                    >
                                      معاينة
                                    </button>
                                  </td>
                                </tr>
                              );
                            })
                          )}
                        </tbody>
                      </table>
                    </div>
                  </div>
                ) : (
                  /* Grid Card View */
                  <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4 gap-3">
                    {paginatedEmployees.map((emp) => {
                      const incompleteDays = emp.shifts.filter(s => s.status === 'ناقصة').length;
                      return (
                        <div
                          key={emp.id}
                          onClick={() => setSelectedEmployee(emp)}
                          className="p-3.5 rounded-xl border border-slate-200 dark:border-slate-800 bg-white dark:bg-slate-900 hover:border-blue-400 dark:hover:border-blue-500 cursor-pointer transition-all flex flex-col justify-between gap-3 shadow-xs"
                        >
                          <div className="flex items-start justify-between">
                            <div className="flex items-center gap-2">
                              <div className="w-8 h-8 rounded-lg bg-slate-100 dark:bg-slate-800 text-slate-700 dark:text-slate-300 flex items-center justify-center font-bold text-xs shrink-0">
                                {emp.name.charAt(0)}
                              </div>
                              <div className="min-w-0">
                                <h3 className="font-semibold text-slate-900 dark:text-white truncate text-xs">
                                  {emp.name}
                                </h3>
                                <p className="font-mono text-slate-400 text-[10px] tabular-nums">
                                  #{emp.id}
                                </p>
                              </div>
                            </div>
                            <span className={`px-1.5 py-0.5 rounded text-[10px] font-semibold border ${
                              emp.isEligible 
                                ? 'bg-emerald-50 text-emerald-700 border-emerald-200/80 dark:bg-emerald-950/40 dark:text-emerald-300 dark:border-emerald-800/40' 
                                : 'bg-rose-50 text-rose-700 border-rose-200/80 dark:bg-rose-950/40 dark:text-rose-300 dark:border-rose-800/40'
                            }`}>
                              {emp.isEligible ? 'مستوفٍ' : 'به نواقص'}
                            </span>
                          </div>

                          <div className="flex items-center justify-between text-[11px] pt-2 border-t border-slate-100 dark:border-slate-800/80 text-slate-500 dark:text-slate-400">
                            <span>{emp.shifts.length} أيام نوبة</span>
                            <span className="font-mono tabular-nums font-semibold text-slate-700 dark:text-slate-300">
                              {emp.totalPunches} بصمة
                            </span>
                          </div>

                          <div className="flex items-center justify-between text-[11px]">
                            <span className="text-slate-500">الالتزام: <strong className="text-slate-800 dark:text-slate-200 font-mono">{emp.completionRate}%</strong></span>
                            {incompleteDays > 0 && (
                              <span className="text-rose-600 dark:text-rose-400 font-medium text-[10px]">
                                {incompleteDays} يوم ناقص
                              </span>
                            )}
                          </div>
                        </div>
                      );
                    })}
                  </div>
                )}

                {/* Pagination Controls */}
                {totalPages > 1 && (
                  <div className="flex items-center justify-between pt-3 border-t border-slate-200 dark:border-slate-800 text-xs text-slate-500 dark:text-slate-400">
                    <div>
                      عرض الصفحة <strong className="text-slate-800 dark:text-white font-mono">{currentPage}</strong> من أصل <strong className="text-slate-800 dark:text-white font-mono">{totalPages}</strong> ({filteredEmployees.length} موظف إجمالاً)
                    </div>
                    <div className="flex items-center gap-1">
                      <button
                        onClick={() => setCurrentPage(p => Math.max(1, p - 1))}
                        disabled={currentPage === 1}
                        className="px-2.5 py-1 rounded border border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-800 disabled:opacity-40 hover:bg-slate-50 transition-colors"
                      >
                        السابق
                      </button>
                      <button
                        onClick={() => setCurrentPage(p => Math.min(totalPages, p + 1))}
                        disabled={currentPage === totalPages}
                        className="px-2.5 py-1 rounded border border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-800 disabled:opacity-40 hover:bg-slate-50 transition-colors"
                      >
                        التالي
                      </button>
                    </div>
                  </div>
                )}
              </>
            ) : (
              /* VIEW 2: Bulk Audit & Deficiency Report View */
              <div className="space-y-4">
                
                <div className="flex items-center justify-between">
                  <div>
                    <h2 className="text-base font-bold text-slate-900 dark:text-white">
                      تقرير النواقص ومطابقة الحضور
                    </h2>
                    <p className="text-slate-500 dark:text-slate-400 text-xs mt-0.5">
                      كشف تفصيلي بالحالات غير المستوفية لكامل البصمات المعتمدة ({rules.REQUIRED_PUNCHES_PER_SHIFT} بصمات)
                    </p>
                  </div>
                  <button
                    onClick={() => setCurrentView('dashboard')}
                    className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg border border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-800 hover:bg-slate-50 transition-colors text-xs font-semibold"
                  >
                    <ChevronLeft size={14} />
                    <span>العودة للوحة السجلات</span>
                  </button>
                </div>

                {/* Audit Analysis Card */}
                <div className="border border-slate-200 dark:border-slate-800 rounded-xl bg-white dark:bg-slate-900 p-5 shadow-xs">
                  <div className="flex items-center justify-between mb-3 pb-2.5 border-b border-slate-100 dark:border-slate-800">
                    <div className="flex items-center gap-2">
                      <ShieldCheck size={16} className="text-blue-600 dark:text-blue-400" />
                      <h3 className="font-bold text-xs text-slate-800 dark:text-white">
                        خلاصة التدقيق الإداري
                      </h3>
                    </div>
                    <span className="text-[11px] font-mono text-slate-400">
                      القاعدة المقررة: {rules.REQUIRED_PUNCHES_PER_SHIFT} بصمات
                    </span>
                  </div>
                  <div className="bg-slate-50 dark:bg-slate-800/50 p-4 rounded-lg font-mono text-xs leading-relaxed text-slate-700 dark:text-slate-300 whitespace-pre-line border border-slate-200/60 dark:border-slate-700/60">
                    {isAnalyzing ? (
                      <div className="flex items-center gap-2 text-blue-600 dark:text-blue-400">
                        <RefreshCcw size={14} className="animate-spin" />
                        <span>جاري تدقيق البيانات ومطابقة الفترات الزمنية...</span>
                      </div>
                    ) : (
                      bulkAiReport || "جميع سجلات الموظفين مكتملة ومطابقة للقواعد المقررة بدون أي نواقص."
                    )}
                  </div>
                </div>

                {/* List of Deficient Employees */}
                <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                  {employees.filter(e => !e.isEligible).map(emp => (
                    <div
                      key={emp.id}
                      className="border border-slate-200 dark:border-slate-800 rounded-xl bg-white dark:bg-slate-900 p-4 shadow-xs"
                    >
                      <div className="flex items-center justify-between mb-2.5 pb-2 border-b border-slate-100 dark:border-slate-800">
                        <div className="flex items-center gap-2">
                          <div className="w-7 h-7 rounded bg-rose-50 dark:bg-rose-950/40 text-rose-700 dark:text-rose-400 flex items-center justify-center font-bold text-xs border border-rose-200/60 dark:border-rose-800/40">
                            {emp.name.charAt(0)}
                          </div>
                          <div>
                            <h4 className="font-semibold text-xs text-slate-900 dark:text-white">
                              {emp.name}
                            </h4>
                            <p className="font-mono text-slate-400 text-[10px]">
                              رقم: #{emp.id} · التزام: {emp.completionRate}%
                            </p>
                          </div>
                        </div>
                        <button
                          onClick={() => setSelectedEmployee(emp)}
                          className="px-2 py-0.5 rounded text-[11px] font-medium border border-slate-200 dark:border-slate-700 hover:bg-slate-50 text-slate-600 dark:text-slate-300"
                        >
                          معاينة السجل
                        </button>
                      </div>

                      <div className="space-y-1.5">
                        {emp.shifts.filter(s => s.status === 'ناقصة').map((shift, idx) => {
                          const found = shift.punches_details.filter(p => !p.is_duplicate).map(p => p.period);
                          const allPossible = ["دخول صباحي", "بصمة وسطية", "بصمة مسائية", "بصمة ليلية", "بصمة خروج"];
                          const missing = allPossible.filter(p => !found.includes(p));
                          return (
                            <div
                              key={idx}
                              className="flex items-center justify-between px-2.5 py-1.5 rounded bg-slate-50 dark:bg-slate-800/60 text-[11px] border border-slate-100 dark:border-slate-800"
                            >
                              <span className="font-mono font-medium text-slate-700 dark:text-slate-300">
                                {shift.shift_date}
                              </span>
                              <div className="flex flex-wrap gap-1 justify-end">
                                {missing.map(m => (
                                  <span
                                    key={m}
                                    className="px-1.5 py-0.2 rounded text-[10px] bg-rose-50 dark:bg-rose-950/50 text-rose-700 dark:text-rose-300 border border-rose-200/60 dark:border-rose-900/40"
                                  >
                                    {m}
                                  </span>
                                ))}
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    </div>
                  ))}
                </div>

              </div>
            )}

          </div>
        )}

      </main>

      {/* MODAL 1: Employee Shift Log Inspector */}
      {selectedEmployee && (
        <EmployeeDetailModal
          employee={selectedEmployee}
          requiredPunches={rules.REQUIRED_PUNCHES_PER_SHIFT}
          onClose={() => setSelectedEmployee(null)}
        />
      )}

      {/* MODAL 2: Settings & Period Time Windows */}
      {isSettingsOpen && (
        <SettingsModal
          rules={rules}
          onSave={handleSaveRules}
          onClose={() => setIsSettingsOpen(false)}
        />
      )}

    </div>
  );
};

/* --- SUB-COMPONENTS (Clean Enterprise Architecture) --- */

interface EmployeeModalProps {
  employee: EmployeeSummary;
  requiredPunches: number;
  onClose: () => void;
}

const EmployeeDetailModal: React.FC<EmployeeModalProps> = ({ employee, requiredPunches, onClose }) => {
  const [report, setReport] = useState<string | null>(null);
  const [loading, setLoading] = useState(false);

  const handleAuditInspection = async () => {
    setLoading(true);
    try {
      const res = await generateDetailedAuditSummaryAI(employee);
      setReport(res.text);
    } catch {
      setReport("تعذر إكمال الفحص الآلي.");
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-3 sm:p-6 bg-slate-900/40 backdrop-blur-xs">
      <div className="bg-white dark:bg-slate-900 w-full max-w-4xl max-h-[90vh] rounded-xl overflow-hidden flex flex-col shadow-lg border border-slate-200 dark:border-slate-800">
        
        {/* Modal Top Bar */}
        <div className="px-5 py-3.5 border-b border-slate-200 dark:border-slate-800 bg-slate-50/70 dark:bg-slate-800/70 flex items-center justify-between shrink-0">
          <div className="flex items-center gap-3">
            <div className="w-8 h-8 rounded-lg bg-slate-900 dark:bg-slate-700 text-white flex items-center justify-center font-bold text-xs">
              {employee.name.charAt(0)}
            </div>
            <div>
              <h2 className="font-bold text-sm text-slate-900 dark:text-white">
                {employee.name}
              </h2>
              <p className="text-[11px] font-mono text-slate-500 dark:text-slate-400">
                الرقم الوظيفي: #{employee.id} · نسبة الالتزام: {employee.completionRate}%
              </p>
            </div>
          </div>
          <button
            onClick={onClose}
            className="p-1 rounded-md border border-slate-200 dark:border-slate-700 text-slate-500 hover:bg-slate-200 dark:hover:bg-slate-700 transition-colors"
          >
            <X size={16} />
          </button>
        </div>

        {/* Modal Body */}
        <div className="flex-1 flex flex-col md:flex-row overflow-hidden">
          
          {/* Summary Column */}
          <div className="w-full md:w-60 p-4 border-b md:border-b-0 md:border-l border-slate-200 dark:border-slate-800 bg-slate-50/40 dark:bg-slate-950/40 overflow-y-auto space-y-3 shrink-0">
            <div className="p-3 rounded-lg border border-slate-200 dark:border-slate-800 bg-white dark:bg-slate-900 text-center">
              <span className="text-[10px] text-slate-400 block font-medium">معدل الانضباط</span>
              <span className="text-2xl font-bold font-mono text-blue-600 dark:text-blue-400 tabular-nums">
                {employee.completionRate}%
              </span>
            </div>

            <div className="grid grid-cols-2 gap-2 text-center text-xs">
              <div className="p-2 rounded-lg border border-slate-200 dark:border-slate-800 bg-white dark:bg-slate-900">
                <span className="text-[9px] text-slate-400 block">أيام النوبات</span>
                <span className="font-bold font-mono text-slate-800 dark:text-slate-200">
                  {employee.shifts.length}
                </span>
              </div>
              <div className="p-2 rounded-lg border border-slate-200 dark:border-slate-800 bg-white dark:bg-slate-900">
                <span className="text-[9px] text-slate-400 block">إجمالي البصمات</span>
                <span className="font-bold font-mono text-slate-800 dark:text-slate-200">
                  {employee.totalPunches}
                </span>
              </div>
            </div>

            {/* Quick Inspection Audit */}
            <div className="p-3 rounded-lg border border-slate-200 dark:border-slate-800 bg-white dark:bg-slate-900 space-y-2">
              <div className="flex items-center justify-between">
                <span className="text-[11px] font-bold text-slate-700 dark:text-slate-300">
                  فحص النواقص
                </span>
                <button
                  onClick={handleAuditInspection}
                  disabled={loading}
                  className="text-[10px] font-semibold text-blue-600 hover:underline"
                >
                  {loading ? 'جاري الفحص...' : 'تحديث'}
                </button>
              </div>
              <div className="text-[11px] font-mono leading-relaxed bg-slate-50 dark:bg-slate-800 p-2 rounded text-slate-600 dark:text-slate-300 whitespace-pre-line min-h-[60px] border border-slate-200/60 dark:border-slate-700/60">
                {report || "انقر لإجراء تدقيق فوري لمطابقة بصمات هذا الموظف."}
              </div>
            </div>
          </div>

          {/* Detailed Shift Logs Table */}
          <div className="flex-1 p-4 overflow-y-auto custom-scroll bg-white dark:bg-slate-900">
            <h3 className="text-xs font-bold text-slate-700 dark:text-slate-300 mb-3">
              سجل النوبات والبصمات المسجلة بالتاريخ
            </h3>

            <div className="space-y-2">
              {employee.shifts.map((shift, i) => (
                <div
                  key={i}
                  className={`p-3 rounded-lg border text-xs ${
                    shift.status === 'مكتملة'
                      ? 'border-slate-200 dark:border-slate-800 bg-slate-50/40 dark:bg-slate-800/30'
                      : 'border-rose-200/80 dark:border-rose-900/40 bg-rose-50/20 dark:bg-rose-950/20'
                  }`}
                >
                  <div className="flex items-center justify-between mb-2">
                    <div className="flex items-center gap-2">
                      <span className="font-bold font-mono text-slate-800 dark:text-slate-200">
                        {shift.shift_date}
                      </span>
                    </div>
                    <span className={`px-2 py-0.5 rounded text-[10px] font-semibold border ${
                      shift.status === 'مكتملة'
                        ? 'bg-emerald-50 text-emerald-700 border-emerald-200 dark:bg-emerald-950/40 dark:text-emerald-300'
                        : 'bg-rose-50 text-rose-700 border-rose-200 dark:bg-rose-950/40 dark:text-rose-300'
                    }`}>
                      {shift.status} ({shift.punches_count} من {requiredPunches})
                    </span>
                  </div>

                  {/* 5 Period Slots */}
                  <div className="grid grid-cols-5 gap-1.5 text-center">
                    {["دخول صباحي", "بصمة وسطية", "بصمة مسائية", "بصمة ليلية", "بصمة خروج"].map(period => {
                      const punch = shift.punches_details.find((pd: any) => pd.period === period && !pd.is_duplicate);
                      return (
                        <div
                          key={period}
                          className={`p-1.5 rounded border text-[11px] ${
                            punch
                              ? 'bg-white dark:bg-slate-800 border-slate-200 dark:border-slate-700 font-semibold'
                              : 'bg-slate-100/50 dark:bg-slate-800/20 border-transparent text-slate-400'
                          }`}
                        >
                          <span className="text-[9px] text-slate-400 block truncate">{period}</span>
                          <span className="font-mono tabular-nums">{punch ? punch.time : '--:--'}</span>
                        </div>
                      );
                    })}
                  </div>
                </div>
              ))}
            </div>
          </div>

        </div>
      </div>
    </div>
  );
};

interface SettingsModalProps {
  rules: Rules;
  onSave: (rules: Rules) => void;
  onClose: () => void;
}

const SettingsModal: React.FC<SettingsModalProps> = ({ rules, onSave, onClose }) => {
  const [localRules, setLocalRules] = useState<Rules>(rules);

  const handlePeriodChange = (periodName: string, field: 'start' | 'end', value: string) => {
    setLocalRules(prev => ({
      ...prev,
      TIME_PERIODS: {
        ...prev.TIME_PERIODS,
        [periodName]: {
          ...prev.TIME_PERIODS[periodName],
          [field]: value
        }
      }
    }));
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-3 bg-slate-900/40 backdrop-blur-xs">
      <div className="bg-white dark:bg-slate-900 w-full max-w-lg rounded-xl overflow-hidden shadow-lg border border-slate-200 dark:border-slate-800">
        
        <div className="px-5 py-3.5 border-b border-slate-200 dark:border-slate-800 bg-slate-50/70 dark:bg-slate-800/70 flex items-center justify-between">
          <div className="flex items-center gap-2">
            <Settings size={16} className="text-slate-700 dark:text-slate-300" />
            <h2 className="font-bold text-sm text-slate-900 dark:text-white">
              إعدادات قواعد البصمة والورديات
            </h2>
          </div>
          <button
            onClick={onClose}
            className="p-1 rounded-md border border-slate-200 dark:border-slate-700 text-slate-500 hover:bg-slate-200 dark:hover:bg-slate-700 transition-colors"
          >
            <X size={15} />
          </button>
        </div>

        <div className="p-5 space-y-4 text-xs">
          
          <div>
            <label className="block font-bold text-slate-700 dark:text-slate-300 mb-1.5">
              عدد البصمات الإلزامية لكل وردية
            </label>
            <div className="flex gap-2">
              <button
                type="button"
                onClick={() => setLocalRules({ ...localRules, REQUIRED_PUNCHES_PER_SHIFT: 3 })}
                className={`flex-1 py-1.5 px-3 rounded-lg font-bold border transition-colors ${
                  localRules.REQUIRED_PUNCHES_PER_SHIFT === 3
                    ? 'bg-blue-50 dark:bg-blue-950/50 border-blue-500 text-blue-700 dark:text-blue-300'
                    : 'bg-white dark:bg-slate-800 border-slate-200 dark:border-slate-700 text-slate-600 dark:text-slate-300'
                }`}
              >
                3 بصمات (أمر التخفيض)
              </button>
              <button
                type="button"
                onClick={() => setLocalRules({ ...localRules, REQUIRED_PUNCHES_PER_SHIFT: 5 })}
                className={`flex-1 py-1.5 px-3 rounded-lg font-bold border transition-colors ${
                  localRules.REQUIRED_PUNCHES_PER_SHIFT === 5
                    ? 'bg-blue-50 dark:bg-blue-950/50 border-blue-500 text-blue-700 dark:text-blue-300'
                    : 'bg-white dark:bg-slate-800 border-slate-200 dark:border-slate-700 text-slate-600 dark:text-slate-300'
                }`}
              >
                5 بصمات (الاعتيادي)
              </button>
            </div>
          </div>

          <div>
            <label className="block font-bold text-slate-700 dark:text-slate-300 mb-2">
              النوافذ الزمنية لفترات البصمات
            </label>
            <div className="space-y-2">
              {Object.entries(localRules.TIME_PERIODS).map(([periodName, times]) => (
                <div key={periodName} className="flex items-center justify-between p-2 rounded-lg border border-slate-200 dark:border-slate-800 bg-slate-50/50 dark:bg-slate-800/40">
                  <span className="font-semibold text-slate-800 dark:text-slate-200">{periodName}</span>
                  <div className="flex items-center gap-1.5 font-mono">
                    <input
                      type="time"
                      value={times.start}
                      onChange={(e) => handlePeriodChange(periodName, 'start', e.target.value)}
                      className="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-700 rounded px-1.5 py-0.5 text-xs text-slate-800 dark:text-slate-200"
                    />
                    <span className="text-slate-400">إلى</span>
                    <input
                      type="time"
                      value={times.end}
                      onChange={(e) => handlePeriodChange(periodName, 'end', e.target.value)}
                      className="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-700 rounded px-1.5 py-0.5 text-xs text-slate-800 dark:text-slate-200"
                    />
                  </div>
                </div>
              ))}
            </div>
          </div>

        </div>

        <div className="px-5 py-3 border-t border-slate-200 dark:border-slate-800 bg-slate-50 dark:bg-slate-800/60 flex justify-end gap-2">
          <button
            onClick={onClose}
            className="px-3 py-1.5 rounded-lg border border-slate-200 dark:border-slate-700 text-slate-600 dark:text-slate-300 hover:bg-slate-100 font-semibold text-xs transition-colors"
          >
            إلغاء
          </button>
          <button
            onClick={() => onSave(localRules)}
            className="px-4 py-1.5 rounded-lg bg-blue-600 hover:bg-blue-500 text-white font-semibold text-xs shadow-xs transition-colors"
          >
            حفظ وتطبيق القواعد
          </button>
        </div>

      </div>
    </div>
  );
};

export default App;
