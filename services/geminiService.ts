import { GoogleGenAI } from "@google/genai";
import { EmployeeSummary } from "../types";

const getAiClient = () => {
  // Use environment key or fallback if available
  const apiKey = (typeof process !== 'undefined' && process.env?.GEMINI_API_KEY) || 
                 (typeof import.meta !== 'undefined' && (import.meta as any).env?.VITE_GEMINI_API_KEY) ||
                 "";
  if (!apiKey) return null;
  return new GoogleGenAI({ apiKey });
};

export const generateDetailedAuditSummaryAI = async (employee: EmployeeSummary): Promise<{ text: string }> => {
  const missingDays = employee.shifts
    .filter(s => s.status === 'ناقصة')
    .map(s => {
      const found = s.punches_details.filter(p => !p.is_duplicate).map(p => p.period);
      const allPossible = ["دخول صباحي", "بصمة وسطية", "بصمة مسائية", "بصمة ليلية", "بصمة خروج"];
      const missing = allPossible.filter(p => !found.includes(p));
      return `تاريخ ${s.shift_date}: نقص في [${missing.join('، ')}]`;
    });

  if (missingDays.length === 0) {
    return { text: `الموظف: ${employee.name}\nالحالة: جميع سجلات البصمة مكتملة ومطابقة للقواعد المعتمدة.` };
  }

  try {
    const ai = getAiClient();
    if (!ai) {
      // Deterministic senior-grade fallback summary without external latency
      return {
        text: `الموظف: ${employee.name} (رقم: ${employee.id})\n` +
              `الحالة: توجد نواقص في ${missingDays.length} يوم عمل.\n\n` +
              `بيان النواقص التفصيلي:\n` +
              missingDays.map(d => `• ${d}`).join('\n')
      };
    }

    const prompt = `أنت مدقق أنظمة حضور وانصراف حكومية ومؤسسية. قدم تقريراً موجزاً ورصيناً للموظف (${employee.name}) بالصيغة التالية فقط بدون مقدمات:
الموظف: ${employee.name} (الرقم: ${employee.id})
الحالة: سجل به نواقص (${missingDays.length} أيام)
بيان النواقص:
${missingDays.join('\n')}`;

    const response = await ai.models.generateContent({
      model: 'gemini-2.5-flash',
      contents: prompt,
    });

    return { text: response.text || missingDays.join('\n') };
  } catch (error) {
    return {
      text: `الموظف: ${employee.name} (رقم: ${employee.id})\n` +
            `الحالة: توجد نواقص في ${missingDays.length} يوم عمل.\n\n` +
            missingDays.map(d => `• ${d}`).join('\n')
    };
  }
};

export const generateBulkAuditSummaryAI = async (employees: EmployeeSummary[]): Promise<{ text: string }> => {
  const deficientEmployees = employees.filter(e => !e.isEligible);
  
  if (deficientEmployees.length === 0) {
    return { text: "تقرير التدقيق الشامل:\nكافة موظفي الكادر المسجلين في هذا الملف منضبطون ومستوفون لجميع البصمات المقررة بدون أي استثناءات." };
  }

  const deterministicReport = deficientEmployees.map(e => {
    const missingDays = e.shifts.filter(s => s.status === 'ناقصة').map(s => {
      const found = s.punches_details.filter(p => !p.is_duplicate).map(p => p.period);
      const allPossible = ["دخول صباحي", "بصمة وسطية", "بصمة مسائية", "بصمة ليلية", "بصمة خروج"];
      const missing = allPossible.filter(p => !found.includes(p));
      return `    - ${s.shift_date}: غياب (${missing.join('، ')})`;
    }).join('\n');
    return `• ${e.name} [رقم ${e.id}] - نسبة الالتزام: ${e.completionRate}%\n${missingDays}`;
  }).join('\n\n');

  try {
    const ai = getAiClient();
    if (!ai) {
      return {
        text: `تقرير التدقيق الشامل (النواقص المكتشفة):\n` +
              `إجمالي الحالات غير المكتملة: ${deficientEmployees.length} من أصل ${employees.length} موظف.\n\n` +
              deterministicReport
      };
    }

    const prompt = `أنت مدقق أنظمة حضور وانصراف. لخص حالات النواقص التالية بشكل إداري منظم وموجز جداً:
${deterministicReport}`;

    const response = await ai.models.generateContent({
      model: 'gemini-2.5-flash',
      contents: prompt,
    });

    return { text: response.text || deterministicReport };
  } catch (error) {
    return {
      text: `تقرير التدقيق الشامل (النواقص المكتشفة):\n` +
            `إجمالي الحالات غير المكتملة: ${deficientEmployees.length} من أصل ${employees.length} موظف.\n\n` +
            deterministicReport
    };
  }
};
