
import { GoogleGenAI } from "@google/genai";
import { KpiId } from '../types';

if (!process.env.API_KEY) {
  throw new Error("API_KEY environment variable not set");
}

const ai = new GoogleGenAI({ apiKey: process.env.API_KEY });

export const generateReport = async (
  selectedKpisData: Partial<Record<KpiId, { label: string; data: string }>>,
  audience: string
): Promise<string> => {
  const kpiEntries = Object.values(selectedKpisData);
  
  if (kpiEntries.length === 0) {
    return "錯誤：請至少選擇一個KPI並提供相關資料。";
  }

  const selectedKpiNames = kpiEntries.map(kpi => kpi.label).join(', ');

  const dataSection = kpiEntries
    .map(kpi => `
---
成果項目: ${kpi.label}
原始資料:
${kpi.data}
---
    `)
    .join('\n');

  const prompt = `
    你是一位專業的助理，專門為政府計畫撰寫進度與成果報告。你的任務是根據我提供的原始資料，生成一份結構完整、語氣正式的報告草稿。

    **報告基本資訊:**
    - **報告對象/單位:** ${audience || '未指定，請使用通用正式語氣'}
    - **涵蓋的成果項目 (KPIs):** ${selectedKpiNames}

    **原始資料:**
    ${dataSection}

    **撰寫要求:**
    1.  **結構化:** 請將提供的分散資料整合成一份連貫的報告。報告應包含簡短的引言、各個成果項目的分段說明，以及一個總結。
    2.  **語氣:** 使用符合政府單位要求的正式、專業且客觀的書面語氣。
    3.  **忠於原文:** 請僅使用我提供的原始資料進行撰寫，不要虛構任何不存在的資訊或數據。
    4.  **聚焦重點:** 報告內容需緊扣選定的成果項目。
    5.  **輸出格式:** 請直接提供完整的報告文案，使其可以直接被複製使用。不需要任何額外的標題或說明，例如「報告草稿：」。
    `;

  try {
    const response = await ai.models.generateContent({
        model: 'gemini-2.5-flash',
        contents: prompt
    });
    return response.text;
  } catch (error) {
    console.error("Gemini API call failed:", error);
    return "生成報告時發生錯誤，請檢查主控台以獲取詳細資訊。可能是API金鑰設定有誤或網路問題。";
  }
};
