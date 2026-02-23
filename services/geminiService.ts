
import { GoogleGenAI, Type } from "@google/genai";

// Always use the named parameter and process.env.API_KEY for initialization
const ai = new GoogleGenAI({ apiKey: process.env.API_KEY });

export const analyzeCase = async (kronologi: string, kategori: string) => {
  try {
    // Using gemini-3-pro-preview for complex reasoning tasks like psychological analysis
    const response = await ai.models.generateContent({
      model: "gemini-3-pro-preview",
      contents: `Analisislah kasus bimbingan konseling berikut ini dan berikan saran penanganan yang edukatif, humanis, dan sesuai dengan etika konseling sekolah.
      
      Kategori: ${kategori}
      Kronologi: ${kronologi}
      
      Berikan respon dalam format JSON dengan struktur:
      {
        "analisis": "penjelasan singkat akar masalah",
        "saran": ["saran 1", "saran 2", "saran 3"],
        "pendekatan": "pendekatan psikologis yang disarankan"
      }`,
      config: {
        responseMimeType: "application/json",
        responseSchema: {
          type: Type.OBJECT,
          properties: {
            analisis: { type: Type.STRING },
            saran: { 
              type: Type.ARRAY, 
              items: { type: Type.STRING } 
            },
            pendekatan: { type: Type.STRING }
          },
          required: ["analisis", "saran", "pendekatan"]
        }
      }
    });

    // Directly access the .text property (not a method)
    const text = response.text;
    if (!text) return null;
    return JSON.parse(text.trim());
  } catch (error) {
    console.error("AI Analysis error:", error);
    return null;
  }
};
