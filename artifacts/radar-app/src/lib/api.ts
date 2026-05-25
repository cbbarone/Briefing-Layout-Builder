import type { ParseResult, GeneratePayload } from "./types";

const BASE = import.meta.env.BASE_URL.replace(/\/$/, "");
const API = `${BASE}/api`;

export async function parseExcel(file: File): Promise<ParseResult> {
  const form = new FormData();
  form.append("file", file);
  const res = await fetch(`${API}/radar/parse`, { method: "POST", body: form });
  if (!res.ok) {
    const data = await res.json().catch(() => ({ error: "Erro desconhecido" }));
    throw new Error(data.error || `Erro ${res.status}`);
  }
  return res.json();
}

export async function generatePptx(payload: GeneratePayload): Promise<Blob> {
  const res = await fetch(`${API}/radar/generate`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  if (!res.ok) {
    const data = await res.json().catch(() => ({ error: "Erro desconhecido" }));
    throw new Error(data.error || `Erro ${res.status}`);
  }
  return res.blob();
}

export function downloadBlob(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}
