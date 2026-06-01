import { Router, type IRouter } from "express";
import multer from "multer";
import XLSX from "xlsx";
import { generatePptx, type CategoryData } from "../lib/pptxGenerator.js";

const router: IRouter = Router();
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 20 * 1024 * 1024 },
});

const EXCLUDED_STAGES = ["cancelado", "identificado"];
const DIRECTED_STAGES = [
  "concluido",
  "solucao experimentada",
];

function normalizeStr(s: string): string {
  return String(s)
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim();
}

function normCategory(s: string): string {
  return String(s)
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/_/g, " ")
    .trim();
}

router.post(
  "/radar/parse",
  upload.single("file") as any,
  async (req: any, res: any) => {
    try {
      if (!req.file) {
        return res.status(400).json({ error: "Nenhum arquivo enviado." });
      }

      const workbook = XLSX.read(req.file.buffer, { type: "buffer" });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const rows: any[][] = XLSX.utils.sheet_to_json(sheet, {
        header: 1,
        defval: "",
      }) as any[][];

      let headerRow = -1;
      const colMap = { categoria: -1, titulo: -1, etapa: -1 };

      for (let i = 0; i < Math.min(rows.length, 20); i++) {
        const row = rows[i];
        for (let j = 0; j < row.length; j++) {
          const h = normalizeStr(String(row[j] || ""));
          if (h === "categorias" || h === "categoria") colMap.categoria = j;
          if (h === "resumo" || h === "titulo") colMap.titulo = j;
          if (h === "status" || h === "etapa") colMap.etapa = j;
        }
        if (
          colMap.categoria >= 0 &&
          colMap.titulo >= 0 &&
          colMap.etapa >= 0
        ) {
          headerRow = i;
          break;
        }
      }

      if (headerRow < 0) {
        return res.status(400).json({
          error:
            "Colunas obrigatórias não encontradas. A planilha deve ter as colunas: CATEGORIAS, RESUMO, STATUS.",
        });
      }

      const projects: Array<{
        id: string;
        categoria: string;
        titulo: string;
        etapa: string;
        isDirected: boolean;
      }> = [];

      const categoryMap = new Map<
        string,
        { total: number; directed: number; displayName: string }
      >();

      for (let i = headerRow + 1; i < rows.length; i++) {
        const row = rows[i];
        const categoriaRaw = String(row[colMap.categoria] || "").trim();
        const titulo = String(row[colMap.titulo] || "").trim();
        const etapa = String(row[colMap.etapa] || "").trim();

        if (!categoriaRaw || !titulo || !etapa) continue;

        const categoria = normCategory(categoriaRaw);
        const etapaNorm = normalizeStr(etapa);

        if (EXCLUDED_STAGES.includes(etapaNorm)) continue;
        if (EXCLUDED_STAGES.some((ex) => etapaNorm.includes(ex))) continue;

        const isDirected = DIRECTED_STAGES.some((d) => etapaNorm.includes(d));

        const id = `${i}-${titulo}`;
        projects.push({ id, categoria, titulo, etapa, isDirected });

        if (!categoryMap.has(categoria)) {
          categoryMap.set(categoria, {
            total: 0,
            directed: 0,
            displayName: categoria.toUpperCase(),
          });
        }
        const catData = categoryMap.get(categoria)!;
        catData.total++;
        if (isDirected) catData.directed++;
      }

      const categories: CategoryData[] = Array.from(
        categoryMap.entries(),
      ).map(([name, data]) => ({
        name,
        displayName: data.displayName,
        total: data.total,
        directed: data.directed,
        directedPct: data.total > 0 ? data.directed / data.total : 0,
      }));

      const totalProjects = categories.reduce((s, c) => s + c.total, 0);
      const totalDirected = categories.reduce((s, c) => s + c.directed, 0);
      const totalInProgress = totalProjects - totalDirected;

      res.json({
        projects,
        categories,
        totalProjects,
        totalDirected,
        totalInProgress,
      });
    } catch (err) {
      req.log.error({ err }, "Error parsing Excel");
      res
        .status(500)
        .json({ error: "Erro ao processar a planilha. Verifique o formato." });
    }
  },
);

router.post("/radar/generate", async (req: any, res: any) => {
  try {
    const { categories, totalProjects, totalDirected, totalInProgress, highlights } =
      req.body;

    if (!categories || !Array.isArray(categories) || categories.length === 0) {
      return res.status(400).json({ error: "Dados de categorias inválidos." });
    }

    const buffer = await generatePptx({
      categories,
      totalProjects: Number(totalProjects) || 0,
      totalDirected: Number(totalDirected) || 0,
      totalInProgress: Number(totalInProgress) || 0,
      highlights: Array.isArray(highlights) ? highlights : [],
    });

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    );
    res.setHeader(
      "Content-Disposition",
      'attachment; filename="Radar_Inovacao.pptx"',
    );
    res.send(buffer);
  } catch (err) {
    req.log.error({ err }, "Error generating PPTX");
    res
      .status(500)
      .json({ error: "Erro ao gerar o PowerPoint." });
  }
});

export default router;
