import JSZip from "jszip";
import fs from "fs/promises";
import path from "path";
import { fileURLToPath } from "url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

const TEMPLATE_PATH = path.join(
  __dirname,
  "..",
  "..",
  "..",
  "attached_assets",
  "02_-_Comite_Inovação_e_Novos_Negocios_17_04_2026_1_1779743624123.pptx",
);

const CENTER_X = 6096000;
const CENTER_Y = 3429000;
const OUTER_RADIUS = 2880000;
const INNER_RADIUS = 1440000;
const MIN_BUBBLE_DIAM = 600000;
const MAX_BUBBLE_DIAM = 2700000;

const DYNAMIC_SHAPE_IDS = [15, 17, 20, 21, 22, 32, 47, 48, 8, 10];

export interface CategoryData {
  name: string;
  displayName: string;
  total: number;
  directed: number;
  directedPct: number;
}

export interface GenerateInput {
  categories: CategoryData[];
  totalProjects: number;
  totalDirected: number;
  totalInProgress: number;
  highlights: Array<{ title: string; category: string }>;
}

function escapeXml(str: string): string {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function removeShapesByIds(xml: string, ids: number[]): string {
  let result = xml;
  result = result.replace(/<p:cxnSp>[\s\S]*?<\/p:cxnSp>/g, "");
  for (const id of ids) {
    const searchFor = `<p:cNvPr id="${id}"`;
    const idPos = result.indexOf(searchFor);
    if (idPos < 0) continue;
    const before = result.slice(0, idPos);
    const lastSpStart = before.lastIndexOf("<p:sp>");
    if (lastSpStart < 0) continue;
    const afterSp = result.slice(lastSpStart);
    const closingIdx = afterSp.indexOf("</p:sp>");
    if (closingIdx < 0) continue;
    result =
      result.slice(0, lastSpStart) +
      result.slice(lastSpStart + closingIdx + "</p:sp>".length);
  }
  return result;
}

function updateShapeTextByName(
  xml: string,
  shapeName: string,
  newText: string,
): string {
  const marker = `name="${shapeName}"`;
  const namePos = xml.indexOf(marker);
  if (namePos < 0) return xml;
  const before = xml.slice(0, namePos);
  const lastSpStart = before.lastIndexOf("<p:sp>");
  if (lastSpStart < 0) return xml;
  const afterSp = xml.slice(lastSpStart);
  const closingIdx = afterSp.indexOf("</p:sp>");
  if (closingIdx < 0) return xml;
  const shapeBlock = afterSp.slice(0, closingIdx + "</p:sp>".length);
  const rest = afterSp.slice(closingIdx + "</p:sp>".length);
  const updatedBlock = shapeBlock.replace(
    /<a:t>([^<]*)<\/a:t>/,
    `<a:t>${escapeXml(newText)}</a:t>`,
  );
  return xml.slice(0, lastSpStart) + updatedBlock + rest;
}

function splitIntoLines(text: string, maxChars: number): string[] {
  const words = text.split(" ");
  const lines: string[] = [];
  let current = "";
  for (const word of words) {
    const candidate = current ? current + " " + word : word;
    if (candidate.length > maxChars && current) {
      lines.push(current);
      current = word;
    } else {
      current = candidate;
    }
  }
  if (current) lines.push(current);
  return lines;
}

function generateBubbleXml(
  id: number,
  centerX: number,
  centerY: number,
  diameter: number,
  count: number,
  categoryName: string,
): string {
  const x = Math.round(centerX - diameter / 2);
  const y = Math.round(centerY - diameter / 2);
  const cx = Math.round(diameter);
  const cy = Math.round(diameter);

  const countFontSize =
    diameter > 2200000
      ? 4800
      : diameter > 1600000
        ? 3600
        : diameter > 1100000
          ? 2800
          : diameter > 750000
            ? 2000
            : 1600;

  const nameFontSize =
    diameter > 2200000
      ? 1300
      : diameter > 1600000
        ? 1100
        : diameter > 1100000
          ? 900
          : diameter > 750000
            ? 700
            : 600;

  const maxCharsPerLine =
    diameter > 2200000
      ? 18
      : diameter > 1600000
        ? 14
        : diameter > 1100000
          ? 11
          : 9;

  const lines = splitIntoLines(categoryName, maxCharsPerLine);

  const nameParagraphs = lines
    .map(
      (line) =>
        `<a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="pt-BR" sz="${nameFontSize}" b="1" dirty="0"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:latin typeface="Montserrat" pitchFamily="2" charset="77"/></a:rPr><a:t>${escapeXml(line)}</a:t></a:r></a:p>`,
    )
    .join("");

  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="BubbleCat${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="060386"/></a:solidFill><a:ln w="38100"><a:solidFill><a:srgbClr val="060386"/></a:solidFill></a:ln></p:spPr><p:style><a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="15000"/></a:schemeClr></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style><p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="pt-BR" sz="${countFontSize}" b="1" dirty="0"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:latin typeface="Montserrat Black" pitchFamily="2" charset="77"/></a:rPr><a:t>${count}</a:t></a:r></a:p>${nameParagraphs}</p:txBody></p:sp>`;
}

function generateHighlightSection(
  id: number,
  x: number,
  y: number,
  cx: number,
  categoryName: string,
  projects: string[],
): string {
  const catPara = `<a:p><a:pPr><a:buNone/></a:pPr><a:r><a:rPr lang="pt-BR" sz="900" b="1" dirty="0"><a:solidFill><a:srgbClr val="060386"/></a:solidFill><a:latin typeface="Montserrat" pitchFamily="2" charset="77"/></a:rPr><a:t>${escapeXml(categoryName)}</a:t></a:r></a:p>`;
  const projectParas = projects
    .map(
      (p) =>
        `<a:p><a:pPr><a:buNone/></a:pPr><a:r><a:rPr lang="pt-BR" sz="800" b="1" dirty="0"><a:solidFill><a:srgbClr val="050193"/></a:solidFill><a:latin typeface="Montserrat" pitchFamily="2" charset="77"/></a:rPr><a:t>${escapeXml(p)}</a:t></a:r></a:p>`,
    )
    .join("");

  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Highlight${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="100000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr wrap="square" rtlCol="0"><a:spAutoFit/></a:bodyPr><a:lstStyle/>${catPara}${projectParas}</p:txBody></p:sp>`;
}

interface BubblePos {
  x: number;
  y: number;
  diameter: number;
}

function calculatePositions(categories: CategoryData[]): BubblePos[] {
  const n = categories.length;
  if (n === 0) return [];

  const counts = categories.map((c) => c.total);
  const minCount = Math.min(...counts);
  const maxCount = Math.max(...counts);

  return categories.map((cat, i) => {
    const angle = (2 * Math.PI * i) / n - Math.PI / 2 + Math.PI / 8;
    const directedPct = Math.min(1, Math.max(0, cat.directedPct));
    const d =
      INNER_RADIUS +
      (OUTER_RADIUS - INNER_RADIUS) * (1 - directedPct) +
      300000;

    let diameter: number;
    if (maxCount === minCount) {
      diameter = (MIN_BUBBLE_DIAM + MAX_BUBBLE_DIAM) / 2;
    } else {
      const ratio = (cat.total - minCount) / (maxCount - minCount);
      diameter = MIN_BUBBLE_DIAM + ratio * (MAX_BUBBLE_DIAM - MIN_BUBBLE_DIAM);
    }

    return {
      x: Math.round(CENTER_X + d * Math.cos(angle)),
      y: Math.round(CENTER_Y + d * Math.sin(angle)),
      diameter: Math.round(diameter),
    };
  });
}

function resolveOverlaps(positions: BubblePos[]): BubblePos[] {
  const result = positions.map((p) => ({ ...p }));
  for (let iter = 0; iter < 80; iter++) {
    let moved = false;
    for (let i = 0; i < result.length; i++) {
      for (let j = i + 1; j < result.length; j++) {
        const dx = result[j].x - result[i].x;
        const dy = result[j].y - result[i].y;
        const dist = Math.sqrt(dx * dx + dy * dy);
        const minDist = (result[i].diameter + result[j].diameter) / 2 + 180000;
        if (dist < minDist && dist > 1) {
          const push = (minDist - dist) / 2;
          const nx = dx / dist;
          const ny = dy / dist;
          result[i].x -= Math.round(nx * push);
          result[i].y -= Math.round(ny * push);
          result[j].x += Math.round(nx * push);
          result[j].y += Math.round(ny * push);
          moved = true;
        }
      }
    }
    if (!moved) break;
  }
  return result;
}

export async function generatePptx(input: GenerateInput): Promise<Buffer> {
  const templateBuffer = await fs.readFile(TEMPLATE_PATH);
  const zip = await (JSZip as any).loadAsync(templateBuffer);

  for (const slideFile of [
    "ppt/slides/slide1.xml",
    "ppt/slides/slide2.xml",
  ]) {
    const file = zip.file(slideFile);
    if (!file) continue;

    let xml: string = await file.async("string");
    const isSlide2 = slideFile.includes("slide2");

    xml = xml.replace(
      /(RADAR DE INOVAÇÃO E NOVOS NEGÓCIOS: )\d+( PROJETOS)/,
      `$1${input.totalProjects}$2`,
    );

    xml = updateShapeTextByName(
      xml,
      "Retângulo Arredondado 50",
      String(input.totalInProgress),
    );
    xml = updateShapeTextByName(
      xml,
      "Retângulo Arredondado 54",
      String(input.totalDirected),
    );

    xml = removeShapesByIds(xml, DYNAMIC_SHAPE_IDS);

    const positions = resolveOverlaps(calculatePositions(input.categories));
    let newShapes = "";
    let idCounter = 300;

    for (let i = 0; i < input.categories.length; i++) {
      const cat = input.categories[i];
      const pos = positions[i];
      newShapes += generateBubbleXml(
        idCounter++,
        pos.x,
        pos.y,
        pos.diameter,
        cat.total,
        cat.displayName,
      );
    }

    if (isSlide2 && input.highlights.length > 0) {
      const grouped = new Map<string, string[]>();
      for (const h of input.highlights) {
        if (!grouped.has(h.category)) grouped.set(h.category, []);
        grouped.get(h.category)!.push(h.title);
      }

      const entries = Array.from(grouped.entries());
      const leftEntries = entries.slice(0, Math.ceil(entries.length / 2));
      const rightEntries = entries.slice(Math.ceil(entries.length / 2));

      let leftY = 600000;
      for (const [catName, projects] of leftEntries) {
        if (leftY > 6200000) break;
        newShapes += generateHighlightSection(
          idCounter++,
          200000,
          leftY,
          2700000,
          catName,
          projects,
        );
        leftY += (projects.length + 1) * 310000 + 250000;
      }

      let rightY = 600000;
      for (const [catName, projects] of rightEntries) {
        if (rightY > 6200000) break;
        newShapes += generateHighlightSection(
          idCounter++,
          9400000,
          rightY,
          2600000,
          catName,
          projects,
        );
        rightY += (projects.length + 1) * 310000 + 250000;
      }
    }

    xml = xml.replace("</p:spTree>", newShapes + "</p:spTree>");
    zip.file(slideFile, xml);
  }

  const buffer = await zip.generateAsync({
    type: "nodebuffer",
    compression: "DEFLATE",
  });
  return buffer as Buffer;
}
