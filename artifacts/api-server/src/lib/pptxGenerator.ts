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
  "Radar_Inovacao_(3)_1780409449539.pptx",
);

const SLIDE_W = 12192000;
const SLIDE_H = 6858000;
const CENTER_X = 6096000;
const CENTER_Y = 3429000;
const OUTER_RADIUS = 2880000;
const INNER_RADIUS = 1440000;
const MIN_BUBBLE_DIAM = 600000;
const MAX_BUBBLE_DIAM = 2700000;
const SMALL_BUBBLE_THRESHOLD = 900000;
const HIGHLIGHT_BOX_WIDTH = 2300000;
const HIGHLIGHT_GAP = 220000;

const DYNAMIC_SHAPE_IDS: number[] = [];

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
  showNameInside: boolean,
): string {
  const x = Math.round(centerX - diameter / 2);
  const y = Math.round(centerY - diameter / 2);
  const cx = Math.round(diameter);
  const cy = Math.round(diameter);

  const countFontSize =
    diameter > 2200000 ? 4800 :
    diameter > 1600000 ? 3600 :
    diameter > 1100000 ? 2800 :
    diameter > 750000  ? 2000 : 3200;

  const nameFontSize =
    diameter > 2200000 ? 1300 :
    diameter > 1600000 ? 1100 :
    diameter > 1100000 ? 900 :
    diameter > 750000  ? 700 : 600;

  const maxCharsPerLine =
    diameter > 2200000 ? 18 :
    diameter > 1600000 ? 14 :
    diameter > 1100000 ? 11 : 9;

  const nameParagraphs = showNameInside
    ? splitIntoLines(categoryName, maxCharsPerLine)
        .map(
          (line) =>
            `<a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="pt-BR" sz="${nameFontSize}" b="0" dirty="0">` +
            `<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>` +
            `<a:latin typeface="Montserrat" pitchFamily="2" charset="77"/></a:rPr>` +
            `<a:t>${escapeXml(line)}</a:t></a:r></a:p>`,
        )
        .join("")
    : "";

  return (
    `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="BubbleCat${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr>` +
    `<a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
    `<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>` +
    `<a:solidFill><a:srgbClr val="060386"/></a:solidFill>` +
    `<a:ln w="38100"><a:solidFill><a:srgbClr val="060386"/></a:solidFill></a:ln>` +
    `</p:spPr>` +
    `<p:style>` +
    `<a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="15000"/></a:schemeClr></a:lnRef>` +
    `<a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>` +
    `<a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>` +
    `<a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef>` +
    `</p:style>` +
    `<p:txBody>` +
    `<a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0" rtlCol="0" anchor="ctr"/>` +
    `<a:lstStyle/>` +
    `<a:p><a:pPr algn="ctr"/><a:r>` +
    `<a:rPr lang="pt-BR" sz="${countFontSize}" b="1" dirty="0">` +
    `<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>` +
    `<a:latin typeface="Montserrat Black" pitchFamily="2" charset="77"/>` +
    `</a:rPr><a:t>${count}</a:t></a:r></a:p>` +
    nameParagraphs +
    `</p:txBody></p:sp>`
  );
}

function generateCategoryLabel(
  id: number,
  x: number,
  y: number,
  categoryName: string,
): string {
  const W = 1600000;
  return (
    `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="CatLabel${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr>` +
    `<a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${W}" cy="200000"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
    `<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>` +
    `<a:ln><a:noFill/></a:ln>` +
    `</p:spPr>` +
    `<p:txBody>` +
    `<a:bodyPr wrap="square" lIns="91440" rIns="91440" tIns="45720" bIns="45720" rtlCol="0"><a:spAutoFit/></a:bodyPr>` +
    `<a:lstStyle/>` +
    `<a:p><a:pPr algn="ctr"><a:buNone/></a:pPr><a:r>` +
    `<a:rPr lang="pt-BR" sz="1000" b="0" dirty="0">` +
    `<a:solidFill><a:srgbClr val="060386"/></a:solidFill>` +
    `<a:latin typeface="Montserrat" pitchFamily="2" charset="77"/>` +
    `</a:rPr><a:t>${escapeXml(categoryName)}</a:t></a:r></a:p>` +
    `</p:txBody></p:sp>`
  );
}

function generateHighlightBox(
  id: number,
  x: number,
  y: number,
  width: number,
  categoryName: string,
  projects: string[],
): string {
  const catPara =
    `<a:p><a:pPr><a:buNone/></a:pPr><a:r>` +
    `<a:rPr lang="pt-BR" sz="900" b="1" dirty="0">` +
    `<a:solidFill><a:srgbClr val="060386"/></a:solidFill>` +
    `<a:latin typeface="Montserrat" pitchFamily="2" charset="77"/>` +
    `</a:rPr><a:t>${escapeXml(categoryName)}</a:t></a:r></a:p>`;

  const projectParas = projects
    .map(
      (p) =>
        `<a:p><a:pPr><a:buNone/></a:pPr><a:r>` +
        `<a:rPr lang="pt-BR" sz="800" dirty="0">` +
        `<a:solidFill><a:srgbClr val="1A1A1A"/></a:solidFill>` +
        `<a:latin typeface="Montserrat" pitchFamily="2" charset="77"/>` +
        `</a:rPr><a:t>\u2022 ${escapeXml(p)}</a:t></a:r></a:p>`,
    )
    .join("");

  return (
    `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Highlight${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr>` +
    `<a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${width}" cy="100000"/></a:xfrm>` +
    `<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 16667"/></a:avLst></a:prstGeom>` +
    `<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>` +
    `<a:ln w="38100"><a:solidFill><a:srgbClr val="060386"/></a:solidFill></a:ln>` +
    `</p:spPr>` +
    `<p:txBody>` +
    `<a:bodyPr wrap="square" lIns="91440" rIns="91440" tIns="0" bIns="0" rtlCol="0"><a:spAutoFit/></a:bodyPr>` +
    `<a:lstStyle/>` +
    catPara +
    projectParas +
    `</p:txBody></p:sp>`
  );
}

function generateArrow(
  id: number,
  x1: number,
  y1: number,
  x2: number,
  y2: number,
  startShapeId?: number,
  startCxnIdx?: number,
  endShapeId?: number,
  endCxnIdx?: number,
): string {
  const rx1 = Math.round(x1);
  const ry1 = Math.round(y1);
  const rx2 = Math.round(x2);
  const ry2 = Math.round(y2);
  const offX = Math.min(rx1, rx2);
  const offY = Math.min(ry1, ry2);
  const extCX = Math.max(1, Math.abs(rx2 - rx1));
  const extCY = Math.max(1, Math.abs(ry2 - ry1));
  const flipH = rx1 > rx2 ? ' flipH="1"' : "";
  const flipV = ry1 > ry2 ? ' flipV="1"' : "";

  const stCxn = startShapeId !== undefined
    ? `<a:stCxn id="${startShapeId}" idx="${startCxnIdx ?? 0}"/>`
    : "";
  const endCxn = endShapeId !== undefined
    ? `<a:endCxn id="${endShapeId}" idx="${endCxnIdx ?? 0}"/>`
    : "";

  return (
    `<p:cxnSp>` +
    `<p:nvCxnSpPr>` +
    `<p:cNvPr id="${id}" name="Arrow${id}"/>` +
    `<p:cNvCxnSpPr>${stCxn}${endCxn}</p:cNvCxnSpPr><p:nvPr/>` +
    `</p:nvCxnSpPr>` +
    `<p:spPr>` +
    `<a:xfrm${flipH}${flipV}><a:off x="${offX}" y="${offY}"/><a:ext cx="${extCX}" cy="${extCY}"/></a:xfrm>` +
    `<a:prstGeom prst="bentConnector3"><a:avLst/></a:prstGeom>` +
    `<a:noFill/>` +
    `<a:ln w="38100">` +
    `<a:solidFill><a:srgbClr val="060386"/></a:solidFill>` +
    `<a:tailEnd type="triangle" w="med" len="med"/>` +
    `</a:ln>` +
    `</p:spPr>` +
    `<p:style>` +
    `<a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef>` +
    `<a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef>` +
    `<a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>` +
    `<a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef>` +
    `</p:style>` +
    `</p:cxnSp>`
  );
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

    let diameter: number;
    if (maxCount === minCount) {
      diameter = (MIN_BUBBLE_DIAM + MAX_BUBBLE_DIAM) / 2;
    } else {
      const ratio = (cat.total - minCount) / (maxCount - minCount);
      diameter = MIN_BUBBLE_DIAM + ratio * (MAX_BUBBLE_DIAM - MIN_BUBBLE_DIAM);
    }

    const bubbleRadius = diameter / 2;

    // A distância do centro da bolha ao centro do radar é proporcional ao
    // % direcionado. Quanto maior o %, mais a bolha invade o círculo interno.
    // d = INNER_RADIUS + r * (1 - 2 * directedPct)
    //   0%  -> borda da bolha toca o círculo interno por fora
    //   50% -> centro da bolha na borda do círculo interno
    //   100% -> bolha inteira dentro do círculo interno
    let d = INNER_RADIUS + bubbleRadius * (1 - 2 * directedPct);

    // Garantir que a bolha caiba dentro do slide (não ultrapasse o outer radius)
    d = Math.min(d, OUTER_RADIUS - bubbleRadius);

    return {
      x: Math.round(CENTER_X + d * Math.cos(angle)),
      y: Math.round(CENTER_Y + d * Math.sin(angle)),
      diameter: Math.round(diameter),
    };
  });
}

interface ExclusionRect { x1: number; y1: number; x2: number; y2: number }

// Fixed areas that bubbles must not overlap (title bar, legend block)
const EXCLUSION_ZONES: ExclusionRect[] = [
  { x1: 0,       y1: 0,       x2: SLIDE_W, y2: 860000  }, // title strip (top)
  { x1: 8350000, y1: 0,       x2: SLIDE_W, y2: 1700000 }, // legend (top-right)
  { x1: 0,       y1: 5600000, x2: 3800000, y2: SLIDE_H }, // legend (bottom-left)
];

// Zone where highlight boxes must not be placed (bottom-left legend)
const HIGHLIGHT_LEGEND_ZONE: ExclusionRect = { x1: 0, y1: 5600000, x2: 3800000, y2: SLIDE_H };

function pushBubbleFromRect(
  bx: number, by: number, r: number,
  rect: ExclusionRect, margin: number,
): { dx: number; dy: number } {
  const closestX = Math.max(rect.x1, Math.min(bx, rect.x2));
  const closestY = Math.max(rect.y1, Math.min(by, rect.y2));
  const dx = bx - closestX;
  const dy = by - closestY;
  const dist = Math.sqrt(dx * dx + dy * dy);
  const minDist = r + margin;
  if (dist >= minDist) return { dx: 0, dy: 0 };
  if (dist < 1) {
    // Center is inside the rect — push downward (away from top)
    return { dx: 0, dy: minDist };
  }
  const push = minDist - dist;
  return { dx: (dx / dist) * push, dy: (dy / dist) * push };
}

function resolveOverlaps(
  positions: BubblePos[],
  categories: CategoryData[],
): BubblePos[] {
  const result = positions.map((p) => ({ ...p }));
  const MARGIN = 200000;

  for (let iter = 0; iter < 120; iter++) {
    let moved = false;

    // Bubble ↔ bubble repulsion
    for (let i = 0; i < result.length; i++) {
      for (let j = i + 1; j < result.length; j++) {
        const dx = result[j].x - result[i].x;
        const dy = result[j].y - result[i].y;
        const dist = Math.sqrt(dx * dx + dy * dy);
        const minDist = (result[i].diameter + result[j].diameter) / 2 + MARGIN;
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

    // Bubble ↔ exclusion zone repulsion
    for (let i = 0; i < result.length; i++) {
      const r = Math.round(result[i].diameter / 2);
      for (const zone of EXCLUSION_ZONES) {
        const { dx, dy } = pushBubbleFromRect(result[i].x, result[i].y, r, zone, MARGIN);
        if (dx !== 0 || dy !== 0) {
          result[i].x += Math.round(dx);
          result[i].y += Math.round(dy);
          moved = true;
        }
      }
    }

    // Clamp bubbles inside slide bounds
    for (let i = 0; i < result.length; i++) {
      const r = Math.round(result[i].diameter / 2);
      const clampedX = Math.max(r, Math.min(SLIDE_W - r, result[i].x));
      const clampedY = Math.max(r, Math.min(SLIDE_H - r, result[i].y));
      if (clampedX !== result[i].x || clampedY !== result[i].y) {
        result[i].x = clampedX;
        result[i].y = clampedY;
        moved = true;
      }
    }

    // Enforce directed-constraint: a distância do centro da bolha
    // ao centro do radar deve ser proporcional ao % direcionado.
    // Quanto maior o % direcionado, mais a bolha invade o círculo interno.
    for (let i = 0; i < result.length; i++) {
      const directedPct = Math.min(1, Math.max(0, categories[i].directedPct));
      if (categories[i].directed > 0) {
        const r = result[i].diameter / 2;
        // Centro da bolha deve estar a INNER_RADIUS + r*(1 - 2*pct)
        // Isso faz a proporção de invasão ser exatamente igual ao % direcionado
        const targetDist = INNER_RADIUS + r * (1 - 2 * directedPct);
        const dx = result[i].x - CENTER_X;
        const dy = result[i].y - CENTER_Y;
        const dist = Math.sqrt(dx * dx + dy * dy);
        if (dist > targetDist + 100000) {
          const nx = dx / dist;
          const ny = dy / dist;
          result[i].x = Math.round(CENTER_X + nx * targetDist);
          result[i].y = Math.round(CENTER_Y + ny * targetDist);
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

  const positions = resolveOverlaps(calculatePositions(input.categories), input.categories);

  const highlightsByCatIdx = new Map<number, string[]>();
  for (const h of input.highlights) {
    const idx = input.categories.findIndex(
      (c) =>
        c.displayName === h.category ||
        c.name.toUpperCase() === h.category,
    );
    if (idx < 0) continue;
    if (!highlightsByCatIdx.has(idx)) highlightsByCatIdx.set(idx, []);
    highlightsByCatIdx.get(idx)!.push(h.title);
  }

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

    let newShapes = "";
    let idCounter = 300;

    // Track the bubble shape ID for each category index (for arrow connections)
    const bubbleIds: number[] = [];

    for (let i = 0; i < input.categories.length; i++) {
      const cat = input.categories[i];
      const pos = positions[i];
      const isSmall = pos.diameter < SMALL_BUBBLE_THRESHOLD;

      const bubbleId = idCounter;
      bubbleIds.push(bubbleId);

      newShapes += generateBubbleXml(
        idCounter++,
        pos.x,
        pos.y,
        pos.diameter,
        cat.total,
        cat.displayName,
        !isSmall,
      );

      if (isSmall) {
        const onLeft = pos.x < CENTER_X;
        const LW = 1600000;
        const halfD = Math.round(pos.diameter / 2);
        const labelX = onLeft
          ? Math.round(Math.max(50000, pos.x - halfD - LW - 150000))
          : Math.round(Math.min(SLIDE_W - LW - 50000, pos.x + halfD + 150000));
        const labelY = Math.round(pos.y - 160000);
        newShapes += generateCategoryLabel(
          idCounter++,
          labelX,
          labelY,
          cat.displayName,
        );
      }
    }

    if (isSlide2 && highlightsByCatIdx.size > 0) {
      const MARGIN = 200000;
      const BOX_W = HIGHLIGHT_BOX_WIDTH;
      const LINE_H = 290000;

      for (const [catIdx, projects] of highlightsByCatIdx.entries()) {
        const pos = positions[catIdx];
        const cat = input.categories[catIdx];
        const bubbleR = Math.round(pos.diameter / 2);

        const boxH = 360000 + projects.length * LINE_H + 180000;

        const onLeft = pos.x < CENTER_X;
        const rawBoxX = onLeft
          ? pos.x - bubbleR - HIGHLIGHT_GAP - BOX_W
          : pos.x + bubbleR + HIGHLIGHT_GAP;
        const boxX = Math.round(Math.max(
          MARGIN,
          Math.min(SLIDE_W - BOX_W - MARGIN, rawBoxX),
        ));
        const rawBoxY = pos.y - Math.round(boxH / 2);
        let boxY = Math.round(Math.max(
          MARGIN + 350000,
          Math.min(SLIDE_H - boxH - MARGIN, rawBoxY),
        ));

        // Push box up if it would overlap the bottom-left legend zone
        if (
          boxX < HIGHLIGHT_LEGEND_ZONE.x2 &&
          boxX + BOX_W > HIGHLIGHT_LEGEND_ZONE.x1 &&
          boxY + boxH > HIGHLIGHT_LEGEND_ZONE.y1
        ) {
          boxY = Math.max(
            MARGIN + 350000,
            HIGHLIGHT_LEGEND_ZONE.y1 - boxH - MARGIN,
          );
        }

        const highlightBoxId = idCounter;
        newShapes += generateHighlightBox(
          idCounter++,
          boxX,
          boxY,
          BOX_W,
          cat.displayName,
          projects,
        );

        const bubbleId = bubbleIds[catIdx];
        // Connection point indices for ellipse/roundRect: 0=top,1=right,2=bottom,3=left
        // Arrow goes from bubble edge → highlight box edge
        const bubbleCxnIdx = onLeft ? 3 : 1;   // bubble left or right
        const boxCxnIdx = onLeft ? 1 : 3;       // box right or left

        const arrowStartX = onLeft ? pos.x - bubbleR : pos.x + bubbleR;
        const arrowStartY = pos.y;
        const arrowEndX = onLeft ? boxX + BOX_W : boxX;
        const arrowEndY = Math.round(boxY + boxH / 2);

        newShapes += generateArrow(
          idCounter++,
          arrowStartX,
          arrowStartY,
          arrowEndX,
          arrowEndY,
          bubbleId,
          bubbleCxnIdx,
          highlightBoxId,
          boxCxnIdx,
        );
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
