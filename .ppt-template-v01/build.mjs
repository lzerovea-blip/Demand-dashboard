import fs from "node:fs/promises";
import path from "node:path";
import { Presentation, PresentationFile } from "@oai/artifact-tool";

const ROOT = "/Users/lin/Documents/解决方案分层分级";
const WORK = path.join(ROOT, ".ppt-template-v01", "output");
const OUTPUT = path.join(ROOT, "需求路标模板-共创版-v0.1.pptx");

const C = {
  ink: "#101820",
  navy: "#17232E",
  navy2: "#21313D",
  paper: "#F3F6F7",
  white: "#FFFFFF",
  muted: "#70808B",
  faint: "#DCE4E8",
  mint: "#85E8B8",
  mintDeep: "#30A978",
  blue: "#247CF0",
  orange: "#EE6B56",
  violet: "#6C5DD3",
  lilac: "#F0EEFF",
  paleMint: "#E9FAF2",
  paleBlue: "#EAF2FF",
  paleOrange: "#FFF0EC",
};

const FONT = "HarmonyOS Sans SC";

function box(slide, { x, y, w, h, fill = "none", line = "none", radius = false, name }) {
  return slide.shapes.add({
    geometry: radius ? "roundRect" : "rect",
    name,
    position: { left: x, top: y, width: w, height: h },
    fill,
    line: { style: "solid", fill: line, width: line === "none" ? 0 : 1 },
    ...(radius ? { borderRadius: "rounded-xl" } : {}),
  });
}

function ellipse(slide, { x, y, w, h, fill, name }) {
  return slide.shapes.add({
    geometry: "ellipse",
    name,
    position: { left: x, top: y, width: w, height: h },
    fill,
    line: { style: "solid", fill: "none", width: 0 },
  });
}

function textBox(slide, text, { x, y, w, h, size = 18, color = C.ink, bold = false, align = "left", fill = "none", line = "none", radius = false, name, italic = false }) {
  const shape = box(slide, { x, y, w, h, fill, line, radius, name });
  shape.text = text;
  shape.text.style = {
    fontFamily: FONT,
    fontSize: size,
    color,
    bold,
    italic,
    alignment: align,
  };
  return shape;
}

function line(slide, x, y, w, h, fill = C.faint, name) {
  return box(slide, { x, y, w, h, fill, name });
}

function addTopBar(slide, section, page, dark = false) {
  const fg = dark ? C.white : C.ink;
  const muted = dark ? "#AFC0CB" : C.muted;
  textBox(slide, "需求路标工作台", { x: 54, y: 31, w: 270, h: 25, size: 13, color: muted, bold: true, name: `brand-${page}` });
  textBox(slide, section, { x: 895, y: 31, w: 270, h: 25, size: 13, color: muted, bold: true, align: "right", name: `section-${page}` });
  textBox(slide, String(page).padStart(2, "0"), { x: 1172, y: 31, w: 55, h: 25, size: 13, color: fg, bold: true, align: "right", name: `page-${page}` });
}

function addTitle(slide, eyebrow, title, subtitle, dark = false) {
  const titleColor = dark ? C.white : C.ink;
  const subColor = dark ? "#B6C5CF" : C.muted;
  textBox(slide, eyebrow.toUpperCase(), { x: 64, y: 82, w: 520, h: 28, size: 13, color: C.mintDeep, bold: true, name: `eyebrow-${title}` });
  textBox(slide, title, { x: 64, y: 112, w: 1136, h: 70, size: 36, color: titleColor, bold: true, name: `title-${title}` });
  if (subtitle) textBox(slide, subtitle, { x: 64, y: 181, w: 1040, h: 38, size: 17, color: subColor, name: `subtitle-${title}` });
}

function addPill(slide, label, x, y, w, fill, color = C.ink) {
  return textBox(slide, label, { x, y, w, h: 32, size: 13, color, bold: true, align: "center", fill, line: "none", radius: true });
}

function addMetric(slide, x, label, value, accent) {
  textBox(slide, label, { x, y: 267, w: 238, h: 28, size: 14, color: C.muted, bold: true });
  textBox(slide, value, { x, y: 298, w: 238, h: 70, size: 45, color: C.ink, bold: true });
  line(slide, x, 383, 238, 6, accent);
}

function addRoadmapGrid(slide, { empty = false, health = false }) {
  const x0 = 166;
  const y0 = 255;
  const monthW = 164;
  const rowH = 106;
  const months = ["1月", "2月", "3月", "4月", "5月", "6月"];
  const levels = ["基础", "进阶", "高阶"];

  textBox(slide, "2027", { x: 64, y: 227, w: 88, h: 32, size: 17, color: C.ink, bold: true });
  levels.forEach((level, row) => {
    const accent = row === 0 ? C.orange : row === 1 ? C.blue : C.violet;
    ellipse(slide, { x: 76, y: y0 + 43 + row * rowH, w: 10, h: 10, fill: accent });
    textBox(slide, level, { x: 94, y: y0 + 31 + row * rowH, w: 60, h: 32, size: 15, color: C.ink, bold: true });
    months.forEach((month, col) => {
      box(slide, { x: x0 + col * monthW, y: y0 + row * rowH, w: monthW - 8, h: rowH - 8, fill: C.white, line: C.faint, radius: true, name: `cell-${row}-${col}` });
      if (row === 0) textBox(slide, month, { x: x0 + col * monthW + 12, y: y0 - 34, w: 96, h: 28, size: 13, color: C.muted, bold: true, align: "center" });
    });
  });

  if (health) {
    const x = x0 + 2 * monthW + 7;
    const y = y0 + 8;
    box(slide, { x, y, w: monthW - 22, h: rowH - 24, fill: C.ink, line: "none", radius: true, name: "sleep-card" });
    ellipse(slide, { x: x + 13, y: y + 13, w: 10, h: 10, fill: C.orange });
    textBox(slide, "睡眠", { x: x + 29, y: y + 7, w: 87, h: 27, size: 15, color: C.white, bold: true });
    textBox(slide, "2 条需求 · 2 人月", { x: x + 13, y: y + 37, w: 121, h: 25, size: 11, color: "#B9C5CE" });
  }

  if (empty) {
    textBox(slide, "暂无运动需求", { x: 469, y: 386, w: 500, h: 48, size: 23, color: C.muted, bold: true, align: "center" });
    textBox(slide, "录入运动基础 / 进阶 / 高阶来源后，将自动落入对应月份泳道", { x: 360, y: 438, w: 718, h: 40, size: 15, color: C.muted, align: "center" });
  }
}

async function writeBlob(filePath, blob) {
  await fs.writeFile(filePath, new Uint8Array(await blob.arrayBuffer()));
}

await fs.mkdir(WORK, { recursive: true });
const deck = Presentation.create({ slideSize: { width: 1280, height: 720 } });

// 01 — Cover
{
  const slide = deck.slides.add();
  slide.background.fill = C.ink;
  addTopBar(slide, "模板共创版 v0.1", 1, true);
  line(slide, 64, 94, 90, 7, C.mint);
  textBox(slide, "运动健康\n需求路标", { x: 64, y: 146, w: 720, h: 182, size: 55, color: C.white, bold: true, name: "cover-title" });
  textBox(slide, "从本地需求池，自动生成可汇报、可编辑、可复用的 PowerPoint", { x: 67, y: 344, w: 720, h: 52, size: 20, color: "#B8C6CE", name: "cover-subtitle" });
  addPill(slide, "离线处理", 67, 426, 114, C.paleMint, C.mintDeep);
  addPill(slide, "结构化汇总", 193, 426, 132, C.paleBlue, C.blue);
  addPill(slide, "模板确认后固定输出", 337, 426, 196, C.paleOrange, C.orange);

  box(slide, { x: 846, y: 135, w: 340, h: 428, fill: C.navy2, line: "#334654", radius: true, name: "cover-roadmap-preview" });
  textBox(slide, "2027 H1", { x: 879, y: 170, w: 150, h: 33, size: 16, color: C.mint, bold: true });
  textBox(slide, "健康路标", { x: 879, y: 206, w: 220, h: 47, size: 29, color: C.white, bold: true });
  [0, 1, 2].forEach((row) => {
    const c = row === 0 ? C.orange : row === 1 ? C.blue : C.violet;
    ellipse(slide, { x: 882, y: 288 + row * 75, w: 9, h: 9, fill: c });
    textBox(slide, ["基础", "进阶", "高阶"][row], { x: 901, y: 277 + row * 75, w: 54, h: 30, size: 13, color: "#B6C5CF", bold: true });
    line(slide, 968, 291 + row * 75, 179, 2, "#40515E");
  });
  box(slide, { x: 1001, y: 267, w: 116, h: 55, fill: C.white, line: "none", radius: true });
  textBox(slide, "睡眠", { x: 1016, y: 277, w: 85, h: 24, size: 14, color: C.ink, bold: true });
  textBox(slide, "2 人月", { x: 1016, y: 300, w: 85, h: 18, size: 10, color: C.muted });
  textBox(slide, "共创日期 · 2026.08.19", { x: 67, y: 641, w: 260, h: 24, size: 12, color: "#71828D" });
}

// 02 — Summary
{
  const slide = deck.slides.add();
  slide.background.fill = C.paper;
  addTopBar(slide, "汇总分析", 2);
  addTitle(slide, "SUMMARY", "一页看清投入、结构和节奏", "示例采用当前有效需求：社交时差、体动；测试数据不计入模板样例。", false);

  addMetric(slide, 64, "需求数量", "2 条", C.orange);
  addMetric(slide, 337, "工作量", "2 人月", C.mintDeep);
  addMetric(slide, 610, "健康来源占比", "100%", C.blue);
  addMetric(slide, 883, "体验优化占比", "100%", C.violet);

  textBox(slide, "上线节奏", { x: 64, y: 443, w: 200, h: 30, size: 16, color: C.ink, bold: true });
  line(slide, 96, 530, 1035, 5, C.faint);
  const points = ["1月", "2月", "3月", "4月", "5月", "6月"];
  points.forEach((month, index) => {
    const x = 96 + index * 207;
    ellipse(slide, { x: x - 8, y: 524, w: 17, h: 17, fill: index === 2 ? C.orange : C.white });
    if (index !== 2) box(slide, { x: x - 8, y: 524, w: 17, h: 17, fill: C.white, line: C.faint, radius: true });
    textBox(slide, month, { x: x - 34, y: 550, w: 70, h: 24, size: 13, color: index === 2 ? C.ink : C.muted, bold: index === 2, align: "center" });
  });
  textBox(slide, "2 条需求 · 2 人月", { x: 421, y: 477, w: 210, h: 35, size: 16, color: C.white, bold: true, align: "center", fill: C.ink, radius: true });
  textBox(slide, "汇总页建议固定保留：数量、工作量、来源、分类、月份节奏", { x: 64, y: 624, w: 720, h: 30, size: 14, color: C.muted });
}

// 03 — Health roadmap
{
  const slide = deck.slides.add();
  slide.background.fill = C.paper;
  addTopBar(slide, "健康路标", 3);
  addTitle(slide, "HEALTH ROADMAP", "健康需求集中在 2027 年 3 月的基础泳道", "同领域、同来源、同上线年月自动合并为一张路标卡片。", false);
  addPill(slide, "2 条需求", 907, 181, 108, C.white, C.ink);
  addPill(slide, "2 人月", 1027, 181, 102, C.white, C.ink);
  addRoadmapGrid(slide, { health: true });
  textBox(slide, "卡片标题、摘要可在工作台中单独编辑；详情页继续保留每条原始需求。", { x: 64, y: 626, w: 870, h: 30, size: 14, color: C.muted });
}

// 04 — Sports roadmap
{
  const slide = deck.slides.add();
  slide.background.fill = C.paper;
  addTopBar(slide, "运动路标", 4);
  addTitle(slide, "SPORT ROADMAP", "当前区间暂无运动需求，版式仍保持完整", "空泳道不删除，确保后续半年与跨版本输出的位置稳定。", false);
  addRoadmapGrid(slide, { empty: true });
  textBox(slide, "空状态也是模板的一部分：能明确表达“暂无”，而不是漏数或未生成。", { x: 64, y: 626, w: 820, h: 30, size: 14, color: C.muted });
}

// 05 — Detail
{
  const slide = deck.slides.add();
  slide.background.fill = C.white;
  addTopBar(slide, "领域详情", 5);
  addTitle(slide, "DOMAIN DETAIL", "睡眠领域包含 2 条可展开需求", "健康基础 · 2027年3月 · 工作量合计 2 人月", false);

  box(slide, { x: 64, y: 245, w: 742, h: 354, fill: C.paper, line: "none", radius: true, name: "detail-card-1" });
  ellipse(slide, { x: 92, y: 276, w: 13, h: 13, fill: C.orange });
  textBox(slide, "01", { x: 118, y: 266, w: 46, h: 30, size: 13, color: C.muted, bold: true });
  textBox(slide, "社交时差", { x: 169, y: 260, w: 350, h: 44, size: 26, color: C.ink, bold: true });
  addPill(slide, "体验优化", 624, 261, 116, C.lilac, C.violet);
  textBox(slide, "需求描述", { x: 92, y: 329, w: 120, h: 30, size: 14, color: C.muted, bold: true });
  textBox(slide, "评估周末和周内的睡眠规律性，帮助用户理解作息节奏差异。", { x: 92, y: 367, w: 638, h: 65, size: 18, color: C.ink });
  line(slide, 92, 458, 626, 1, C.faint);
  textBox(slide, "工作量", { x: 92, y: 481, w: 92, h: 26, size: 13, color: C.muted, bold: true });
  textBox(slide, "1 人月", { x: 92, y: 511, w: 100, h: 31, size: 18, color: C.ink, bold: true });
  textBox(slide, "上线", { x: 246, y: 481, w: 70, h: 26, size: 13, color: C.muted, bold: true });
  textBox(slide, "2027.03", { x: 246, y: 511, w: 100, h: 31, size: 18, color: C.ink, bold: true });

  box(slide, { x: 838, y: 245, w: 378, h: 231, fill: C.ink, line: "none", radius: true, name: "image-slot" });
  ellipse(slide, { x: 998, y: 303, w: 58, h: 58, fill: C.navy2 });
  textBox(slide, "+", { x: 1005, y: 307, w: 44, h: 44, size: 28, color: C.mint, bold: true, align: "center" });
  textBox(slide, "本地图片自动填充", { x: 899, y: 382, w: 256, h: 34, size: 17, color: C.white, bold: true, align: "center" });
  textBox(slide, "最多 5 张 · 保持比例裁切", { x: 899, y: 418, w: 256, h: 26, size: 12, color: "#AFC0CB", align: "center" });

  box(slide, { x: 838, y: 496, w: 378, h: 103, fill: C.paper, line: "none", radius: true, name: "detail-card-2" });
  textBox(slide, "02", { x: 864, y: 516, w: 42, h: 24, size: 12, color: C.muted, bold: true });
  textBox(slide, "体动", { x: 912, y: 508, w: 150, h: 32, size: 20, color: C.ink, bold: true });
  textBox(slide, "描述待补充", { x: 912, y: 547, w: 150, h: 25, size: 13, color: C.muted, italic: true });
  textBox(slide, "1 人月", { x: 1102, y: 523, w: 82, h: 30, size: 16, color: C.ink, bold: true, align: "right" });
  textBox(slide, "详情页规则：描述为空时保留“待补充”；有图时优先使用右侧视觉区。", { x: 64, y: 630, w: 920, h: 28, size: 14, color: C.muted });
}

// 06 — Confirmation
{
  const slide = deck.slides.add();
  slide.background.fill = C.ink;
  addTopBar(slide, "模板确认", 6, true);
  addTitle(slide, "TEMPLATE REVIEW", "确认这 4 类页面后，工作台将按同一规范持续输出", "v0.1 用于共同修改；确认后设为内置模板，不再要求导入 PPTX。", true);

  const cards = [
    ["01", "汇总分析", "指标与月份节奏"],
    ["02", "运动路标", "三级泳道与空状态"],
    ["03", "健康路标", "自动合并与卡片摘要"],
    ["04", "领域详情", "需求描述与本地图片"],
  ];
  cards.forEach(([num, title, desc], i) => {
    const x = 64 + i * 287;
    box(slide, { x, y: 271, w: 263, h: 147, fill: C.navy2, line: "#344754", radius: true });
    textBox(slide, num, { x: x + 22, y: 293, w: 44, h: 24, size: 13, color: C.mint, bold: true });
    textBox(slide, title, { x: x + 22, y: 326, w: 210, h: 34, size: 20, color: C.white, bold: true });
    textBox(slide, desc, { x: x + 22, y: 371, w: 210, h: 26, size: 13, color: "#AFC0CB" });
  });

  textBox(slide, "需要一起确认", { x: 64, y: 480, w: 220, h: 34, size: 18, color: C.white, bold: true });
  const checks = ["字体与颜色", "信息密度", "图片呈现", "详情页拆分规则"];
  checks.forEach((label, i) => {
    const x = 64 + i * 287;
    ellipse(slide, { x, y: 540, w: 20, h: 20, fill: C.mint });
    textBox(slide, label, { x: x + 34, y: 532, w: 210, h: 35, size: 16, color: C.white, bold: true });
  });
  textBox(slide, "下一版只改模板，不改历史需求数据。", { x: 64, y: 637, w: 540, h: 26, size: 13, color: "#81939F" });
}

for (const [index, slide] of deck.slides.items.entries()) {
  const stem = `slide-${String(index + 1).padStart(2, "0")}`;
  await writeBlob(path.join(WORK, `${stem}.png`), await deck.export({ slide, format: "png", scale: 1 }));
  const layout = await slide.export({ format: "layout" });
  await fs.writeFile(path.join(WORK, `${stem}.layout.json`), await layout.text());
}

await writeBlob(path.join(WORK, "deck-montage.webp"), await deck.export({ format: "webp", montage: true, scale: 1 }));
const pptx = await PresentationFile.exportPptx(deck);
await pptx.save(OUTPUT);
console.log(OUTPUT);
