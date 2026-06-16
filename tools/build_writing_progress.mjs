import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const rootDir = path.resolve(__dirname, "..");
const outputDir = path.join(rootDir, "outputs");
const outputPath = path.join(outputDir, "writing_progress_tracker.xlsx");

const workbook = Workbook.create();
const dashboard = workbook.worksheets.add("统计数据");
const log = workbook.worksheets.add("写作记录");

const today = new Date(2026, 5, 16);
const start = new Date(today);

function excelDate(d) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

function addMonths(d, count) {
  return new Date(d.getFullYear(), d.getMonth() + count, 1);
}

function fmtMonth(d) {
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
}

function monthName(d) {
  return `${d.getFullYear()}年${d.getMonth() + 1}月`;
}

const weekdays = ["日", "一", "二", "三", "四", "五", "六"];

for (const sheet of [dashboard, log]) {
  sheet.showGridLines = false;
}

dashboard.getRange("A1:N1").merge();
dashboard.getRange("A1").values = [["写作进程统计"]];
dashboard.getRange("A1").format = {
  fill: "#294C60",
  font: { bold: true, color: "#FFFFFF", size: 18 },
  horizontalAlignment: "center",
  verticalAlignment: "center",
};
dashboard.getRange("A1:N1").format.rowHeightPx = 36;

dashboard.getRange("A3:D3").values = [["指标", "从今天起12个月", "说明", ""]];
dashboard.getRange("A3:D3").format = {
  fill: "#E8F1F2",
  font: { bold: true, color: "#1F2933" },
  borders: { preset: "all", style: "thin", color: "#C8D5D9" },
};
dashboard.getRange("A4:C7").values = [
  ["记录天数", null, "写作记录表中已有日期的天数"],
  ["总字数", null, "字数列会自动汇总"],
  ["日均字数", null, "从今天起12个月每日平均写作字数"],
  ["目标完成率", null, "按“目标”列中 1 的比例计算"],
];
dashboard.getRange("B4:B7").formulas = [
  [`=COUNT('写作记录'!$A$2:$A$366)`],
  [`=SUM('写作记录'!$D$2:$D$366)`],
  [`=IFERROR(AVERAGE('写作记录'!$D$2:$D$366),0)`],
  [`=IFERROR(COUNTIF('写作记录'!$E$2:$E$366,1)/COUNT('写作记录'!$E$2:$E$366),0)`],
];
dashboard.getRange("A4:C7").format = {
  borders: { preset: "all", style: "thin", color: "#D6DEE2" },
  fill: "#FFFFFF",
};
dashboard.getRange("A4:A7").format.font = { bold: true, color: "#294C60" };
dashboard.getRange("B4:B6").format.numberFormat = "#,##0";
dashboard.getRange("B7").format.numberFormat = "0.0%";

dashboard.getRange("A10").values = [["字数频率分布"]];
dashboard.getRange("A10:C10").merge();
dashboard.getRange("A10:C10").format = {
  fill: "#294C60",
  font: { bold: true, color: "#FFFFFF" },
  horizontalAlignment: "center",
};
dashboard.getRange("A11:C11").values = [["字数区间", "频率", "说明"]];
dashboard.getRange("A11:C11").format = {
  fill: "#E8F1F2",
  font: { bold: true },
  borders: { preset: "all", style: "thin", color: "#C8D5D9" },
};

const bins = [
  ["0-499", 0, 499],
  ["500-999", 500, 999],
  ["1000-1499", 1000, 1499],
  ["1500-1999", 1500, 1999],
  ["2000-2499", 2000, 2499],
  ["2500-2999", 2500, 2999],
  ["3000+", 3000, 999999],
];
dashboard.getRange("A12:A18").values = bins.map(([label]) => [label]);
dashboard.getRange("C12:C18").values = bins.map(([, min, max]) => [`${min} 到 ${max === 999999 ? "以上" : max} 字`]);
dashboard.getRange("B12:B18").formulas = bins.map(([, min, max]) => [
  max === 999999
    ? `=COUNTIFS('写作记录'!$D$2:$D$366,">=${min}",'写作记录'!$D$2:$D$366,"<>")`
    : `=COUNTIFS('写作记录'!$D$2:$D$366,">=${min}",'写作记录'!$D$2:$D$366,"<=${max}",'写作记录'!$D$2:$D$366,"<>")`,
]);
dashboard.getRange("A11:C18").format.borders = { preset: "all", style: "thin", color: "#D6DEE2" };

dashboard.getRange("F10").values = [["每月目标完成率"]];
dashboard.getRange("F10:H10").merge();
dashboard.getRange("F10:H10").format = {
  fill: "#294C60",
  font: { bold: true, color: "#FFFFFF" },
  horizontalAlignment: "center",
};
dashboard.getRange("F11:H11").values = [["月份", "完成率", "完成天数/记录天数"]];
dashboard.getRange("F11:H11").format = {
  fill: "#E8F1F2",
  font: { bold: true },
  borders: { preset: "all", style: "thin", color: "#C8D5D9" },
};
const monthStart = new Date(today.getFullYear(), today.getMonth(), 1);
const monthRows = [];
for (let i = 0; i < 12; i++) {
  const d = addMonths(monthStart, i);
  const next = addMonths(d, 1);
  monthRows.push([fmtMonth(d), d, next, monthName(d)]);
}
dashboard.getRange("F12:F23").values = monthRows.map(([, , , label]) => [label]);
dashboard.getRange("G12:G23").formulas = monthRows.map(([, d, next]) => [
  `=IFERROR(COUNTIFS('写作记录'!$A$2:$A$366,">=${fmtMonth(d)}-01",'写作记录'!$A$2:$A$366,"<${fmtMonth(next)}-01",'写作记录'!$E$2:$E$366,1)/COUNTIFS('写作记录'!$A$2:$A$366,">=${fmtMonth(d)}-01",'写作记录'!$A$2:$A$366,"<${fmtMonth(next)}-01",'写作记录'!$E$2:$E$366,"<>"),0)`,
]);
dashboard.getRange("H12:H23").formulas = monthRows.map(([, d, next]) => [
  `=COUNTIFS('写作记录'!$A$2:$A$366,">=${fmtMonth(d)}-01",'写作记录'!$A$2:$A$366,"<${fmtMonth(next)}-01",'写作记录'!$E$2:$E$366,1)&"/"&COUNTIFS('写作记录'!$A$2:$A$366,">=${fmtMonth(d)}-01",'写作记录'!$A$2:$A$366,"<${fmtMonth(next)}-01",'写作记录'!$E$2:$E$366,"<>")`,
]);
dashboard.getRange("F11:H23").format.borders = { preset: "all", style: "thin", color: "#D6DEE2" };
dashboard.getRange("G12:G23").format.numberFormat = "0.0%";

const wordChart = dashboard.charts.add("bar", dashboard.getRange("A11:B18"));
wordChart.title = "从今天起12个月每日写作字数分布";
wordChart.hasLegend = false;
wordChart.xAxis = { axisType: "textAxis" };
wordChart.yAxis = { numberFormatCode: "0" };
wordChart.xAxis.title.text = "字数";
wordChart.yAxis.title.text = "频率";
wordChart.setPosition("A25", "N42");

const completionChart = dashboard.charts.add("line", dashboard.getRange("F11:G23"));
completionChart.title = "从今天起12个月每月写作目标完成率";
completionChart.hasLegend = false;
completionChart.xAxis = { axisType: "textAxis", textStyle: { fontSize: 9 } };
completionChart.yAxis = { numberFormatCode: "0%", min: 0, max: 1 };
completionChart.xAxis.title.text = "月份";
completionChart.yAxis.title.text = "写作目标完成率";
completionChart.setPosition("A44", "N62");

dashboard.getRange("A:N").format.font = { name: "Microsoft YaHei", size: 10 };
dashboard.getRange("A:A").format.columnWidthPx = 110;
dashboard.getRange("B:B").format.columnWidthPx = 80;
dashboard.getRange("C:C").format.columnWidthPx = 190;
dashboard.getRange("F:F").format.columnWidthPx = 100;
dashboard.getRange("G:G").format.columnWidthPx = 90;
dashboard.getRange("H:H").format.columnWidthPx = 130;

log.getRange("A1:G1").values = [["日期", "年份", "星期", "字数", "目标", "任务", "Comments"]];
log.getRange("A1:G1").format = {
  fill: "#294C60",
  font: { bold: true, color: "#FFFFFF" },
  horizontalAlignment: "center",
  borders: { preset: "all", style: "thin", color: "#294C60" },
};
const rows = [];
for (let i = 0; i < 365; i++) {
  const d = new Date(start);
  d.setDate(start.getDate() + i);
  rows.push([
    excelDate(d),
    d.getFullYear(),
    `星期${weekdays[d.getDay()]}`,
    null,
    null,
    null,
    null,
  ]);
}
log.getRange("A2:G366").values = rows;
log.getRange("A2:A366").format.numberFormat = "yyyy-mm-dd";
log.getRange("B2:E366").format.numberFormat = "0";
log.getRange("F2:G366").format.numberFormat = "@";
log.getRange("A1:G366").format.borders = { preset: "all", style: "thin", color: "#D6DEE2" };
log.getRange("A2:C366").format.fill = "#F7FAFA";
log.getRange("D2:G366").format.fill = "#FFFFFF";
log.getRange("A:G").format.font = { name: "Microsoft YaHei", size: 10 };
log.getRange("A:A").format.columnWidthPx = 105;
log.getRange("B:B").format.columnWidthPx = 70;
log.getRange("C:C").format.columnWidthPx = 75;
log.getRange("D:D").format.columnWidthPx = 85;
log.getRange("E:E").format.columnWidthPx = 85;
log.getRange("F:F").format.columnWidthPx = 360;
log.getRange("G:G").format.columnWidthPx = 260;
log.getRange("F2:G366").format.wrapText = true;
log.getRange("A2:G366").format.rowHeightPx = 22;
log.freezePanes.freezeRows(1);
const logTable = log.tables.add("A1:G366", true, "WritingLog");
logTable.style = "TableStyleMedium2";
log.getRange("E2:E366").dataValidation = { rule: { type: "list", values: ["0", "1"] } };
log.getRange("D2:D366").dataValidation = { rule: { type: "whole", operator: "greaterThanOrEqual", formula1: 0 } };
log.getRange("J1:K5").values = [
  ["填写说明", ""],
  ["字数", "每天开始和结束写作时记录差值"],
  ["目标", "0=未完成，1=完成"],
  ["任务", "如实写下当天推进的研究任务"],
  ["Comments", "补充备注、卡点、资料来源或下一步想法"],
];
log.getRange("J1:K1").merge();
log.getRange("J1:K1").format = {
  fill: "#E8F1F2",
  font: { bold: true, color: "#294C60" },
  horizontalAlignment: "center",
};
log.getRange("J2:K5").format = {
  fill: "#FFFFFF",
  borders: { preset: "all", style: "thin", color: "#D6DEE2" },
};
log.getRange("K2:K5").format.wrapText = true;
log.getRange("J2:K5").format.rowHeightPx = 28;
log.getRange("J:K").format.font = { name: "Microsoft YaHei", size: 10 };
log.getRange("J:J").format.columnWidthPx = 80;
log.getRange("K:K").format.columnWidthPx = 300;

const previewDash = await workbook.render({ sheetName: "统计数据", autoCrop: "all", scale: 1, format: "png" });
await fs.mkdir(outputDir, { recursive: true });
await fs.writeFile(path.join(outputDir, "writing_progress_dashboard_preview.png"), new Uint8Array(await previewDash.arrayBuffer()));
const previewLog = await workbook.render({ sheetName: "写作记录", range: "A1:K24", scale: 1, format: "png" });
await fs.writeFile(path.join(outputDir, "writing_progress_log_preview.png"), new Uint8Array(await previewLog.arrayBuffer()));

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 100 },
  summary: "formula error scan",
});
console.log(errors.ndjson);

const dashCheck = await workbook.inspect({
  kind: "table",
  range: "统计数据!A1:H23",
  include: "values,formulas",
  tableMaxRows: 25,
  tableMaxCols: 8,
  maxChars: 4000,
});
console.log(dashCheck.ndjson);

const xlsx = await SpreadsheetFile.exportXlsx(workbook);
await xlsx.save(outputPath);
console.log(outputPath);
