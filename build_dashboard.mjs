import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = "outputs";
await fs.mkdir(outputDir, { recursive: true });

const rows = [
  ["Bidcom Fulfillment PREPARACION EN EE ENVIO STORE 68158112 29/05", "https://docs.google.com/spreadsheets/d/1m5Olc3LjrZah7EMihBCvP7y73N3dYeHRYxNsy5emTps/edit?gid=31416541#gid=31416541", "1m5Olc3LjrZah7EMihBCvP7y73N3dYeHRYxNsy5emTps", 3, 989, 0, 1, 5],
  ["Bidcom Fulfillment PREPARACION EN EE ENVIO BIDCOM #68160444 1/06", "https://docs.google.com/spreadsheets/d/163YF0pLeeVuxLu2qPo-XNsULU3NQiS2qethrH-VOBlM/edit?gid=31416541#gid=31416541", "163YF0pLeeVuxLu2qPo-XNsULU3NQiS2qethrH-VOBlM", 144, 5444, 4777, 1, 2],
  ["Preparacion Fulfillment EE ENVIO BIDCOM #68156370 29/05", "https://docs.google.com/spreadsheets/d/1uuSFNGIiP4tt5jrXVm57SK9DQzXOmn2IBOctdmG28c0/edit?gid=31416541#gid=31416541", "1uuSFNGIiP4tt5jrXVm57SK9DQzXOmn2IBOctdmG28c0", 8, 937, 0, 1, 6],
  ["PREPARACION FULL RENOVADO - BIDCOM STORE N68165989 1/06 08:00HS", "https://docs.google.com/spreadsheets/d/11_DPRMezgoiQkNOVSOA1c-95RsRu2D_FUwfYbGZGxoQ/edit?gid=31416541#gid=31416541", "11_DPRMezgoiQkNOVSOA1c-95RsRu2D_FUwfYbGZGxoQ", 56, 976, 187, 3, 3],
  ["FULL RENOVADO - ENVIO BIDCOM N68160444 1/06 08:30hs", "https://docs.google.com/spreadsheets/d/13TgDi51AMnBygpaYxGtzBo-GE8KDFwnJivcld6JVsr4/edit?gid=31416541#gid=31416541", "13TgDi51AMnBygpaYxGtzBo-GE8KDFwnJivcld6JVsr4", 55, 882, 0, 2, 4],
  ["Preparacion Fulfillment EE BIDCOM STORE #68165989 1/06", "https://docs.google.com/spreadsheets/d/1GDmkGykS2I8NUoZHgbaLDRGUz150bz_QL3AQpOnDfqQ/edit?gid=31416541#gid=31416541", "1GDmkGykS2I8NUoZHgbaLDRGUz150bz_QL3AQpOnDfqQ", 173, 5492, 266, 2, 25],
];

const wb = Workbook.create();
const data = wb.worksheets.add("Datos");
const dash = wb.worksheets.add("Dashboard");
const aux = wb.worksheets.add("Aux");

data.showGridLines = false;
dash.showGridLines = false;
aux.showGridLines = false;

const headers = ["Planillas", "URL", "ID", "SKU", "TOTAL", "PENDIENTES", "PROGRESO", "USUARIOS", "PALLETS"];
data.getRange("A1:I1").values = [headers];
data.getRange("A2:F7").values = rows.map((r) => r.slice(0, 6));
data.getRange("H2:I7").values = rows.map((r) => r.slice(6, 8));
data.getRange("G2").formulas = [["=IFERROR((E2-F2)/E2,0)"]];
data.getRange("G2:G7").fillDown();
data.getRange("A8:C8").values = [["", "", ""]];
data.getRange("D8").formulas = [["=SUM(D2:D7)"]];
data.getRange("E8").formulas = [["=SUM(E2:E7)"]];
data.getRange("F8").formulas = [["=SUM(F2:F7)"]];
data.getRange("G8").formulas = [["=IFERROR((E8-F8)/E8,0)"]];
data.getRange("H8").formulas = [["=SUM(H2:H7)"]];
data.getRange("I8").formulas = [["=SUM(I2:I7)"]];

data.getRange("A1:I1").format = {
  fill: "#075C45",
  font: { bold: true, color: "#FFFFFF" },
};
data.getRange("A8:I8").format = {
  fill: "#E8F5EF",
  font: { bold: true, color: "#0B1F3A" },
};
data.getRange("A1:I8").format.borders = {
  insideHorizontal: { style: "Continuous", color: "#D7DEE8" },
  insideVertical: { style: "Continuous", color: "#D7DEE8" },
  edgeBottom: { style: "Continuous", color: "#D7DEE8" },
  edgeTop: { style: "Continuous", color: "#D7DEE8" },
  edgeLeft: { style: "Continuous", color: "#D7DEE8" },
  edgeRight: { style: "Continuous", color: "#D7DEE8" },
};
data.getRange("G2:G8").format.numberFormat = "0.00%";
data.getRange("D2:F8").format.numberFormat = "#,##0";
data.getRange("H2:I8").format.numberFormat = "#,##0";
data.getRange("A:A").format.columnWidthPx = 390;
data.getRange("B:B").format.columnWidthPx = 120;
data.getRange("C:C").format.columnWidthPx = 160;
data.getRange("D:I").format.columnWidthPx = 92;
data.getRange("A1:I8").format.wrapText = true;
data.freezePanes.freezeRows(1);
data.tables.add("A1:I8", true, "TablaPlanillas");

dash.getRange("A1:N1").merge();
dash.getRange("A1").values = [["2. DASHBOARD DE PROGRESO (GOOGLE SHEETS)"]];
dash.getRange("A1").format = {
  fill: "#075C45",
  font: { bold: true, color: "#FFFFFF", size: 18 },
  horizontalAlignment: "Center",
  verticalAlignment: "Center",
};
dash.getRange("A1").format.rowHeightPx = 42;

dash.getRange("A3:N3").merge();
dash.getRange("A3").values = [["DASHBOARD - PROGRESO DE PEDIDOS"]];
dash.getRange("A3").format = {
  font: { bold: true, color: "#111C4E", size: 15 },
};

const kpiLabels = ["TOTAL PEDIDOS", "PICKEADOS TOTALES", "% PROGRESO GENERAL", "SKU TOTALES", "PALLETS"];
const kpiFormulas = ["=Datos!E8", "=Datos!E8-Datos!F8", "=Datos!G8", "=Datos!D8", "=Datos!I8"];
const kpiStarts = ["A5", "D5", "G5", "J5", "M5"];
for (let i = 0; i < kpiStarts.length; i += 1) {
  const col = kpiStarts[i].replace(/\d/g, "");
  const endCol = String.fromCharCode(col.charCodeAt(0) + 1);
  dash.getRange(`${col}5:${endCol}5`).merge();
  dash.getRange(`${col}5`).values = [[kpiLabels[i]]];
  dash.getRange(`${col}5`).format = {
    fill: "#FFFFFF",
    font: { bold: true, color: "#1F2937", size: 8 },
    horizontalAlignment: "Center",
    verticalAlignment: "Top",
  };
  dash.getRange(`${col}6:${endCol}7`).merge();
  dash.getRange(`${col}6`).formulas = [[kpiFormulas[i]]];
  dash.getRange(`${col}6`).format = {
    fill: "#FFFFFF",
    font: { bold: true, color: "#111827", size: 17 },
    horizontalAlignment: "Center",
    verticalAlignment: "Center",
  };
}
dash.getRange("G6:H6").format.numberFormat = "0.00%";
dash.getRange("A6:B6").format.numberFormat = "#,##0";
dash.getRange("D6:E6").format.numberFormat = "#,##0";
dash.getRange("J6:K6").format.numberFormat = "#,##0";
dash.getRange("M6:N6").format.numberFormat = "#,##0";

dash.getRange("A9:I9").merge();
dash.getRange("A9").values = [["PROGRESO POR PLANILLA"]];
dash.getRange("A10:D10").values = [["PLANILLA", "PICKEADOS", "% PROGRESO", "PEDIDOS"]];
dash.getRange("A11:D16").formulas = [
  ["=Datos!A2", "=Datos!E2-Datos!F2", "=Datos!G2", "=Datos!E2"],
  ["=Datos!A3", "=Datos!E3-Datos!F3", "=Datos!G3", "=Datos!E3"],
  ["=Datos!A4", "=Datos!E4-Datos!F4", "=Datos!G4", "=Datos!E4"],
  ["=Datos!A5", "=Datos!E5-Datos!F5", "=Datos!G5", "=Datos!E5"],
  ["=Datos!A6", "=Datos!E6-Datos!F6", "=Datos!G6", "=Datos!E6"],
  ["=Datos!A7", "=Datos!E7-Datos!F7", "=Datos!G7", "=Datos!E7"],
];
dash.getRange("A9:I16").format = { fill: "#FFFFFF" };
dash.getRange("A9").format = { font: { bold: true, color: "#111827", size: 11 } };
dash.getRange("A10:D10").format = {
  fill: "#F3F6FA",
  font: { bold: true, color: "#111827", size: 8 },
};
dash.getRange("A11:A16").format.wrapText = true;
dash.getRange("B11:D16").format.horizontalAlignment = "Center";
dash.getRange("B11:B16").format.numberFormat = "#,##0";
dash.getRange("C11:C16").format.numberFormat = "0.00%";
dash.getRange("D11:D16").format.numberFormat = "#,##0";

dash.getRange("K9:N16").format = { fill: "#FFFFFF" };
dash.getRange("K9:N9").merge();
dash.getRange("K9").values = [["QUIEN ES EL MEJOR?"]];
dash.getRange("K9").format = { font: { bold: true, color: "#111827", size: 11 }, horizontalAlignment: "Center" };
dash.getRange("K11:N11").merge();
dash.getRange("K11").values = [["TROFEO"]];
dash.getRange("K11").format = { font: { bold: true, color: "#F5B301", size: 30 }, horizontalAlignment: "Center" };
dash.getRange("K12:N12").merge();
dash.getRange("K12").formulas = [["=INDEX(Datos!A2:A7,MATCH(MAX(Datos!G2:G7),Datos!G2:G7,0))"]];
dash.getRange("K12").format = { font: { bold: true, color: "#111827", size: 12 }, horizontalAlignment: "Center", wrapText: true };
dash.getRange("K14:L14").merge();
dash.getRange("M14:N14").merge();
dash.getRange("K14").formulas = [["=MAX(Datos!E2:E7-Datos!F2:F7)"]];
dash.getRange("M14").formulas = [["=MAX(Datos!G2:G7)"]];
dash.getRange("K14:N14").format = { font: { bold: true, color: "#111827", size: 14 }, horizontalAlignment: "Center" };
dash.getRange("K15:L15").merge();
dash.getRange("M15:N15").merge();
dash.getRange("K15").values = [["PICKEADOS"]];
dash.getRange("M15").values = [["DE PROGRESO"]];
dash.getRange("K15:N15").format = { font: { bold: true, color: "#111827", size: 8 }, horizontalAlignment: "Center" };
dash.getRange("K14:L14").format.numberFormat = "#,##0";
dash.getRange("M14:N14").format.numberFormat = "0.00%";

dash.getRange("A18:H18").merge();
dash.getRange("A18").values = [["PEDIDOS POR PLANILLA"]];
dash.getRange("A18").format = { font: { bold: true, color: "#111827", size: 11 } };
dash.getRange("J18:N18").merge();
dash.getRange("J18").values = [["ESTADO DE PEDIDOS"]];
dash.getRange("J18").format = { font: { bold: true, color: "#111827", size: 11 } };

aux.getRange("A1:B4").values = [
  ["Estado", "Cantidad"],
  ["Completados", null],
  ["Pendientes", null],
  ["Total", null],
];
aux.getRange("B2").formulas = [["=Datos!E8-Datos!F8"]];
aux.getRange("B3").formulas = [["=Datos!F8"]];
aux.getRange("B4").formulas = [["=Datos!E8"]];

aux.getRange("D1:F7").values = [["Planilla", "Pedidos", "Pendientes"], ...rows.map((r, idx) => [`P${idx + 1}`, r[4], r[5]])];

const bar = dash.charts.add("bar", aux.getRange("D1:F7"));
bar.title = "Pedidos y pendientes";
bar.hasLegend = true;
bar.xAxis = { axisType: "textAxis" };
bar.yAxis = { numberFormatCode: "#,##0" };
bar.setPosition("A19", "H33");

const doughnut = dash.charts.add("doughnut", aux.getRange("A1:B3"));
doughnut.title = "Estado de pedidos";
doughnut.hasLegend = true;
doughnut.setPosition("J19", "N33");

dash.getRange("A5:N16").format.borders = {
  insideHorizontal: { style: "Continuous", color: "#E5EAF1" },
  insideVertical: { style: "Continuous", color: "#E5EAF1" },
  edgeBottom: { style: "Continuous", color: "#D7DEE8" },
  edgeTop: { style: "Continuous", color: "#D7DEE8" },
  edgeLeft: { style: "Continuous", color: "#D7DEE8" },
  edgeRight: { style: "Continuous", color: "#D7DEE8" },
};
dash.getRange("A18:N33").format.borders = {
  edgeBottom: { style: "Continuous", color: "#D7DEE8" },
  edgeTop: { style: "Continuous", color: "#D7DEE8" },
  edgeLeft: { style: "Continuous", color: "#D7DEE8" },
  edgeRight: { style: "Continuous", color: "#D7DEE8" },
};
dash.getRange("A:N").format.columnWidthPx = 78;
dash.getRange("A:A").format.columnWidthPx = 120;
dash.getRange("B:B").format.columnWidthPx = 70;
dash.getRange("D:D").format.columnWidthPx = 78;
dash.getRange("G:G").format.columnWidthPx = 80;
dash.getRange("J:J").format.columnWidthPx = 78;
dash.getRange("M:M").format.columnWidthPx = 78;
dash.getRange("A1:N33").format.fill = "#F7F9FC";
dash.getRange("A5:N7").format.fill = "#FFFFFF";
dash.getRange("A9:I16").format.fill = "#FFFFFF";
dash.getRange("K9:N16").format.fill = "#FFFFFF";
dash.getRange("A18:N33").format.fill = "#FFFFFF";
dash.getRange("A1:N1").format = {
  fill: "#075C45",
  font: { bold: true, color: "#FFFFFF", size: 18 },
  horizontalAlignment: "Center",
  verticalAlignment: "Center",
};
dash.getRange("A3:N3").format = {
  fill: "#F7F9FC",
  font: { bold: true, color: "#111C4E", size: 15 },
};

const preview = await wb.render({ sheetName: "Dashboard", autoCrop: "all", scale: 1, format: "png" });
await fs.writeFile(`${outputDir}/dashboard_preview.png`, new Uint8Array(await preview.arrayBuffer()));

const errors = await wb.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 100 },
  summary: "formula error scan",
});
console.log(errors.ndjson);

const exported = await SpreadsheetFile.exportXlsx(wb);
await exported.save(`${outputDir}/dashboard_progreso_pedidos.xlsx`);
console.log(`${outputDir}/dashboard_progreso_pedidos.xlsx`);
