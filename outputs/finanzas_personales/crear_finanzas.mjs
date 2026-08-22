import fs from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outDir = fileURLToPath(new URL("./", import.meta.url));
const wb = Workbook.create();
const inicio = wb.worksheets.add("Inicio");
const config = wb.worksheets.add("Configuración");
const registro = wb.worksheets.add("Registro diario");
const pagos = wb.worksheets.add("Pagos fijos");
const resumen = wb.worksheets.add("Resumen mensual");

wb.comments.setSelf({ displayName: "Vanessa M.V." });

const C = {
  rose: "#C9828B", roseDark: "#9E5F69", blush: "#F5D9DD", pale: "#FFF4F5",
  beige: "#E9DCCB", cream: "#FBF7F1", taupe: "#806D62", cocoa: "#4F413B",
  green: "#8FAF9B", greenLight: "#E5F0E8", gold: "#D6A85F", red: "#C75C5C",
  white: "#FFFFFF", gray: "#6E6763", line: "#E4D8D0", input: "#FFF2CC"
};
const money = '"S/ "#,##0.00;[Red]("S/ "#,##0.00);-';
const dateFmt = "dd-mmm-yyyy";
const monthFmt = "mmm yyyy";
const title = (sheet, range, text) => {
  range.merge(); range.values = [[text]];
  range.format = { fill: C.roseDark, font: { bold: true, color: C.white, size: 18 },
    horizontalAlignment: "left", verticalAlignment: "center" };
  range.format.rowHeight = 34;
};
const section = (range) => {
  range.format = { fill: C.beige, font: { bold: true, color: C.cocoa },
    borders: { preset: "outside", style: "thin", color: C.line } };
};
const header = (range) => {
  range.format = { fill: C.rose, font: { bold: true, color: C.white },
    horizontalAlignment: "center", verticalAlignment: "center", wrapText: true,
    borders: { preset: "inside", style: "thin", color: "#EBC5CA" } };
};
const card = (range, fill=C.pale) => {
  range.format = { fill, borders: { preset: "outside", style: "thin", color: C.line },
    font: { color: C.cocoa }, verticalAlignment: "center" };
};
const setWidths = (sheet, widths) => {
  for (const [col, px] of Object.entries(widths)) sheet.getRange(`${col}:${col}`).format.columnWidthPx = px;
};

for (const s of [inicio, config, registro, pagos, resumen]) s.showGridLines = false;

// CONFIGURACIÓN
title(config, config.getRange("A1:F1"), "Mi configuración financiera");
config.getRange("A2:F2").merge();
config.getRange("A2").values = [["Celdas amarillas = datos que puedes cambiar. Los cálculos se actualizarán solos."]];
config.getRange("A2:F2").format = { fill: C.cream, font: { italic: true, color: C.gray }, wrapText: true };

config.getRange("A4:C4").values = [["Entradas mensuales", "Monto", "Uso"]]; header(config.getRange("A4:C4"));
config.getRange("A5:C6").values = [
  ["Sueldo", 3233, "Efectivo disponible"],
  ["Tarjeta de alimentos", 200, "Solo alimentación"]
];
config.getRange("B5:B6").format = { fill: C.input, font: { color: "#0000FF" }, numberFormat: money };
config.getRange("A7:C7").values = [["Total de recursos", null, "No todo es efectivo"]];
config.getRange("B7").formulas = [["=SUM(B5:B6)"]]; section(config.getRange("A7:C7")); config.getRange("B7").format.numberFormat = money;

config.getRange("A9:F9").values = [["Tipo", "Concepto", "Presupuesto mensual", "Día(s) de vencimiento", "Inicio", "Nota"]]; header(config.getRange("A9:F9"));
const fixed = [
  ["Gasto fijo", "Luz", 100, "3", new Date(2026,7,1), "Monto aproximado"],
  ["Gasto fijo", "Celular Claro", 31, "14", new Date(2026,7,1), "Mensual"],
  ["Gasto fijo", "Celular Entel", 84, "1", new Date(2026,7,1), "Mensual"],
  ["Gasto fijo", "Gas", 100, "1 y 15", new Date(2026,7,1), "S/ 50 cada 15 días"],
  ["Gasto fijo", "Internet", 75, "1", new Date(2026,7,1), "Mensual"],
  ["Deuda", "Préstamo 1", 851.80, "3", new Date(2026,7,1), "Mensual"],
  ["Deuda", "Préstamo 2", 489.50, "3", new Date(2026,9,1), "Empieza en octubre de 2026"],
  ["Deuda", "Tarjeta", 250, "20", new Date(2026,7,1), "Mensual"]
];
config.getRange("A10:F17").values = fixed;
config.getRange("C10:E17").format = { fill: C.input, font: { color: "#0000FF" } };
config.getRange("C10:C17").format.numberFormat = money; config.getRange("E10:E17").format.numberFormat = dateFmt;
config.getRange("A18:F18").values = [["Total vigente desde octubre", "", null, "", "", ""]];
config.getRange("C18").formulas = [["=SUM(C10:C17)"]]; section(config.getRange("A18:F18")); config.getRange("C18").format.numberFormat = money;

config.getRange("A20:B20").values = [["Categoría variable", "Presupuesto mensual sugerido"]]; header(config.getRange("A20:B20"));
config.getRange("A21:B26").values = [
  ["Pasaje",250],["Comida",200],["Ropa",80],["Viajes o paseos",100],["Regalos",50],["Otros",50]
];
config.getRange("B21:B26").format = { fill: C.input, font: { color: "#0000FF" }, numberFormat: money };
config.getRange("A27:B27").values = [["Total variable sugerido", null]]; config.getRange("B27").formulas = [["=SUM(B21:B26)"]]; section(config.getRange("A27:B27")); config.getRange("B27").format.numberFormat = money;
config.getRange("D20:E20").values = [["Meta", "Valor"]]; header(config.getRange("D20:E20"));
config.getRange("D21:E22").values = [["Ahorro mínimo mensual",323.30],["Mes inicial",new Date(2026,7,1)]];
config.getRange("D23").values = [["Mes seleccionado"]]; config.getRange("E23").formulas = [["='Inicio'!B4"]];
config.getRange("E21:E22").format = { fill: C.input, font: { color: "#0000FF" } };
config.getRange("E23").format = { fill: C.cream, font: { color: "#008000" } };
config.getRange("E21").format.numberFormat = money; config.getRange("E22:E23").format.numberFormat = monthFmt;
config.getRange("D25:F25").merge(); config.getRange("D25").values = [["Presupuesto conservador pensado para que desde octubre puedas ahorrar y conservar un pequeño margen. Ajusta los importes amarillos según tu realidad."]];
config.getRange("D25:F25").format = { fill: C.greenLight, font: { color: C.cocoa }, wrapText: true }; config.getRange("D25:F25").format.rowHeight = 56;
setWidths(config,{A:130,B:140,C:145,D:150,E:115,F:190}); config.freezePanes.freezeRows(3);

// Comments for user-provided inputs
for (const cell of ["B5","B6","C10","C11","C12","C13","C14","C15","C16","C17"]) {
  wb.comments.addThread({ cell: config.getRange(cell) }, "Fuente: información proporcionada por Vanessa en esta conversación (10-ago-2026). El valor es editable.");
}

// REGISTRO DIARIO
title(registro, registro.getRange("A1:G1"), "Registro diario de gastos variables");
registro.getRange("A2:G2").merge(); registro.getRange("A2").values = [["Desde el celular: agrega una fila por cada gasto. Elige categoría y medio de pago; no necesitas tocar las fórmulas."]];
registro.getRange("A2:G2").format = { fill: C.cream, font: { color: C.gray }, wrapText: true }; registro.getRange("A2:G2").format.rowHeight = 40;
registro.getRange("A4:G4").values = [["Fecha", "Mes", "Categoría", "Descripción", "Monto", "Medio de pago", "Nota"]]; header(registro.getRange("A4:G4"));
registro.getRange("B5").formulas = [["=IF(A5=\"\",\"\",DATE(YEAR(A5),MONTH(A5),1))"]]; registro.getRange("B5:B504").fillDown();
registro.getRange("A5:A504").format.numberFormat = dateFmt; registro.getRange("B5:B504").format.numberFormat = monthFmt;
registro.getRange("E5:E504").format.numberFormat = money;
registro.getRange("A5:A504").format.fill = C.input; registro.getRange("C5:G504").format.fill = C.input;
registro.getRange("B5:B504").format = { fill: C.cream, font: { color: C.gray }, numberFormat: monthFmt };
registro.getRange("C5:C504").dataValidation = { rule: { type: "list", formula1: "'Configuración'!$A$21:$A$26" } };
registro.getRange("F5:F504").dataValidation = { rule: { type: "list", values: ["Efectivo/Débito", "Tarjeta alimentos", "Tarjeta crédito"] } };
registro.tables.add("A4:G504", true, "RegistroDiario").style = "TableStyleMedium2";
setWidths(registro,{A:105,B:90,C:125,D:190,E:100,F:145,G:190}); registro.freezePanes.freezeRows(4);

// PAGOS FIJOS (12 meses x 8 conceptos)
title(pagos, pagos.getRange("A1:H1"), "Control mensual de pagos fijos y deudas");
pagos.getRange("A2:H2").merge(); pagos.getRange("A2").values = [["Completa Monto real, Estado y Fecha pagada. Si el monto real queda vacío, el informe no lo contará como pagado."]];
pagos.getRange("A2:H2").format = { fill: C.cream, font: { color: C.gray }, wrapText: true }; pagos.getRange("A2:H2").format.rowHeight = 40;
pagos.getRange("A4:H4").values = [["Mes", "Tipo", "Concepto", "Presupuesto", "Monto real", "Vencimiento", "Estado", "Fecha pagada"]]; header(pagos.getRange("A4:H4"));
const rows=[];
for(let m=0;m<12;m++){
  const month = new Date(2026,7+m,1);
  for(const item of fixed){
    const active = month >= item[4];
    let dueDay = Number.parseInt(item[3],10) || 1;
    rows.push([month,item[0],item[1],active?item[2]:0,null,new Date(month.getFullYear(),month.getMonth(),dueDay),active?"Pendiente":"No aplica",null]);
  }
}
pagos.getRange("A5:H100").values = rows;
pagos.getRange("A5:A100").format.numberFormat = monthFmt; pagos.getRange("D5:E100").format.numberFormat = money;
pagos.getRange("F5:F100").format.numberFormat = dateFmt; pagos.getRange("H5:H100").format.numberFormat = dateFmt;
pagos.getRange("E5:E100").format.fill = C.input; pagos.getRange("G5:H100").format.fill = C.input;
pagos.getRange("G5:G100").dataValidation = { rule: { type: "list", values: ["Pendiente", "Pagado", "No aplica"] } };
pagos.getRange("G5:G100").conditionalFormats.add("containsText", { text:"Pagado", format:{ fill:C.greenLight, font:{color:"#356346",bold:true} } });
pagos.getRange("G5:G100").conditionalFormats.add("containsText", { text:"Pendiente", format:{ fill:"#FCE8E6", font:{color:C.red,bold:true} } });
pagos.tables.add("A4:H100", true, "ControlPagos").style = "TableStyleMedium2";
setWidths(pagos,{A:95,B:95,C:135,D:105,E:105,F:105,G:100,H:110}); pagos.freezePanes.freezeRows(4);

// RESUMEN MENSUAL
title(resumen, resumen.getRange("A1:K1"), "Evolución mensual de mis finanzas");
resumen.getRange("A3:K3").values = [["Mes", "Sueldo", "Tarjeta alimentos", "Pagos presup.", "Pagos reales", "Variables efectivo", "Variables alimentos", "Gasto total efectivo", "Saldo efectivo", "Meta ahorro", "Cumplimiento"]]; header(resumen.getRange("A3:K3"));
for(let i=0;i<12;i++) resumen.getRange(`A${4+i}`).values = [[new Date(2026,7+i,1)]];
resumen.getRange("A4:A15").format.numberFormat = monthFmt;
resumen.getRange("B4").formulas = [["='Configuración'!$B$5"]]; resumen.getRange("B4:B15").fillDown();
resumen.getRange("C4").formulas = [["='Configuración'!$B$6"]]; resumen.getRange("C4:C15").fillDown();
resumen.getRange("D4").formulas = [["=SUMIFS('Pagos fijos'!$D$5:$D$100,'Pagos fijos'!$A$5:$A$100,A4)"]]; resumen.getRange("D4:D15").fillDown();
resumen.getRange("E4").formulas = [["=SUMIFS('Pagos fijos'!$E$5:$E$100,'Pagos fijos'!$A$5:$A$100,A4)"]]; resumen.getRange("E4:E15").fillDown();
resumen.getRange("F4").formulas = [["=SUMIFS('Registro diario'!$E$5:$E$504,'Registro diario'!$B$5:$B$504,A4,'Registro diario'!$F$5:$F$504,\"<>Tarjeta alimentos\")"]]; resumen.getRange("F4:F15").fillDown();
resumen.getRange("G4").formulas = [["=SUMIFS('Registro diario'!$E$5:$E$504,'Registro diario'!$B$5:$B$504,A4,'Registro diario'!$F$5:$F$504,\"Tarjeta alimentos\")"]]; resumen.getRange("G4:G15").fillDown();
resumen.getRange("H4").formulas = [["=E4+F4"]]; resumen.getRange("H4:H15").fillDown();
resumen.getRange("I4").formulas = [["=B4-H4"]]; resumen.getRange("I4:I15").fillDown();
resumen.getRange("J4").formulas = [["='Configuración'!$E$21"]]; resumen.getRange("J4:J15").fillDown();
resumen.getRange("K4").formulas = [["=IF(AND(E4=0,F4=0,G4=0),\"Sin datos\",IF(I4>=J4,\"Meta lograda\",IF(I4>=0,\"Saldo positivo\",\"Saldo negativo\")))"]]; resumen.getRange("K4:K15").fillDown();
resumen.getRange("B4:J15").format.numberFormat = money;
resumen.getRange("I4:I15").conditionalFormats.add("cellIs", {operator:"lessThan",formula:0,format:{fill:"#FCE8E6",font:{color:C.red,bold:true}}});
resumen.getRange("I4:I15").conditionalFormats.add("cellIs", {operator:"greaterThanOrEqual",formula:0,format:{fill:C.greenLight,font:{color:"#356346",bold:true}}});
resumen.getRange("K4:K15").conditionalFormats.add("containsText", {text:"Meta lograda",format:{fill:C.greenLight,font:{color:"#356346",bold:true}}});
resumen.getRange("K4:K15").conditionalFormats.add("containsText", {text:"negativo",format:{fill:"#FCE8E6",font:{color:C.red,bold:true}}});
resumen.tables.add("A3:K15", true, "ResumenMensual").style = "TableStyleMedium2";
resumen.getRange("A18:C18").values = [["Lectura rápida", "Desde agosto", "Desde octubre"]]; header(resumen.getRange("A18:C18"));
resumen.getRange("A19:C21").values = [["Pagos presupuestados",null,null],["Variable sugerido",null,null],["Margen antes de ahorro",null,null]];
resumen.getRange("B19").formulas = [["=D4"]]; resumen.getRange("C19").formulas = [["=D6"]];
resumen.getRange("B20:C20").formulas = [["='Configuración'!$B$27","='Configuración'!$B$27"]];
resumen.getRange("B21").formulas = [["='Configuración'!$B$5-B19-B20"]]; resumen.getRange("C21").formulas = [["='Configuración'!$B$5-C19-C20"]];
resumen.getRange("B19:C21").format.numberFormat = money; card(resumen.getRange("A19:C21"),C.cream);
setWidths(resumen,{A:95,B:95,C:110,D:100,E:100,F:120,G:120,H:125,I:105,J:100,K:120}); resumen.freezePanes.freezeRows(3);

// INICIO / DASHBOARD
title(inicio, inicio.getRange("A1:L1"), "Mis finanzas, con calma y propósito");
inicio.getRange("A2:L2").merge(); inicio.getRange("A2").values = [["Registra tus gastos diarios, actualiza tus pagos reales y revisa aquí cómo avanza tu mes."]];
inicio.getRange("A2:L2").format = { fill:C.cream,font:{color:C.gray,italic:true},horizontalAlignment:"left" };
inicio.getRange("A4:B4").values = [["Mes revisado", new Date(2026,7,1)]];
inicio.getRange("B4").format = { fill:C.input,font:{color:"#0000FF",bold:true},numberFormat:monthFmt }; section(inicio.getRange("A4:B4"));

const cards=[
  ["A6:C6","A7:C8","Ingresos en efectivo","='Configuración'!$B$5",C.blush],
  ["D6:F6","D7:F8","Gastos reales del mes","=SUMIFS('Pagos fijos'!$E$5:$E$100,'Pagos fijos'!$A$5:$A$100,$B$4)+SUMIFS('Registro diario'!$E$5:$E$504,'Registro diario'!$B$5:$B$504,$B$4,'Registro diario'!$F$5:$F$504,\"<>Tarjeta alimentos\")",C.beige],
  ["G6:I6","G7:I8","Saldo disponible","=A7-D7",C.greenLight],
  ["J6:L6","J7:L8","Meta de ahorro","='Configuración'!$E$21",C.blush]
];
for(const [rh,rv,label,formula,fill] of cards){
  inicio.getRange(rh).merge(); inicio.getRange(rh.split(":")[0]).values=[[label]]; inicio.getRange(rh).format={fill,font:{bold:true,color:C.cocoa},horizontalAlignment:"center"};
  inicio.getRange(rv).merge(); inicio.getRange(rv.split(":")[0]).formulas=[[formula]]; inicio.getRange(rv).format={fill,font:{bold:true,color:C.roseDark,size:18},horizontalAlignment:"center",verticalAlignment:"center",numberFormat:money,borders:{preset:"outside",style:"thin",color:C.line}};
}
inicio.getRange("G7:I8").conditionalFormats.add("cellIs",{operator:"lessThan",formula:0,format:{fill:"#FCE8E6",font:{color:C.red,bold:true}}});

inicio.getRange("A10:F10").merge(); inicio.getRange("A10").values=[["¿En qué se está yendo mi dinero?"]]; section(inicio.getRange("A10:F10"));
inicio.getRange("A11:C11").values=[["Categoría","Presupuesto","Gasto real"]]; header(inicio.getRange("A11:C11"));
for(let i=0;i<6;i++){
  const r=12+i, cr=21+i;
  inicio.getRange(`A${r}`).formulas=[[`='Configuración'!A${cr}`]];
  inicio.getRange(`B${r}`).formulas=[[`='Configuración'!B${cr}`]];
  inicio.getRange(`C${r}`).formulas=[[`=SUMIFS('Registro diario'!$E$5:$E$504,'Registro diario'!$B$5:$B$504,$B$4,'Registro diario'!$C$5:$C$504,A${r})`]];
}
inicio.getRange("B12:C17").format.numberFormat=money;
inicio.getRange("B12:C17").conditionalFormats.add("dataBar",{color:C.rose,gradient:true});
inicio.getRange("E11:F11").values=[["Indicador","Resultado"]]; header(inicio.getRange("E11:F11"));
inicio.getRange("E12:E16").values=[["Pagos pendientes"],["Uso tarjeta alimentos"],["Saldo tarjeta alimentos"],["Saldo según registros"],["Estado del mes"]];
inicio.getRange("F12").formulas=[["=COUNTIFS('Pagos fijos'!$A$5:$A$100,$B$4,'Pagos fijos'!$G$5:$G$100,\"Pendiente\")"]];
inicio.getRange("F13").formulas=[["=SUMIFS('Registro diario'!$E$5:$E$504,'Registro diario'!$B$5:$B$504,$B$4,'Registro diario'!$F$5:$F$504,\"Tarjeta alimentos\")"]];
inicio.getRange("F14").formulas=[["='Configuración'!$B$6-F13"]];
inicio.getRange("F15").formulas=[["=G7"]];
inicio.getRange("F16").formulas=[["=IF(F12>0,\"Faltan pagos por registrar\",IF(G7<0,\"Necesita ajuste\",IF(G7<J7,\"Vas en positivo\",\"Meta lograda\")))"]];
inicio.getRange("F13:F15").format.numberFormat=money; card(inicio.getRange("E12:F16"),C.cream);
inicio.getRange("F16").conditionalFormats.add("containsText",{text:"Meta lograda",format:{fill:C.greenLight,font:{color:"#356346",bold:true}}});
inicio.getRange("F16").conditionalFormats.add("containsText",{text:"Necesita",format:{fill:"#FCE8E6",font:{color:C.red,bold:true}}});

inicio.getRange("A20:L20").merge(); inicio.getRange("A20").values=[["Consejo: registra incluso los gastos pequeños. La constancia dará una imagen real y permitirá ajustar sin adivinar."]];
inicio.getRange("A20:L20").format={fill:C.greenLight,font:{color:C.cocoa,italic:true},wrapText:true}; inicio.getRange("A20:L20").format.rowHeight=38;

// Chart helpers: category and 12-month trend
inicio.getRange("N1:O1").values=[["Categoría","Gasto real"]];
for(let i=0;i<6;i++) inicio.getRange(`N${2+i}:O${2+i}`).formulas=[[`=A${12+i}`,`=C${12+i}`]];
inicio.getRange("Q1:S1").values=[["Mes","Saldo","Meta"]];
for(let i=0;i<12;i++) inicio.getRange(`Q${2+i}:S${2+i}`).formulas=[[`=TEXT('Resumen mensual'!A${4+i},\"mmm yyyy\")`,`=IF('Resumen mensual'!K${4+i}=\"Sin datos\",\"\",'Resumen mensual'!I${4+i})`,`='Resumen mensual'!J${4+i}`]];
inicio.getRange("Q2:Q13").format.numberFormat=monthFmt; inicio.getRange("R2:S13").format.numberFormat=money;
const pie=inicio.charts.add("doughnut",inicio.getRange("N1:O7")); pie.title="Gastos variables del mes"; pie.hasLegend=true; pie.setPosition("H10","L19");
const line=inicio.charts.add("line",inicio.getRange("Q1:S13")); line.title="Saldo mensual vs. meta de ahorro"; line.hasLegend=true; line.yAxis={numberFormatCode:'"S/ "#,##0'}; line.xAxis={axisType:"textAxis"}; line.setPosition("A22","L38");
inicio.getRange("N:S").format.columnWidthPx=80;
setWidths(inicio,{A:110,B:100,C:100,D:95,E:150,F:115,G:95,H:95,I:95,J:95,K:95,L:95}); inicio.freezePanes.freezeRows(4);

// Verify + render every sheet
await fs.mkdir(outDir,{recursive:true});
for(const [name,range] of [["Inicio","A1:L38"],["Configuración","A1:F27"],["Registro diario","A1:G22"],["Pagos fijos","A1:H28"],["Resumen mensual","A1:K21"]]){
  const blob=await wb.render({sheetName:name,range,scale:1,format:"png"});
  await fs.writeFile(`${outDir}/vista_${name.replaceAll(" ","_")}.png`,new Uint8Array(await blob.arrayBuffer()));
}
const check=await wb.inspect({kind:"table",range:"Inicio!A1:L20",include:"values,formulas",tableMaxRows:20,tableMaxCols:12,maxChars:5000});
console.log("CHECK\n"+check.ndjson);
const errors=await wb.inspect({kind:"match",searchTerm:"#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",options:{useRegex:true,maxResults:100},summary:"final formula error scan",maxChars:3000});
console.log("ERRORS\n"+errors.ndjson);
const xlsx=await SpreadsheetFile.exportXlsx(wb);
await xlsx.save(`${outDir}/Control_de_Finanzas_Vanessa.xlsx`);
console.log(`SAVED ${outDir}/Control_de_Finanzas_Vanessa.xlsx`);
