const ExcelJS = require("exceljs");
const fs = require("fs");
const csvFilePath = "input.csv"; // Nombre del archivo CSV
const outputFilePath = "output.csv"; // Archivo de salida

async function processCSV() {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Sheet1");

  // Leer archivo CSV manualmente línea por línea
  let data = fs
    .readFileSync(csvFilePath, "utf8")
    .split("\n")
    .map((row) => row.split(","));
  let fecha = null;

  for (let row of data) {
    const value = row[0].trim();
    row[1] = row[1]?.replaceAll("\r", "");
    if (/^\d{2}\.\d{2}\.\d{4}$/.test(value)) {
      fecha = value.replaceAll(".", "-");
    }
    row.push(fecha.trim());
  }

  // Guardar el resultado en un nuevo archivo CSV
  const csvOutput = data.map((row) => row.join(",")).join("\n");
  console.log(csvOutput);

  fs.writeFileSync(outputFilePath, csvOutput, "utf8");
  console.log("Proceso completado. Archivo guardado como", outputFilePath);
}

processCSV().catch(console.error);
