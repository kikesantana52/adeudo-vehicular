const { Builder, Browser, By, until } = require("selenium-webdriver");
const ExcelJS = require("exceljs");

async function main() {
  let driver = await new Builder().forBrowser(Browser.CHROME).build();
  console.log("************INICIANDO PROGRAMA*************");
  const vehiculos = await getVehiculos();
  console.log(
    `************${vehiculos.length} vehiculos encontrados************`
  );
  const multasRaw = await obtenerMultas(driver, vehiculos);
  console.log(`************${multasRaw.length} multas encontradas************`);
  console.log(
    `************Obteniendo fechas y motivos de las multas************`
  );
  const multasConFecha = await getMultasConFechas(driver, multasRaw);
  console.log(`************Creando reporte************`);
  crearReporte(multasConFecha);
  await driver.close();
  console.log(
    `************Listo. Fin del programa, ahora puedes revisar el archivo_maestro.xlsx************`
  );
}

const getVehiculos = async () => {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile("./vehiculos_info.xlsx");
  const worksheet = workbook.getWorksheet(1);
  const data = [];
  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber !== 1) {
      const rowData = [];
      row.eachCell((cell, colNumber) => {
        rowData.push(cell.value);
      });
      data.push(rowData);
    }
  });
  return data;
};

async function obtenerMultas(driver, vehiculos) {
  const data = [];
  for (const vehiculo of vehiculos) {
    await driver.get(
      "https://gobiernoenlinea1.jalisco.gob.mx/serviciosVehiculares/adeudos"
    );

    await driver.findElement(By.id("placa")).sendKeys(vehiculo[3]);
    await driver
      .findElement(By.id("numeroSerie"))
      .sendKeys(vehiculo[4].substring(vehiculo[4].length - 5));
    await driver.findElement(By.id("nombrePropietario")).sendKeys(vehiculo[6]);
    await driver.findElement(By.id("numeroMotor")).sendKeys(vehiculo[5]);
    await driver.findElement(By.id("btnConsultar")).click();
    try {
      await driver.wait(until.elementLocated(By.id("swal2-title")), 2000);
      continue;
    } catch (error) {}

    try {
      await driver.wait(until.elementLocated(By.id("adeudosList")), 8000);
    } catch (error) {
      continue;
    }
    const vehiculoRawData = await driver
      .findElement(By.id("adeudosList"))
      .getAttribute("value");
    const multas = JSON.parse(vehiculoRawData)[0].conceptos;
    let folioAnterior = "";
    multas.forEach((multa) => {
      if (multa.folio) {
        folioAnterior = multa.folio;
      }
      data.push([
        vehiculo[0],
        vehiculo[3],
        folioAnterior,
        multa.descripcion,
        multa.importe,
        multa.tipo,
      ]);
    });
  }
  return data;
}

const limpiarHoja = async (index) => {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile("./archivo_maestro.xlsx");
  const worksheet = workbook.getWorksheet(index);
  if (!worksheet) {
    console.log(`La hoja número ${index} no existe`);
    return [];
  }
  const rowCount = worksheet.rowCount;
  const oldData = [];
  for (let i = 2; i <= rowCount; i++) {
    const oldValues = worksheet.getRow(i).values;
    oldValues.shift();
    oldData.push(oldValues);
    worksheet.getRow(i).values = [];
  }
  await workbook.xlsx.writeFile("./archivo_maestro.xlsx");
  return oldData;
};

const crearReporte = async (data) => {
  const multasViejas = await limpiarHoja(1);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile("./archivo_maestro.xlsx");
  const multasNuevasSheet = workbook.getWorksheet(1);
  data.forEach((row, rowIndex) => {
    row.forEach((cell, colIndex) => {
      multasNuevasSheet.getCell(rowIndex + 1, colIndex + 1).value = cell;
    });
  });
  const multasViejasSheet = workbook.getWorksheet(3);
  multasViejas.forEach((row, rowIndex) => {
    row.forEach((cell, colIndex) => {
      multasViejasSheet.getCell(rowIndex + 1, colIndex + 1).value = cell;
    });
  });
  await workbook.xlsx.writeFile(`./archivo_maestro.xlsx`);
};

async function getMultasConFechas(driver, multasRaw) {
  const multas = [
    [
      "Vehiculo",
      "Placas",
      "Folio Multa",
      "Descripcion Multa",
      "Importe",
      "Tipo (P o R)",
      "Fecha",
      "Link",
    ],
  ];
  const fechasPorId = {};
  const VISORES_DE_INFRACCIONES = [
    "https://gobiernoenlinea1.jalisco.gob.mx/serviciosVehiculares/visorInfraccion/SC/",
    "https://gobiernoenlinea1.jalisco.gob.mx/serviciosVehiculares/visorInfraccion/SIGA/",
    "https://gobiernoenlinea1.jalisco.gob.mx/serviciosVehiculares/visorInfraccion/FE/",
  ];
  for (const multa of multasRaw) {
    if (fechasPorId[multa[2].substring(4)]) {
      multa.push(fechasPorId[multa[2].substring(4)]);
      multas.push(multa);
    } else {
      try {
        let fechaValue = "";
        let link = "";
        for (let i = 0; i < VISORES_DE_INFRACCIONES.length; i++) {
          await driver.get(
            `${VISORES_DE_INFRACCIONES[i]}${multa[2].substring(4)}`
          );
          const fecha = await driver
            .findElement(
              By.xpath("/html/body/div/div/form/div/div[2]/div[2]/div")
            )
            .getText();
          fechaValue = fecha.substring(7).trim();
          if (fechaValue !== "") {
            link = `${VISORES_DE_INFRACCIONES[i]}${multa[2].substring(4)}`;
            break;
          }
        }
        fechasPorId[multa[2].substring(4)] = fechaValue;
        multa.push(fechaValue);
        multa.push(link);
      } catch (error) {}
      multas.push(multa);
    }
  }
  return multas;
}

main();
