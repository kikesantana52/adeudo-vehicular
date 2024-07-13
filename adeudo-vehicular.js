const { Builder, Browser, By, until } = require("selenium-webdriver");
const ExcelJS = require("exceljs");

async function main() {
  let driver = await new Builder().forBrowser(Browser.CHROME).build();
  const vehiculos = await getVehiculos();
  const multasRaw = await obtenerMultas(driver, vehiculos);
  console.log(multasRaw);
  const multasConFecha = await getMultasConFechas(driver, multasRaw);
  console.log(multasConFecha);

  crearReporte(multasConFecha);
  await driver.close();
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
  const data = [
    [
      "Vehiculo",
      "Placas",
      "Folio Multa",
      "Descripcion Multa",
      "Importe",
      "Tipo (P o R)",
    ],
  ];
  for (const vehiculo of vehiculos) {
    await driver.get(
      "https://gobiernoenlinea1.jalisco.gob.mx/serviciosVehiculares/adeudos"
    );

    await driver.findElement(By.id("placa")).sendKeys(vehiculo[3]);
    await driver
      .findElement(By.id("numeroSerie"))
      .sendKeys(vehiculo[4].substring(vehiculo[4].length - 5));
    await driver
      .findElement(By.id("nombrePropietario"))
      .sendKeys("CARLOS SANTANA RUELAS");
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

const crearReporte = async (data) => {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Multas");
  data.forEach((row, rowIndex) => {
    row.forEach((cell, colIndex) => {
      worksheet.getCell(rowIndex + 1, colIndex + 1).value = cell;
    });
  });
  await workbook.xlsx.writeFile(`./multas.xlsx`);
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
    ],
  ];
  const fechasPorId = {};
  for (const multa of multasRaw) {
    if (fechasPorId[multa[2].substring(4)]) {
      multa.push(fechasPorId[multa[2].substring(4)]);
      multas.push(multa);
    } else {
      try {
        await driver.get(
          `https://gobiernoenlinea1.jalisco.gob.mx/serviciosVehiculares/visorInfraccion/FE/${multa[2].substring(
            4
          )}`
        );
        const fecha = await driver
          .findElement(
            By.xpath("/html/body/div/div/form/div/div[2]/div[2]/div")
          )
          .getText();
        let fechaValue = fecha.substring(7).trim();

        if (fechaValue === "") {
          await driver.get(
            `https://gobiernoenlinea1.jalisco.gob.mx/serviciosVehiculares/visorInfraccion/SIGA/${multa[2].substring(
              4
            )}`
          );
          const fecha = await driver
            .findElement(
              By.xpath("/html/body/div/div/form/div/div[2]/div[2]/div")
            )
            .getText();
          fechaValue = fecha.substring(7).trim();
        }
        fechasPorId[multa[2].substring(4)] = fechaValue;
        multa.push(fechaValue);
      } catch (error) {
        console.log(multa[2]);
      }
      multas.push(multa);
    }
  }
  return multas;
}

main();
