const ExcelJS = require("exceljs");

const cargarManifiestos = async () => {
  const workbook = new ExcelJS.Workbook();
  const cuestionario = await workbook.xlsx.readFile("cuestionario.xlsx");
  const worksheet = cuestionario.getWorksheet(1);
  const data = [];
  let currentFolio = "";
  let currentManifiesto = {};
  let currentRecoleccion = {};

  worksheet.eachRow((row, rowNumber) => {
    if (row.getCell(1).value !== currentFolio) {
      data.push(currentManifiesto);
      currentFolio = row.getCell(1).value;
      currentManifiesto = {
        folio: row.getCell(1).value,
        codigoCliente: row.getCell(2).value,
        nombreCliente: row.getCell(3).value,
        domicilio: row.getCell(4).value,
        codigoVehiculo: row.getCell(5).value,
        chofer: row.getCell(6).value,
        fechaHora: row.getCell(7).value,
        ruta: row.getCell(8).value,
        telefono: null,
        emailEnviado: false,
        recoleccions: [],
        fotos: [],
      };
    }
    switch (row.getCell(10).value) {
      case "Cantidad":
        currentRecoleccion.cantidad = row.getCell(11).value;
        break;
      case "Elemento":
        currentRecoleccion.codigo_producto = row
          .getCell(11)
          .value.substring(2, row.getCell(11).value.length - 2);
        break;
      case "Descripcion":
        currentRecoleccion.size = row
          .getCell(11)
          .value.substring(2, row.getCell(11).value.length - 2);
        currentManifiesto.recoleccions.push(currentRecoleccion);
        currentRecoleccion = {};
        break;
      case "Nombre, Firma y fotos":
        if (row.getCell(13).value !== null) {
          currentManifiesto.fotos.push(row.getCell(13).value);
        }
        break;
      case "Fotografias tomadas":
        if (row.getCell(13).value !== null) {
          const fotos = row.getCell(13).value.split("\\n\\n");
          currentManifiesto.fotos = [...currentManifiesto.fotos, ...fotos];
        }
        break;
    }
  });
  for (let manifiesto of data) {
    await subirManifiesto(manifiesto);
  }
};

const subirManifiesto = async (body) => {
  const data = await fetch("http://localhost:9001/serviceOrders", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify(body),
  });

  const response = await data.json();
  console.log(response);

  console.log(body.folio);

  return response;
};

cargarManifiestos();
