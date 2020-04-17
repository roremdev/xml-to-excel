const xml2js = require("xml2js");
const fs = require("fs");
const parser = new xml2js.Parser({ attrkey: "ATTR" });
const ExcelJS = require("exceljs");

const { config } = require("./configEnv");


let stackData = [];

const mounths = [
  ["Q1", "Q2"],
  ["Q3", "Q4"],
  ["Q5", "Q6"],
  ["Q7", "Q8"],
  ["Q9", "Q10"],
  ["Q11", "Q12"],
  ["Q13", "Q14"],
  ["Q15", "Q16"],
  ["Q17", "Q18"],
  ["Q19", "Q20"],
  ["Q21", "Q22"],
  ["Q23", "Q24"]
];

function getData(dataXML, string) {
  let value = "";
  switch(string) {
    case "rfc":
      try {
        value = dataXML["cfdi:Comprobante"]["cfdi:Receptor"][0].ATTR.Rfc;
      } catch (error) {
        value = "";
      }
      return value;
    case "nombre":
      try {
        value = dataXML["cfdi:Comprobante"]["cfdi:Receptor"][0].ATTR.Nombre;
      } catch (error) {
        value = "";
      }
      return value;
    case "fechaPago":
      try {
        value = dataXML["cfdi:Comprobante"]["cfdi:Complemento"][0]["nomina12:Nomina"][0].ATTR.FechaPago;
      } catch (error) {
        value = "";
      }
      return value;
    case "fechaInicialPago":
      try {
        value = dataXML["cfdi:Comprobante"]["cfdi:Complemento"][0]["nomina12:Nomina"][0].ATTR.FechaInicialPago;
      } catch (error) {
        value = "";
      }
      return value;
    case "fechaFinalPago":
      try {
        value = dataXML["cfdi:Comprobante"]["cfdi:Complemento"][0]["nomina12:Nomina"][0].ATTR.FechaFinalPago;
      } catch (error) {
        value = "";
      }
      return value;
    case "quincena":
      try {
        value = dataXML["cfdi:Comprobante"]["cfdi:Complemento"][0]["nomina12:Nomina"][0].ATTR.FechaInicialPago;
        value = fortnightlyCalculation(value);
      } catch (error) {
        value = "";
      }
      return value;
    case "periodoPago":
      try {
        value = dataXML["cfdi:Comprobante"]["cfdi:Addenda"][0].RFC[0].ATTR.PeriodoPago;
      } catch (error) {
        value = "";
      }
      return value;
    case "registroPatronal":
      try {
        value = dataXML["cfdi:Comprobante"]["cfdi:Complemento"][0]["nomina12:Nomina"][0]["nomina12:Emisor"][0].ATTR.RegistroPatronal;
      } catch (error) {
        value = "";
      }
      return value;
    case "origenRecurso":
      try {
        value = dataXML["cfdi:Comprobante"]["cfdi:Complemento"][0]["nomina12:Nomina"][0]["nomina12:Emisor"][0]["nomina12:EntidadSNCF"][0].ATTR.OrigenRecurso;
      } catch (error) {
        value = "";
      }
      return value;
    case "totalPercepciones":
      try {
        value = dataXML["cfdi:Comprobante"]["cfdi:Complemento"][0]["nomina12:Nomina"][0].ATTR.TotalPercepciones;
      } catch (error) {
        value = "";
      }
      return value;
    case "totalDeducciones":
      try {
        value = dataXML["cfdi:Comprobante"]["cfdi:Complemento"][0]["nomina12:Nomina"][0].ATTR.TotalDeducciones;
      } catch (error) {
        value = "";
      }
      return value;
    case "clave":
      try {
        value = dataXML["cfdi:Comprobante"]["cfdi:Complemento"][0]["nomina12:Nomina"][0]["nomina12:Deducciones"][0]["nomina12:Deduccion"][0].ATTR.Clave;
      } catch (error) {
        value = "";
      }
      return value;
    case "concepto":
      try {
        value = dataXML["cfdi:Comprobante"]["cfdi:Complemento"][0]["nomina12:Nomina"][0]["nomina12:Deducciones"][0]["nomina12:Deduccion"][0].ATTR.Concepto;
      } catch (error) {
        value = "";
      }
      return value;
    case "importe":
      try {
        value = dataXML["cfdi:Comprobante"]["cfdi:Complemento"][0]["nomina12:Nomina"][0]["nomina12:Deducciones"][0]["nomina12:Deduccion"][0].ATTR.Importe;
      } catch (error) {
        value = "";
      }
      return value;
    case "totalOtrasDeducciones":
      try {
        value = dataXML["cfdi:Comprobante"]["cfdi:Complemento"][0]["nomina12:Nomina"][0]["nomina12:Deducciones"][0].ATTR.TotalOtrasDeducciones;
      } catch (error) {
        value = "";
      }
      return value;
    case "totalImpuestosRetenidos":
      try {
        value = dataXML["cfdi:Comprobante"]["cfdi:Complemento"][0]["nomina12:Nomina"][0]["nomina12:Deducciones"][0].ATTR.TotalImpuestosRetenidos;
      } catch (error) {
        value = "";
      }
      return value;
    case "total":
      try {
        value = dataXML["cfdi:Comprobante"].ATTR.Total;
      } catch (error) {
        value = "";
      }
      return value;
    case "uuid":
      try {
        value = dataXML["cfdi:Comprobante"]["cfdi:Complemento"][0]["tfd:TimbreFiscalDigital"][0].ATTR.UUID;
      } catch (error) {
        value = "";
      }
      return value;
    case "fechaTimbrado":
      try {
        value = dataXML["cfdi:Comprobante"]["cfdi:Complemento"][0]["tfd:TimbreFiscalDigital"][0].ATTR.FechaTimbrado;
      } catch (error) {
        value = "";
      }
      return value;
    case "cedulaProf":
      try {
        value = dataXML["cfdi:Comprobante"]["cfdi:Addenda"][0]["GEP:AddendaEmisor"][0]["GEP:InformacionPago"][0].ATTR.CedulaProf;
      } catch (error) {
        value = "";
      }
      return value;
    case "serie":
      try {
        value = dataXML["cfdi:Comprobante"].ATTR.Serie;
      } catch (error) {
        value = "";
      }
      return value;
    case "folio":
      try {
        value = dataXML["cfdi:Comprobante"].ATTR.Folio;
      } catch (error) {
        value = "";
      }
      return value;
    case "totalOtrosPagos":
      try {
        value = dataXML["cfdi:Comprobante"]["cfdi:Complemento"][0]["nomina12:Nomina"][0].ATTR.TotalOtrosPagos;
      } catch (error) {
        value = "";
      }
      return value;
    case "numEmpleado":
      try {
        value = dataXML["cfdi:Comprobante"]["cfdi:Complemento"][0]["nomina12:Nomina"][0]["nomina12:Receptor"][0].ATTR.NumEmpleado;
      } catch (error) {
        value = "";
      }
      return value;
    case "curp":
      try {
        value = dataXML["cfdi:Comprobante"]["cfdi:Complemento"][0]["nomina12:Nomina"][0]["nomina12:Receptor"][0].ATTR.Curp;
      } catch (error) {
        value = "";
      }
      return value;
    case "numSeguridadSocial":
      try {
        value = dataXML["cfdi:Comprobante"]["cfdi:Complemento"][0]["nomina12:Nomina"][0]["nomina12:Receptor"][0].ATTR.NumSeguridadSocial;
      } catch (error) {
        value = "";
      }
      return value;
  }
}

function readData(dataXML) {

  const dataJSON = {
    rfc:    getData(dataXML, "rfc"),
    nombre: getData(dataXML, "nombre"),
    fechaPago: getData(dataXML, "fechaPago"),
    fechaInicialPago: getData(dataXML, "fechaInicialPago"),
    fechaFinalPago: getData(dataXML, "fechaFinalPago"),
    quincena: getData(dataXML, "quincena"),
    periodoPago: getData(dataXML, "periodoPago"),
    registroPatronal: getData(dataXML, "registroPatronal"),
    origenRecurso: getData(dataXML, "origenRecurso"),
    totalPercepciones: getData(dataXML, "totalPercepciones"),
    totalDeducciones: getData(dataXML, "totalDeducciones"),
    clave: getData(dataXML, "clave"),
    concepto: getData(dataXML, "concepto"),
    importe: getData(dataXML, "importe"),
    totalOtrasDeducciones: getData(dataXML, "totalOtrasDeducciones"),
    totalImpuestosRetenidos: getData(dataXML, "totalImpuestosRetenidos"),
    total: getData(dataXML, "total"),
    uuid: getData(dataXML, "uuid"),
    fechaTimbrado: getData(dataXML, "fechaTimbrado"),
    cedulaProf: getData(dataXML, "cedulaProf"),
    serie:  getData(dataXML, "serie"),
    folio:  getData(dataXML, "folio"),
    totalOtrosPagos: getData(dataXML, "totalOtrosPagos"),
    numEmpleado: getData(dataXML, "numEmpleado"),
    curp: getData(dataXML, "curp"),
    numSeguridadSocial: getData(dataXML, "numSeguridadSocial"),
    tipoNomina: config.typePayRoll,
    carpeta: config.directory
  };
  return dataJSON;
}

function generateExcel(data) {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Data");
  worksheet.columns = [
    { header: "RFC", key: "rfc", width: 20 },
    { header: "Nombre", key: "nombre", width: 20 },
    { header: "Fecha pago", key: "fechaPago", width: 20 },
    { header: "Fecha inicial pago", key: "fechaInicialPago", width: 20 },
    { header: "Fecha final pago", key: "fechaFinalPago", width: 20 },
    { header: "Quincena", key: "quincena", width: 20 },
    { header: "Periodo pago", key: "periodoPago", width: 20 },
    { header: "Registro patronal", key: "registroPatronal", width: 20 },
    { header: "Origen recurso", key: "origenRecurso", width: 20 },
    { header: "Total percepciones", key: "totalPercepciones", width: 20 },
    { header: "Total deducciones", key: "totalDeducciones", width: 20 },
    { header: "Clave", key: "clave", width: 20 },
    { header: "Concepto", key: "concepto", width: 20 },
    { header: "Importe", key: "importe", width: 20 },
    {
      header: "Total otras deducciones",
      key: "totalOtrasDeducciones",
      width: 20
    },
    {
      header: "Total impuestos retenidos",
      key: "totalImpuestosRetenidos",
      width: 20
    },
    { header: "Total", key: "total", width: 20 },
    { header: "UUID", key: "uuid", width: 20 },
    { header: "Fecha timbrado", key: "fechaTimbrado", width: 20 },
    { header: "Cedula profesional", key: "cedulaProf", width: 20 },
    { header: "Serie", key: "serie", width: 20 },
    { header: "Folio", key: "folio", width: 20 },
    { header: "Total otros pagos", key: "totalOtrosPagos", width: 20 },
    { header: "Num empleado", key: "numEmpleado", width: 20 },
    { header: "CURP", key: "curp", width: 20 },
    { header: "Num seguridad social", key: "numSeguridadSocial", width: 20 },
    { header: "Tipo de Nómina", key: "tipoNomina", width: 20 },
    { header: "Carpeta", key: "carpeta", width: 20 }
  ];
  data.forEach(rowData => worksheet.addRow(rowData));
  console.log(data);

  workbook.xlsx.writeFile(config.nameFileExcel);
}

function listDirectory() {
  let pathDirectory = config.pathDirectory;

  fs.readdir(pathDirectory, function(err, items) {
    for (let counterFiles = 0; counterFiles < items.length; counterFiles++) {
      console.log(`File: ${pathDirectory + items[counterFiles]}`);
      let xml_string = fs.readFileSync(
        pathDirectory + items[counterFiles],
        "utf8"
      );
      parser.parseString(xml_string, function(error, result) {
        if (error === null) {
          saveDataArray(readData(result));
        } else {
          console.log(error);
        }
      });
    }
    generateExcel(stackData);
  });
}

function saveDataArray(data) {
  stackData.push(data);
}

function fortnightlyCalculation(datePeriod) {
  const date = datePeriod.split("-");
  let numberMount = Number(date[1]);
  return Number(date[2]) <= 15
    ? mounths[numberMount - 1][0]
    : mounths[numberMount - 1][1];
  // console.log(numberBiweekly);
}

listDirectory();
