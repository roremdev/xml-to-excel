require("dotenv").config();

const config = {
  nameFileExcel: process.env.NAME_FILE_EXCEL,
  directory: process.env.DIRECTORY,
  pathDirectory: process.env.PATH_DIRECTORY,
  typePayRoll: process.env.TYPE_PAYROLL
};

module.exports = { config };