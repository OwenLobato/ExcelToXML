const XLSX = require('xlsx');
const xmlbuilder = require('xmlbuilder');
const fs = require('fs');

const inputFile = 'inputNuevo.xlsx'; // Archivo de Excel (Entrada)
const outputFile = 'outputNuevo.xml'; // Archivo XML (Salida)

const sheetName = 'Hoja1'; // Hoja del excel

const columnHeaders = { // Ubicacion de las columnas
  COL_1: 'B',
  COL_2: 'D',
  COL_3: 'E',
  COL_4: 'F'
};

function getCellName(column, row) {
  return `${column}${row}`;
}

function getCellValue(sheet, column, row) {
  const cellName = getCellName(column, row);
  const cell = sheet[cellName];
  return cell ? cell.v : '';
}

function capitalizeString(str) {
  return str.toLowerCase().replace(/(?:^|\s)\S/g, function (a) { return a.toUpperCase(); });
}

// Cargar el archivo Excel
const workbook = XLSX.readFile(inputFile);
const sheet = workbook.Sheets[sheetName];

// Encontrar el rango de filas
const range = XLSX.utils.decode_range(sheet['!ref']);
let startRow = range.s.r;
let endRow = range.e.r;

// Buscar el inicio de la tabla
for (let i = range.s.r; i <= range.e.r; i++) {
  const codComuna = getCellValue(sheet, columnHeaders.COD_COMUNA, i);
  if (codComuna !== '') {
    startRow = i;
    break;
  }
}

// Buscar el final de la tabla
for (let i = range.e.r; i >= range.s.r; i--) {
  const codComuna = getCellValue(sheet, columnHeaders.COD_COMUNA, i);
  if (codComuna !== '') {
    endRow = i;
    break;
  }
}

// Crear objeto XML
const root = xmlbuilder.create('root');

// Iterar sobre las filas del archivo Excel
for (let i = startRow; i <= endRow; i++) {
  const customObject = root.ele('custom-object', { 'type-id': 'objectId', 'object-id': getCellValue(sheet, columnHeaders.COD_COMUNA, i) });

  customObject.ele('object-attribute', { 'attribute-id': 'columnaA' }, getCellValue(sheet, columnHeaders.COL_1, i));
  customObject.ele('object-attribute', { 'attribute-id': 'columnaB' }, getCellValue(sheet, columnHeaders.COL_2, i));
  customObject.ele('object-attribute', { 'attribute-id': 'columnaC' }, capitalizeString(getCellValue(sheet, columnHeaders.COL_3, i)));
  customObject.ele('object-attribute', { 'attribute-id': 'columnaD' }, getCellValue(sheet, columnHeaders.COL_4, i));
}

let xml = root.end({ pretty: true }); // Genera XML
fs.writeFileSync(outputFile, xml); // Escribir el XML en un archivo

console.log(`Transformación completada, XML guardado en ${outputFile}`);
