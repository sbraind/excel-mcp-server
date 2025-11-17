const ExcelJS = require('exceljs');

async function createTestFile() {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Ventas 2024');
  
  // Headers
  sheet.columns = [
    { header: 'Mes', key: 'mes', width: 15 },
    { header: 'Producto', key: 'producto', width: 20 },
    { header: 'Cantidad', key: 'cantidad', width: 12 },
    { header: 'Precio', key: 'precio', width: 12 },
    { header: 'Total', key: 'total', width: 15 }
  ];
  
  // Datos de ejemplo
  const data = [
    { mes: 'Enero', producto: 'Laptop HP', cantidad: 5, precio: 800, total: 4000 },
    { mes: 'Enero', producto: 'Mouse Logitech', cantidad: 20, precio: 25, total: 500 },
    { mes: 'Febrero', producto: 'Laptop HP', cantidad: 3, precio: 800, total: 2400 },
    { mes: 'Febrero', producto: 'Teclado Mecánico', cantidad: 15, precio: 120, total: 1800 },
    { mes: 'Marzo', producto: 'Monitor Samsung', cantidad: 8, precio: 300, total: 2400 }
  ];
  
  sheet.addRows(data);
  
  // Agregar fórmula en la última fila
  const lastRow = sheet.lastRow.number + 2;
  sheet.getCell(`A${lastRow}`).value = 'TOTAL:';
  sheet.getCell(`A${lastRow}`).font = { bold: true };
  sheet.getCell(`E${lastRow}`).value = { formula: `SUM(E2:E${lastRow-2})` };
  sheet.getCell(`E${lastRow}`).font = { bold: true };
  
  // Estilo a los headers
  sheet.getRow(1).font = { bold: true };
  sheet.getRow(1).fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF4472C4' }
  };
  
  await workbook.xlsx.writeFile('/Users/sebastianbrain/Desktop/Experimentos/Experimentos-1/test-ventas.xlsx');
  console.log('✅ Archivo creado: test-ventas.xlsx');
}

createTestFile().catch(console.error);
