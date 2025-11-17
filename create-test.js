import ExcelJS from 'exceljs';

const workbook = new ExcelJS.Workbook();
const sheet = workbook.addWorksheet('Ventas');

// Agregar encabezados
sheet.addRow(['Producto', 'Cantidad', 'Precio', 'Total']);

// Agregar datos
sheet.addRow(['Laptop', 5, 1200, 6000]);
sheet.addRow(['Mouse', 20, 25, 500]);
sheet.addRow(['Teclado', 15, 75, 1125]);
sheet.addRow(['Monitor', 8, 350, 2800]);

// Aplicar formato a los encabezados
sheet.getRow(1).font = { bold: true };
sheet.getRow(1).fill = {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: 'FF4472C4' }
};
sheet.getRow(1).font = { color: { argb: 'FFFFFFFF' }, bold: true };

// Ajustar anchos de columna
sheet.getColumn(1).width = 15;
sheet.getColumn(2).width = 12;
sheet.getColumn(3).width = 12;
sheet.getColumn(4).width = 12;

// Guardar el archivo
const filePath = '/Users/sebastianbrain/Desktop/Experimentos/Experimentos-1/test-ventas.xlsx';
await workbook.xlsx.writeFile(filePath);

console.log(`✅ Archivo creado: ${filePath}`);
