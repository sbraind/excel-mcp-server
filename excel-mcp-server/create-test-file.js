#!/usr/bin/env node

/**
 * Script to create a test Excel file with sample data
 */

import ExcelJS from 'exceljs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

async function createTestFile() {
  const workbook = new ExcelJS.Workbook();

  // Sheet 1: Sales Data
  const salesSheet = workbook.addWorksheet('Sales');

  // Add headers
  salesSheet.columns = [
    { header: 'Product', key: 'product', width: 20 },
    { header: 'Quantity', key: 'quantity', width: 12 },
    { header: 'Price', key: 'price', width: 12 },
    { header: 'Total', key: 'total', width: 15 },
    { header: 'Date', key: 'date', width: 15 },
  ];

  // Add sample data
  const products = ['Laptop', 'Mouse', 'Keyboard', 'Monitor', 'Headphones', 'Webcam', 'Printer', 'Scanner'];
  const dates = ['2024-01-15', '2024-01-16', '2024-01-17', '2024-01-18', '2024-01-19'];

  for (let i = 0; i < 50; i++) {
    const product = products[Math.floor(Math.random() * products.length)];
    const quantity = Math.floor(Math.random() * 10) + 1;
    const price = Math.floor(Math.random() * 500) + 50;
    const date = dates[Math.floor(Math.random() * dates.length)];

    const row = salesSheet.addRow({
      product,
      quantity,
      price,
      total: { formula: `B${i + 2}*C${i + 2}` },
      date,
    });

    // Format price and total as currency
    row.getCell(3).numFmt = '$#,##0.00';
    row.getCell(4).numFmt = '$#,##0.00';
  }

  // Format header row
  salesSheet.getRow(1).font = { bold: true, size: 12 };
  salesSheet.getRow(1).fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF4472C4' },
  };
  salesSheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };

  // Add totals row
  const lastRow = salesSheet.rowCount + 1;
  const totalsRow = salesSheet.getRow(lastRow);
  totalsRow.getCell(1).value = 'TOTALS';
  totalsRow.getCell(1).font = { bold: true };
  totalsRow.getCell(2).value = { formula: `SUM(B2:B${lastRow - 1})` };
  totalsRow.getCell(4).value = { formula: `SUM(D2:D${lastRow - 1})` };
  totalsRow.getCell(4).numFmt = '$#,##0.00';
  totalsRow.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFE7E6E6' },
  };

  // Sheet 2: Employee Data
  const employeeSheet = workbook.addWorksheet('Employees');

  employeeSheet.columns = [
    { header: 'Employee ID', key: 'id', width: 15 },
    { header: 'Name', key: 'name', width: 20 },
    { header: 'Department', key: 'department', width: 20 },
    { header: 'Salary', key: 'salary', width: 15 },
    { header: 'Start Date', key: 'startDate', width: 15 },
  ];

  const names = ['John Smith', 'Jane Doe', 'Bob Johnson', 'Alice Williams', 'Charlie Brown', 'Diana Prince'];
  const departments = ['Sales', 'Marketing', 'IT', 'HR', 'Finance'];

  for (let i = 0; i < 20; i++) {
    const row = employeeSheet.addRow({
      id: `EMP${String(i + 1).padStart(3, '0')}`,
      name: names[Math.floor(Math.random() * names.length)],
      department: departments[Math.floor(Math.random() * departments.length)],
      salary: Math.floor(Math.random() * 80000) + 40000,
      startDate: new Date(2020 + Math.floor(Math.random() * 4), Math.floor(Math.random() * 12), 1),
    });

    row.getCell(4).numFmt = '$#,##0';
    row.getCell(5).numFmt = 'mm/dd/yyyy';
  }

  // Format header
  employeeSheet.getRow(1).font = { bold: true, size: 12 };
  employeeSheet.getRow(1).fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF70AD47' },
  };
  employeeSheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };

  // Sheet 3: Inventory
  const inventorySheet = workbook.addWorksheet('Inventory');

  inventorySheet.columns = [
    { header: 'SKU', key: 'sku', width: 15 },
    { header: 'Product Name', key: 'name', width: 25 },
    { header: 'Stock', key: 'stock', width: 12 },
    { header: 'Reorder Level', key: 'reorder', width: 15 },
    { header: 'Status', key: 'status', width: 15 },
  ];

  const productNames = [
    'Widget A',
    'Gadget B',
    'Device C',
    'Tool D',
    'Component E',
    'Part F',
    'Assembly G',
    'Module H',
  ];

  for (let i = 0; i < 30; i++) {
    const stock = Math.floor(Math.random() * 200);
    const reorder = Math.floor(Math.random() * 50) + 20;
    const status = stock <= reorder ? 'LOW STOCK' : 'OK';

    const row = inventorySheet.addRow({
      sku: `SKU${String(i + 1).padStart(4, '0')}`,
      name: productNames[Math.floor(Math.random() * productNames.length)],
      stock,
      reorder,
      status,
    });

    // Color code status
    if (status === 'LOW STOCK') {
      row.getCell(5).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFF0000' },
      };
      row.getCell(5).font = { color: { argb: 'FFFFFFFF' }, bold: true };
    } else {
      row.getCell(5).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF92D050' },
      };
    }
  }

  // Format header
  inventorySheet.getRow(1).font = { bold: true, size: 12 };
  inventorySheet.getRow(1).fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFFFC000' },
  };
  inventorySheet.getRow(1).font = { bold: true, color: { argb: 'FF000000' } };

  // Save file
  const filePath = path.join(__dirname, 'test.xlsx');
  await workbook.xlsx.writeFile(filePath);

  console.log(`✅ Test file created successfully: ${filePath}`);
  console.log('\nThe file contains:');
  console.log('  📊 Sheet 1 (Sales): 50 rows of sales data with formulas');
  console.log('  👥 Sheet 2 (Employees): 20 employee records');
  console.log('  📦 Sheet 3 (Inventory): 30 inventory items with status indicators');
}

createTestFile().catch((error) => {
  console.error('❌ Error creating test file:', error);
  process.exit(1);
});
