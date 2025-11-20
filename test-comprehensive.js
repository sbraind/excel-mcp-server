#!/usr/bin/env node

/**
 * Comprehensive test of all live editing features
 */

import {
  isExcelRunning,
  isFileOpenInExcel,
  updateCellViaAppleScript,
  addRowViaAppleScript,
  writeRangeViaAppleScript,
  setFormulaViaAppleScript,
  formatCellViaAppleScript,
  setColumnWidthViaAppleScript,
  setRowHeightViaAppleScript,
  mergeCellsViaAppleScript,
  saveFileViaAppleScript
} from './dist/tools/excel-applescript.js';
import { resolve } from 'path';

async function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

async function comprehensiveTest() {
  console.log('\n🧪 COMPREHENSIVE LIVE EDITING TEST\n');
  console.log('━'.repeat(60));

  const filePath = resolve('./test.xlsx');
  const sheetName = 'Sales';

  try {
    // Verification
    console.log('\n📋 Pre-flight checks:');
    const excelRunning = await isExcelRunning();
    const fileOpen = await isFileOpenInExcel(filePath);
    console.log(`   Excel running: ${excelRunning ? '✅' : '❌'}`);
    console.log(`   File open: ${fileOpen ? '✅' : '❌'}`);

    if (!excelRunning || !fileOpen) {
      console.log('\n❌ Excel is not running or file is not open');
      process.exit(1);
    }

    console.log('\n━'.repeat(60));
    console.log('🚀 Starting tests... Watch Excel for real-time changes!\n');

    // Test 1: Update single cell
    console.log('1️⃣  UPDATE CELL - Setting A1 to "🎯 TEST HEADER"');
    await updateCellViaAppleScript(filePath, sheetName, 'A1', '🎯 TEST HEADER');
    await sleep(1000);
    console.log('   ✅ Done\n');

    // Test 2: Write range
    console.log('2️⃣  WRITE RANGE - Setting B1:D1 with headers');
    await writeRangeViaAppleScript(filePath, sheetName, 'B1', [
      ['Quantity', 'Price', 'Total']
    ]);
    await sleep(1000);
    console.log('   ✅ Done\n');

    // Test 3: Add row with data
    console.log('3️⃣  ADD ROW - Adding new product at end of sheet');
    await addRowViaAppleScript(filePath, sheetName, [
      'New Product',
      100,
      9.99,
      '=B52*C52'
    ]);
    await sleep(1000);
    console.log('   ✅ Done\n');

    // Test 4: Set formula
    console.log('4️⃣  SET FORMULA - Adding formula in D2: =B2*C2');
    await setFormulaViaAppleScript(filePath, sheetName, 'D2', 'B2*C2');
    await sleep(1000);
    console.log('   ✅ Done\n');

    // Test 5: Format cell (make header bold)
    console.log('5️⃣  FORMAT CELL - Making A1 bold with blue background');
    await formatCellViaAppleScript(filePath, sheetName, 'A1', {
      bold: true,
      fontSize: 14,
      fontColor: 'FFFFFF',
      fillColor: '0066CC',
      horizontalAlign: 'center'
    });
    await sleep(1000);
    console.log('   ✅ Done\n');

    // Test 6: Set column width
    console.log('6️⃣  COLUMN WIDTH - Making column A wider (25 units)');
    await setColumnWidthViaAppleScript(filePath, sheetName, 'A', 25);
    await sleep(1000);
    console.log('   ✅ Done\n');

    // Test 7: Set row height
    console.log('7️⃣  ROW HEIGHT - Making row 1 taller (30 units)');
    await setRowHeightViaAppleScript(filePath, sheetName, 1, 30);
    await sleep(1000);
    console.log('   ✅ Done\n');

    // Test 8: Merge cells
    console.log('8️⃣  MERGE CELLS - Merging E1:F1 for title');
    await updateCellViaAppleScript(filePath, sheetName, 'E1', '📊 Statistics');
    await mergeCellsViaAppleScript(filePath, sheetName, 'E1:F1');
    await sleep(1000);
    console.log('   ✅ Done\n');

    // Save
    console.log('9️⃣  SAVE - Saving all changes');
    await saveFileViaAppleScript(filePath);
    console.log('   ✅ Done\n');

    console.log('━'.repeat(60));
    console.log('\n🎉 ALL TESTS PASSED!\n');
    console.log('Results summary:');
    console.log('  ✅ Cell update');
    console.log('  ✅ Range writing');
    console.log('  ✅ Row addition');
    console.log('  ✅ Formula setting');
    console.log('  ✅ Cell formatting');
    console.log('  ✅ Column width adjustment');
    console.log('  ✅ Row height adjustment');
    console.log('  ✅ Cell merging');
    console.log('  ✅ File saving');
    console.log('\n💡 All changes were applied INSTANTLY in Excel!\n');

  } catch (error) {
    console.error('\n❌ Test failed:', error.message);
    console.error('Stack:', error.stack);
    process.exit(1);
  }
}

comprehensiveTest().catch(error => {
  console.error('Fatal error:', error);
  process.exit(1);
});
