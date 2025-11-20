#!/usr/bin/env node

/**
 * Test the fixed formatting functions
 */

import {
  isExcelRunning,
  isFileOpenInExcel,
  formatCellViaAppleScript,
  setColumnWidthViaAppleScript,
  setRowHeightViaAppleScript,
  updateCellViaAppleScript,
  saveFileViaAppleScript
} from './dist/tools/excel-applescript.js';
import { resolve } from 'path';

async function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

async function testFormatFixes() {
  console.log('\n🔧 TESTING FIXED FORMATTING FUNCTIONS\n');
  console.log('━'.repeat(60));

  const filePath = resolve('./test.xlsx');
  const sheetName = 'Sales';

  try {
    // Pre-flight checks
    console.log('\n📋 Checks:');
    const excelRunning = await isExcelRunning();
    const fileOpen = await isFileOpenInExcel(filePath);
    console.log(`   Excel: ${excelRunning ? '✅' : '❌'} | File open: ${fileOpen ? '✅' : '❌'}`);

    if (!excelRunning || !fileOpen) {
      console.log('\n❌ Excel not running or file not open');
      process.exit(1);
    }

    console.log('\n🚀 Testing FIXED functions... Watch Excel!\n');
    console.log('━'.repeat(60) + '\n');

    // Setup - Put test value in cell
    console.log('📝 Setup: Writing "FORMAT TEST" to cell G1');
    await updateCellViaAppleScript(filePath, sheetName, 'G1', 'FORMAT TEST');
    await sleep(500);
    console.log('   ✅ Done\n');

    // Test 1: Format Cell (FIXED)
    console.log('1️⃣  FORMAT CELL G1:');
    console.log('   - Font: Arial, 14pt, Bold, Blue');
    console.log('   - Fill: Yellow background');
    console.log('   - Alignment: Center');

    try {
      await formatCellViaAppleScript(filePath, sheetName, 'G1', {
        fontName: 'Arial',
        fontSize: 14,
        fontBold: true,
        fontColor: '0000FF',  // Blue
        fillColor: 'FFFF00',  // Yellow
        horizontalAlign: 'center'
      });
      await sleep(1000);
      console.log('   ✅ SUCCESS - Cell formatted!\n');
    } catch (error) {
      console.log(`   ❌ FAILED: ${error.message}\n`);
    }

    // Test 2: Set Column Width (FIXED)
    console.log('2️⃣  SET COLUMN WIDTH G → 35 units');
    try {
      await setColumnWidthViaAppleScript(filePath, sheetName, 'G', 35);
      await sleep(1000);
      console.log('   ✅ SUCCESS - Column widened!\n');
    } catch (error) {
      console.log(`   ❌ FAILED: ${error.message}\n`);
    }

    // Test 3: Set Row Height (FIXED)
    console.log('3️⃣  SET ROW HEIGHT 1 → 35 units');
    try {
      await setRowHeightViaAppleScript(filePath, sheetName, 1, 35);
      await sleep(1000);
      console.log('   ✅ SUCCESS - Row height increased!\n');
    } catch (error) {
      console.log(`   ❌ FAILED: ${error.message}\n`);
    }

    // Additional formatting tests
    console.log('4️⃣  FORMAT CELL H1 (different style):');
    await updateCellViaAppleScript(filePath, sheetName, 'H1', '✨ STYLED');

    try {
      await formatCellViaAppleScript(filePath, sheetName, 'H1', {
        fontName: 'Courier',
        fontSize: 12,
        fontBold: false,
        fontItalic: true,
        fontColor: 'FF0000',  // Red
        fillColor: 'E0E0E0'   // Light gray
      });
      await sleep(1000);
      console.log('   ✅ SUCCESS - Different style applied!\n');
    } catch (error) {
      console.log(`   ❌ FAILED: ${error.message}\n`);
    }

    // Save
    console.log('5️⃣  SAVE FILE');
    await saveFileViaAppleScript(filePath);
    console.log('   ✅ Done\n');

    console.log('━'.repeat(60));
    console.log('\n🎉 FORMATTING TESTS COMPLETED!\n');
    console.log('Check Excel - you should see:');
    console.log('  • G1: Arial 14pt, Bold, Blue text, Yellow background');
    console.log('  • H1: Courier 12pt, Italic, Red text, Gray background');
    console.log('  • Column G: Wider (35 units)');
    console.log('  • Row 1: Taller (35 units)\n');

  } catch (error) {
    console.error('\n❌ Test failed:', error.message);
    console.error('Stack:', error.stack);
    process.exit(1);
  }
}

testFormatFixes().catch(error => {
  console.error('Fatal error:', error);
  process.exit(1);
});
