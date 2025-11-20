#!/usr/bin/env node

/**
 * Test script to verify AppleScript live editing functionality
 */

import {
  isExcelRunning,
  isFileOpenInExcel,
  updateCellViaAppleScript,
  readCellViaAppleScript,
  saveFileViaAppleScript
} from './dist/tools/excel-applescript.js';
import { resolve } from 'path';

async function testLiveEditing() {
  console.log('\n🧪 Testing Live Editing Functionality\n');
  console.log('━'.repeat(50));

  const filePath = resolve('./test.xlsx');
  const sheetName = 'Sales';

  try {
    // Step 1: Check if Excel is running
    console.log('\n1️⃣  Checking if Excel is running...');
    const excelRunning = await isExcelRunning();
    console.log(`   ✅ Excel running: ${excelRunning}`);

    if (!excelRunning) {
      console.log('   ⚠️  Excel is not running. Please open Excel first.');
      console.log('   💡 Run: open test.xlsx');
      process.exit(1);
    }

    // Step 2: Check if file is open
    console.log('\n2️⃣  Checking if test.xlsx is open...');
    const fileOpen = await isFileOpenInExcel(filePath);
    console.log(`   ✅ File open: ${fileOpen}`);

    if (!fileOpen) {
      console.log('   ⚠️  File is not open in Excel. Please open it.');
      console.log('   💡 Run: open test.xlsx');
      process.exit(1);
    }

    // Step 3: Read current value
    console.log('\n3️⃣  Reading current value from cell A1...');
    try {
      const originalValue = await readCellViaAppleScript(filePath, sheetName, 'A1');
      console.log(`   📖 Original value: "${originalValue}"`);
    } catch (error) {
      console.log(`   ⚠️  Could not read cell: ${error.message}`);
    }

    // Step 4: Update cell with AppleScript
    console.log('\n4️⃣  Updating cell A1 with AppleScript...');
    const testValue = `✨ LIVE EDIT TEST - ${new Date().toLocaleTimeString()}`;
    await updateCellViaAppleScript(filePath, sheetName, 'A1', testValue);
    console.log(`   ✅ Cell updated to: "${testValue}"`);
    console.log('   👀 CHECK EXCEL NOW - you should see the change immediately!');

    // Wait for user to verify
    console.log('\n   ⏸️  Waiting 3 seconds for you to verify...');
    await new Promise(resolve => setTimeout(resolve, 3000));

    // Step 5: Read back the value
    console.log('\n5️⃣  Reading back the updated value...');
    const newValue = await readCellViaAppleScript(filePath, sheetName, 'A1');
    console.log(`   📖 New value: "${newValue}"`);

    // Step 6: Verify
    if (newValue === testValue) {
      console.log('\n   ✅ SUCCESS! Live editing is working!');
    } else {
      console.log('\n   ⚠️  Warning: Value mismatch');
      console.log(`   Expected: "${testValue}"`);
      console.log(`   Got: "${newValue}"`);
    }

    // Step 7: Save the file
    console.log('\n6️⃣  Saving file via AppleScript...');
    await saveFileViaAppleScript(filePath);
    console.log('   ✅ File saved');

    console.log('\n━'.repeat(50));
    console.log('🎉 Live editing test completed successfully!\n');

  } catch (error) {
    console.error('\n❌ Error during test:', error.message);
    console.error('Stack:', error.stack);
    process.exit(1);
  }
}

// Run the test
testLiveEditing().catch(error => {
  console.error('Fatal error:', error);
  process.exit(1);
});
