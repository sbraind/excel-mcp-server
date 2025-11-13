import { loadWorkbook, getSheet } from './helpers.js';
import type { ResponseFormat } from '../types.js';

export async function validateFormulaSyntax(formula: string): Promise<string> {
  // Basic formula validation
  const cleanFormula = formula.startsWith('=') ? formula.substring(1) : formula;

  const errors: string[] = [];

  // Check for basic syntax errors
  const openParens = (cleanFormula.match(/\(/g) || []).length;
  const closeParens = (cleanFormula.match(/\)/g) || []).length;

  if (openParens !== closeParens) {
    errors.push('Mismatched parentheses');
  }

  // Check for invalid characters at the start
  if (/^[+\-*/]/.test(cleanFormula)) {
    errors.push('Formula cannot start with an operator');
  }

  // Check for empty formula
  if (cleanFormula.trim().length === 0) {
    errors.push('Formula is empty');
  }

  // Check for double operators
  if (/[+\-*/]{2,}/.test(cleanFormula)) {
    errors.push('Invalid consecutive operators');
  }

  // Check for basic function names (common Excel functions)
  const functionPattern = /\b([A-Z]+)\(/g;
  const matches = cleanFormula.match(functionPattern);
  const knownFunctions = [
    'SUM', 'AVERAGE', 'COUNT', 'MAX', 'MIN', 'IF', 'VLOOKUP', 'HLOOKUP',
    'INDEX', 'MATCH', 'CONCATENATE', 'LEFT', 'RIGHT', 'MID', 'LEN',
    'UPPER', 'LOWER', 'TRIM', 'ROUND', 'ROUNDUP', 'ROUNDDOWN',
    'TODAY', 'NOW', 'DATE', 'YEAR', 'MONTH', 'DAY', 'SUMIF', 'COUNTIF',
    'AVERAGEIF', 'SUMIFS', 'COUNTIFS', 'AVERAGEIFS', 'AND', 'OR', 'NOT',
  ];

  if (matches) {
    matches.forEach(match => {
      const funcName = match.slice(0, -1);
      if (!knownFunctions.includes(funcName) && !/^[A-Z]{1,3}\d*$/.test(funcName)) {
        errors.push(`Unknown function or potential error: ${funcName}`);
      }
    });
  }

  const isValid = errors.length === 0;

  return JSON.stringify({
    valid: isValid,
    formula: `=${cleanFormula}`,
    errors: isValid ? [] : errors,
    note: 'This is a basic syntax validation. For full validation, the formula should be tested in Excel.',
  }, null, 2);
}

export async function validateExcelRange(range: string): Promise<string> {
  const rangePattern = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/;
  const match = range.match(rangePattern);

  if (!match) {
    return JSON.stringify({
      valid: false,
      range,
      error: 'Invalid range format. Expected format: A1:D10',
    }, null, 2);
  }

  const [, startCol, startRowStr, endCol, endRowStr] = match;
  const startRow = parseInt(startRowStr);
  const endRow = parseInt(endRowStr);

  const errors: string[] = [];

  // Check if start row is less than end row
  if (startRow > endRow) {
    errors.push('Start row must be less than or equal to end row');
  }

  // Check if start column is less than end column (alphabetically)
  if (startCol > endCol) {
    errors.push('Start column must be less than or equal to end column');
  }

  // Check for valid row numbers (Excel max is 1048576)
  if (startRow < 1 || endRow < 1) {
    errors.push('Row numbers must be greater than 0');
  }

  if (startRow > 1048576 || endRow > 1048576) {
    errors.push('Row numbers exceed Excel maximum (1048576)');
  }

  const isValid = errors.length === 0;

  return JSON.stringify({
    valid: isValid,
    range,
    startCell: `${startCol}${startRow}`,
    endCell: `${endCol}${endRow}`,
    errors: isValid ? [] : errors,
  }, null, 2);
}

export async function getDataValidationInfo(
  filePath: string,
  sheetName: string,
  cellAddress: string,
  responseFormat: ResponseFormat = 'json'
): Promise<string> {
  const workbook = await loadWorkbook(filePath);
  const sheet = getSheet(workbook, sheetName);

  const cell = sheet.getCell(cellAddress);

  // ExcelJS doesn't have full data validation support in reading
  // This is a placeholder for the feature
  const validation = {
    cellAddress,
    hasValidation: false,
    type: 'none',
    note: 'ExcelJS has limited data validation reading support. This feature shows basic cell information.',
  };

  const cellInfo = {
    address: cellAddress,
    value: cell.value,
    type: cell.type,
    dataValidation: validation,
  };

  if (responseFormat === 'markdown') {
    let md = `# Data Validation Info: ${cellAddress}\n\n`;
    md += `**Sheet**: ${sheetName}\n`;
    md += `**Value**: ${cell.value}\n`;
    md += `**Type**: ${cell.type}\n\n`;
    md += `**Note**: ${validation.note}\n`;
    return md;
  }

  return JSON.stringify(cellInfo, null, 2);
}
