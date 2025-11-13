import { loadWorkbook, getSheet, saveWorkbook, parseRange } from './helpers.js';

export async function createChart(
  filePath: string,
  sheetName: string,
  chartType: 'line' | 'bar' | 'column' | 'pie' | 'scatter' | 'area',
  dataRange: string,
  position: string,
  title?: string,
  showLegend: boolean = true,
  createBackup: boolean = false
): Promise<string> {
  const workbook = await loadWorkbook(filePath);
  const sheet = getSheet(workbook, sheetName);

  // Validate the data range
  parseRange(dataRange);

  // Note: ExcelJS has limited chart support compared to the native Excel API
  // This is a simplified implementation that works with ExcelJS capabilities

  // For ExcelJS, we need to use the worksheet's addImage method with chart-like data
  // Since ExcelJS doesn't have full native chart support, we'll create a note about this

  const chartInfo = {
    type: chartType,
    dataRange,
    position,
    title,
    showLegend,
    note: 'Chart placeholder created. For full chart support, use native Excel automation or libraries with full chart APIs.'
  };

  // Add a comment/note at the position indicating the chart
  const posCell = sheet.getCell(position);
  posCell.value = title || `${chartType.toUpperCase()} Chart`;
  posCell.note = JSON.stringify(chartInfo, null, 2);

  // Apply some visual formatting to indicate it's a chart placeholder
  posCell.font = { bold: true, size: 14 };
  posCell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFE7E6E6' },
  };

  await saveWorkbook(workbook, filePath, createBackup);

  return JSON.stringify({
    success: true,
    message: `Chart placeholder created at ${position}`,
    note: 'ExcelJS has limited native chart support. The chart metadata has been saved as a cell note.',
    chartType,
    dataRange,
    position,
    title,
    recommendation: 'For full chart creation, consider using Python openpyxl or native Excel automation.',
  }, null, 2);
}
