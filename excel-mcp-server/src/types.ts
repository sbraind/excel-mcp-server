import type { CellValue } from 'exceljs';

export interface CellData {
  address: string;
  value: CellValue;
  formula?: string;
  type?: string;
}

export interface SheetInfo {
  name: string;
  rowCount: number;
  columnCount: number;
  state?: string;
}

export interface WorkbookInfo {
  sheets: SheetInfo[];
  creator?: string;
  created?: Date;
  modified?: Date;
}

export interface CellFormat {
  font?: {
    name?: string;
    size?: number;
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    color?: string;
  };
  fill?: {
    type: 'pattern';
    pattern: 'solid' | 'darkVertical' | 'darkHorizontal' | 'darkGrid';
    fgColor?: string;
    bgColor?: string;
  };
  alignment?: {
    horizontal?: 'left' | 'center' | 'right' | 'fill' | 'justify';
    vertical?: 'top' | 'middle' | 'bottom';
    wrapText?: boolean;
  };
  border?: {
    top?: { style: string; color?: string };
    left?: { style: string; color?: string };
    bottom?: { style: string; color?: string };
    right?: { style: string; color?: string };
  };
  numFmt?: string;
}

export type ResponseFormat = 'json' | 'markdown';

export interface ToolResponse {
  content: Array<{
    type: 'text';
    text: string;
  }>;
}
