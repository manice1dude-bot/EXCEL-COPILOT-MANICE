/**
 * Excel Service - Real-time Excel Integration
 * Handles all Excel operations using Office.js API
 */

export interface ExcelOperation {
  type: string;
  target: string;
  value?: any;
  options?: any;
}

export interface ExcelContext {
  sheet_name?: string;
  selected_range?: string;
  cell_data?: any;
  workbook_info?: any;
  previous_operations?: string[];
}

export interface CellInfo {
  address: string;
  value: any;
  formula?: string;
  format?: any;
}

export interface SheetInfo {
  name: string;
  index: number;
  visible: boolean;
  tabColor?: string;
}

export class ExcelService {
  private undoStack: Array<() => Promise<void>> = [];
  private redoStack: Array<() => Promise<void>> = [];
  private maxUndoSteps = 50;

  constructor() {
    this.initializeExcel();
  }

  /**
   * Initialize Excel and set up event handlers
   */
  private async initializeExcel(): Promise<void> {
    try {
      await Excel.run(async (context) => {
        // Set up change tracking for undo functionality
        const worksheet = context.workbook.worksheets.getActiveWorksheet();
        worksheet.load("name");
        await context.sync();
        
        console.log(`ExcelService initialized on sheet: ${worksheet.name}`);
      });
    } catch (error) {
      console.error("Failed to initialize Excel service:", error);
    }
  }

  /**
   * Get current Excel context for AI processing
   */
  async getCurrentContext(targetRange?: string): Promise<ExcelContext> {
    return Excel.run(async (context) => {
      const workbook = context.workbook;
      const worksheet = workbook.worksheets.getActiveWorksheet();
      
      worksheet.load("name, tabColor");
      workbook.load("name");

      let range: Excel.Range;
      if (targetRange) {
        range = worksheet.getRange(targetRange);
      } else {
        range = context.workbook.getSelectedRange();
      }
      
      range.load("address, values, formulas, format");
      
      await context.sync();

      // Get additional workbook info
      const sheets = workbook.worksheets;
      sheets.load("items/name");
      await context.sync();

      const sheetNames = sheets.items.map(sheet => sheet.name);

      return {
        sheet_name: worksheet.name,
        selected_range: range.address,
        cell_data: {
          address: range.address,
          values: range.values,
          formulas: range.formulas,
          format: {
            numberFormat: range.format.numberFormat,
            font: range.format.font
          }
        },
        workbook_info: {
          name: workbook.name,
          sheets: sheetNames,
          active_sheet: worksheet.name
        },
        previous_operations: [] // Will be populated by operation history
      };
    });
  }

  /**
   * Execute a list of Excel operations
   */
  async executeOperations(operations: ExcelOperation[]): Promise<void> {
    const undoOperations: Array<() => Promise<void>> = [];

    try {
      for (const operation of operations) {
        const undoOp = await this.executeOperation(operation);
        if (undoOp) {
          undoOperations.unshift(undoOp); // Add to front for reverse order
        }
      }

      // Add combined undo operation to stack
      if (undoOperations.length > 0) {
        this.addToUndoStack(async () => {
          for (const undoOp of undoOperations) {
            await undoOp();
          }
        });
      }

    } catch (error) {
      console.error("Failed to execute operations:", error);
      // Attempt to undo partial changes
      for (const undoOp of undoOperations) {
        try {
          await undoOp();
        } catch (undoError) {
          console.error("Failed to undo operation:", undoError);
        }
      }
      throw error;
    }
  }

  /**
   * Execute a single Excel operation
   */
  private async executeOperation(operation: ExcelOperation): Promise<(() => Promise<void>) | null> {
    switch (operation.type) {
      case 'cell_edit':
        return this.setCellValue(operation.target, operation.value, operation.options);
      
      case 'formula':
        return this.setFormula(operation.target, operation.value, operation.options);
      
      case 'format':
        return this.formatRange(operation.target, operation.options);
      
      case 'insert_row':
        return this.insertRows(operation.target, operation.value || 1);
      
      case 'insert_column':
        return this.insertColumns(operation.target, operation.value || 1);
      
      case 'delete_row':
        return this.deleteRows(operation.target);
      
      case 'delete_column':
        return this.deleteColumns(operation.target);
      
      case 'chart':
        return this.createChart(operation.target, operation.options);
      
      case 'conditional_format':
        return this.addConditionalFormatting(operation.target, operation.options);
      
      case 'sort':
        return this.sortRange(operation.target, operation.options);
      
      case 'filter':
        return this.applyFilter(operation.target, operation.options);
      
      case 'pivot_table':
        return this.createPivotTable(operation.target, operation.options);
      
      case 'sheet_rename':
        return this.renameSheet(operation.target, operation.value);
      
      case 'sheet_create':
        return this.createSheet(operation.value, operation.options);
      
      default:
        console.warn(`Unknown operation type: ${operation.type}`);
        return null;
    }
  }

  /**
   * Set cell or range values
   */
  private async setCellValue(target: string, value: any, options?: any): Promise<() => Promise<void>> {
    return Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const range = worksheet.getRange(target);
      
      // Store original values for undo
      range.load("values");
      await context.sync();
      const originalValues = range.values;

      // Set new values
      if (Array.isArray(value)) {
        range.values = value;
      } else {
        range.values = [[value]];
      }

      await context.sync();

      // Return undo function
      return async () => {
        return Excel.run(async (context) => {
          const undoRange = context.workbook.worksheets.getActiveWorksheet().getRange(target);
          undoRange.values = originalValues;
          await context.sync();
        });
      };
    });
  }

  /**
   * Set formula in cell or range
   */
  private async setFormula(target: string, formula: string, options?: any): Promise<() => Promise<void>> {
    return Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const range = worksheet.getRange(target);
      
      range.load("formulas");
      await context.sync();
      const originalFormulas = range.formulas;

      // Set new formula
      range.formulas = [[formula]];
      await context.sync();

      return async () => {
        return Excel.run(async (context) => {
          const undoRange = context.workbook.worksheets.getActiveWorksheet().getRange(target);
          undoRange.formulas = originalFormulas;
          await context.sync();
        });
      };
    });
  }

  /**
   * Format range (colors, fonts, borders, etc.)
   */
  private async formatRange(target: string, options: any): Promise<() => Promise<void>> {
    return Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const range = worksheet.getRange(target);
      
      // Load current format for undo
      range.format.load("font, fill, borders");
      await context.sync();
      
      const originalFormat = {
        font: { ...range.format.font },
        fill: { ...range.format.fill },
        borders: { ...range.format.borders }
      };

      // Apply new formatting
      if (options.font) {
        Object.assign(range.format.font, options.font);
      }
      if (options.fill) {
        Object.assign(range.format.fill, options.fill);
      }
      if (options.borders) {
        Object.assign(range.format.borders, options.borders);
      }
      if (options.numberFormat) {
        range.numberFormat = options.numberFormat;
      }

      await context.sync();

      return async () => {
        return Excel.run(async (context) => {
          const undoRange = context.workbook.worksheets.getActiveWorksheet().getRange(target);
          Object.assign(undoRange.format.font, originalFormat.font);
          Object.assign(undoRange.format.fill, originalFormat.fill);
          Object.assign(undoRange.format.borders, originalFormat.borders);
          await context.sync();
        });
      };
    });
  }

  /**
   * Insert rows
   */
  private async insertRows(target: string, count: number = 1): Promise<() => Promise<void>> {
    return Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const range = worksheet.getRange(target);
      
      range.insert(Excel.InsertShiftDirection.down);
      await context.sync();

      // Return undo function
      return async () => {
        return Excel.run(async (context) => {
          const undoRange = context.workbook.worksheets.getActiveWorksheet().getRange(target);
          undoRange.delete(Excel.DeleteShiftDirection.up);
          await context.sync();
        });
      };
    });
  }

  /**
   * Insert columns
   */
  private async insertColumns(target: string, count: number = 1): Promise<() => Promise<void>> {
    return Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const range = worksheet.getRange(target);
      
      range.insert(Excel.InsertShiftDirection.right);
      await context.sync();

      return async () => {
        return Excel.run(async (context) => {
          const undoRange = context.workbook.worksheets.getActiveWorksheet().getRange(target);
          undoRange.delete(Excel.DeleteShiftDirection.left);
          await context.sync();
        });
      };
    });
  }

  /**
   * Delete rows
   */
  private async deleteRows(target: string): Promise<() => Promise<void>> {
    return Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const range = worksheet.getRange(target);
      
      // Store data for undo
      range.load("values, formulas, format");
      await context.sync();
      
      const backupData = {
        values: range.values,
        formulas: range.formulas,
        format: range.format
      };

      range.delete(Excel.DeleteShiftDirection.up);
      await context.sync();

      return async () => {
        return Excel.run(async (context) => {
          const undoRange = context.workbook.worksheets.getActiveWorksheet().getRange(target);
          undoRange.insert(Excel.InsertShiftDirection.down);
          undoRange.values = backupData.values;
          undoRange.formulas = backupData.formulas;
          await context.sync();
        });
      };
    });
  }

  /**
   * Delete columns
   */
  private async deleteColumns(target: string): Promise<() => Promise<void>> {
    return Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const range = worksheet.getRange(target);
      
      range.load("values, formulas");
      await context.sync();
      
      const backupData = {
        values: range.values,
        formulas: range.formulas
      };

      range.delete(Excel.DeleteShiftDirection.left);
      await context.sync();

      return async () => {
        return Excel.run(async (context) => {
          const undoRange = context.workbook.worksheets.getActiveWorksheet().getRange(target);
          undoRange.insert(Excel.InsertShiftDirection.right);
          undoRange.values = backupData.values;
          undoRange.formulas = backupData.formulas;
          await context.sync();
        });
      };
    });
  }

  /**
   * Create chart
   */
  private async createChart(target: string, options: any): Promise<() => Promise<void>> {
    return Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const range = worksheet.getRange(target);
      
      const chartType = options.type || "ColumnClustered";
      const chart = worksheet.charts.add(chartType, range, "Auto");
      
      if (options.title) {
        chart.title.text = options.title;
      }
      
      if (options.position) {
        chart.left = options.position.left || 100;
        chart.top = options.position.top || 100;
        chart.width = options.position.width || 400;
        chart.height = options.position.height || 300;
      }

      await context.sync();
      const chartId = chart.id;

      return async () => {
        return Excel.run(async (context) => {
          const chartToDelete = context.workbook.worksheets.getActiveWorksheet().charts.getItem(chartId);
          chartToDelete.delete();
          await context.sync();
        });
      };
    });
  }

  /**
   * Add conditional formatting
   */
  private async addConditionalFormatting(target: string, options: any): Promise<() => Promise<void>> {
    return Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const range = worksheet.getRange(target);
      
      const conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType[options.type]);
      
      if (options.type === "cellValue") {
        conditionalFormat.cellValue.format.fill.color = options.fillColor || "red";
        conditionalFormat.cellValue.rule = options.rule;
      }

      await context.sync();

      return async () => {
        return Excel.run(async (context) => {
          const undoRange = context.workbook.worksheets.getActiveWorksheet().getRange(target);
          undoRange.conditionalFormats.clearAll();
          await context.sync();
        });
      };
    });
  }

  /**
   * Sort range
   */
  private async sortRange(target: string, options: any): Promise<() => Promise<void>> {
    return Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const range = worksheet.getRange(target);
      
      // Store original data for undo
      range.load("values");
      await context.sync();
      const originalValues = range.values;

      // Apply sort
      const sortFields = options.fields || [{ key: 0, ascending: true }];
      range.sort.apply(sortFields, options.hasHeaders || false);
      await context.sync();

      return async () => {
        return Excel.run(async (context) => {
          const undoRange = context.workbook.worksheets.getActiveWorksheet().getRange(target);
          undoRange.values = originalValues;
          await context.sync();
        });
      };
    });
  }

  /**
   * Apply filter
   */
  private async applyFilter(target: string, options: any): Promise<() => Promise<void>> {
    return Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const range = worksheet.getRange(target);
      
      range.autoFilter.apply(range);
      
      if (options.columnIndex && options.filterCriteria) {
        range.autoFilter.filters.getItemAt(options.columnIndex).apply(options.filterCriteria);
      }

      await context.sync();

      return async () => {
        return Excel.run(async (context) => {
          const undoRange = context.workbook.worksheets.getActiveWorksheet().getRange(target);
          undoRange.autoFilter.remove();
          await context.sync();
        });
      };
    });
  }

  /**
   * Create pivot table
   */
  private async createPivotTable(target: string, options: any): Promise<() => Promise<void>> {
    return Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const range = worksheet.getRange(target);
      
      const pivotTable = worksheet.pivotTables.add(
        options.name || "ManicePivotTable",
        range,
        worksheet.getRange(options.destination || "H1")
      );

      if (options.rowFields) {
        options.rowFields.forEach((field: string) => {
          pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem(field));
        });
      }

      if (options.dataFields) {
        options.dataFields.forEach((field: string) => {
          pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem(field));
        });
      }

      await context.sync();
      const pivotTableId = pivotTable.id;

      return async () => {
        return Excel.run(async (context) => {
          const pivotTableToDelete = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem(pivotTableId);
          pivotTableToDelete.delete();
          await context.sync();
        });
      };
    });
  }

  /**
   * Rename sheet
   */
  private async renameSheet(currentName: string, newName: string): Promise<() => Promise<void>> {
    return Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getItem(currentName);
      worksheet.name = newName;
      await context.sync();

      return async () => {
        return Excel.run(async (context) => {
          const undoWorksheet = context.workbook.worksheets.getItem(newName);
          undoWorksheet.name = currentName;
          await context.sync();
        });
      };
    });
  }

  /**
   * Create new sheet
   */
  private async createSheet(name: string, options?: any): Promise<() => Promise<void>> {
    return Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.add(name);
      
      if (options?.tabColor) {
        worksheet.tabColor = options.tabColor;
      }
      
      await context.sync();

      return async () => {
        return Excel.run(async (context) => {
          const sheetToDelete = context.workbook.worksheets.getItem(name);
          sheetToDelete.delete();
          await context.sync();
        });
      };
    });
  }

  /**
   * Undo last operation
   */
  async undo(): Promise<void> {
    if (this.undoStack.length > 0) {
      const undoOperation = this.undoStack.pop()!;
      
      // Execute undo
      await undoOperation();
      
      // Move to redo stack (simplified - would need the original operation)
      // this.redoStack.push(originalOperation);
    }
  }

  /**
   * Add operation to undo stack
   */
  private addToUndoStack(undoOperation: () => Promise<void>): void {
    this.undoStack.push(undoOperation);
    
    // Limit stack size
    if (this.undoStack.length > this.maxUndoSteps) {
      this.undoStack.shift();
    }
    
    // Clear redo stack when new operation is added
    this.redoStack = [];
  }

  /**
   * Get all sheet names
   */
  async getSheetNames(): Promise<string[]> {
    return Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.load("items/name");
      await context.sync();
      
      return sheets.items.map(sheet => sheet.name);
    });
  }

  /**
   * Get range data as array
   */
  async getRangeData(target: string): Promise<any[][]> {
    return Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const range = worksheet.getRange(target);
      range.load("values");
      await context.sync();
      
      return range.values as any[][];
    });
  }

  /**
   * Check if range is valid
   */
  isValidRange(rangeAddress: string): boolean {
    try {
      // Simple regex check for Excel range format
      const rangePattern = /^[A-Z]+[0-9]+(:([A-Z]+[0-9]+)?)?$/;
      return rangePattern.test(rangeAddress.toUpperCase());
    } catch {
      return false;
    }
  }

  /**
   * Get current selection info
   */
  async getSelectionInfo(): Promise<{ address: string; values: any[][]; formulas: any[][] }> {
    return Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load("address, values, formulas");
      await context.sync();
      
      return {
        address: range.address,
        values: range.values as any[][],
        formulas: range.formulas as any[][]
      };
    });
  }
}