/**
 * Formula and VBA Engine Service
 * Integration with AI-powered formula generation and VBA macro creation
 */

import { ExcelService, ExcelContext } from './excelService';
import { AIService } from './aiService';

export interface FormulaRequest {
  requirement: string;
  context?: ExcelContext;
  data_sample?: any;
}

export interface VBARequest {
  requirement: string;
  context?: ExcelContext;
  constraints?: {
    security_level?: string;
    performance_priority?: string;
    max_complexity?: string;
  };
}

export interface FormulaResult {
  success: boolean;
  formula?: string;
  explanation?: string;
  category?: string;
  complexity?: string;
  examples?: string[];
  alternatives?: Array<{formula: string; description: string}>;
  prerequisites?: string[];
  confidence?: number;
  errors?: Array<{type: string; description: string; suggestion: string}>;
  solutions?: Array<{formula: string; description: string}>;
  optimizations?: Array<{type: string; description: string}>;
}

export interface VBAResult {
  success: boolean;
  name?: string;
  code?: string;
  description?: string;
  category?: string;
  complexity?: string;
  security_level?: string;
  dependencies?: string[];
  parameters?: Array<{name: string; type: string; description: string}>;
  usage_examples?: string[];
  warnings?: string[];
  performance_notes?: string[];
  analysis?: {
    security_issues?: string[];
    performance_issues?: string[];
    best_practices?: string[];
    code_quality_score?: number;
    maintainability_score?: number;
    overall_rating?: string;
  };
}

export class FormulaEngineService {
  private static instance: FormulaEngineService;
  private excelService: ExcelService;
  private aiService: AIService;
  private baseUrl: string;

  private constructor() {
    this.excelService = new ExcelService();
    this.aiService = new AIService();
    this.baseUrl = 'http://localhost:8000/api/v1';
  }

  static getInstance(): FormulaEngineService {
    if (!FormulaEngineService.instance) {
      FormulaEngineService.instance = new FormulaEngineService();
    }
    return FormulaEngineService.instance;
  }

  /**
   * Generate Excel formula from natural language requirement
   */
  async generateFormula(requirement: string, targetRange?: string): Promise<FormulaResult> {
    try {
      // Get current Excel context
      const context = await this.excelService.getCurrentContext(targetRange);
      
      // Prepare request
      const request: FormulaRequest = {
        requirement,
        context,
        data_sample: this.extractDataSample(context)
      };

      // Call AI service
      const response = await fetch(`${this.baseUrl}/formula/generate`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify(request)
      });

      if (!response.ok) {
        throw new Error(`Formula generation failed: ${response.status}`);
      }

      const result: FormulaResult = await response.json();
      
      // If successful and we have a formula, optionally apply it
      if (result.success && result.formula && targetRange) {
        try {
          await this.applyFormula(result.formula, targetRange);
        } catch (error) {
          console.warn('Failed to apply generated formula:', error);
        }
      }

      return result;

    } catch (error) {
      console.error('Error generating formula:', error);
      return {
        success: false,
        explanation: `Error: ${error.message}`
      };
    }
  }

  /**
   * Debug an Excel formula
   */
  async debugFormula(formula: string, errorMessage?: string, targetRange?: string): Promise<FormulaResult> {
    try {
      // Get current Excel context
      const context = await this.excelService.getCurrentContext(targetRange);

      const request = {
        formula,
        error_message: errorMessage,
        cell_data: context.cell_data
      };

      const response = await fetch(`${this.baseUrl}/formula/debug`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify(request)
      });

      if (!response.ok) {
        throw new Error(`Formula debugging failed: ${response.status}`);
      }

      return await response.json();

    } catch (error) {
      console.error('Error debugging formula:', error);
      return {
        success: false,
        explanation: `Error: ${error.message}`
      };
    }
  }

  /**
   * Explain how a formula works
   */
  async explainFormula(formula: string): Promise<FormulaResult> {
    try {
      const response = await fetch(`${this.baseUrl}/formula/explain`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({ formula })
      });

      if (!response.ok) {
        throw new Error(`Formula explanation failed: ${response.status}`);
      }

      return await response.json();

    } catch (error) {
      console.error('Error explaining formula:', error);
      return {
        success: false,
        explanation: `Error: ${error.message}`
      };
    }
  }

  /**
   * Generate VBA macro from natural language requirement
   */
  async generateVBAMacro(requirement: string, constraints?: any): Promise<VBAResult> {
    try {
      // Get current Excel context
      const context = await this.excelService.getCurrentContext();

      const request: VBARequest = {
        requirement,
        context,
        constraints: constraints || {
          security_level: 'moderate',
          performance_priority: 'balanced',
          max_complexity: 'moderate'
        }
      };

      const response = await fetch(`${this.baseUrl}/vba/generate`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify(request)
      });

      if (!response.ok) {
        throw new Error(`VBA generation failed: ${response.status}`);
      }

      return await response.json();

    } catch (error) {
      console.error('Error generating VBA macro:', error);
      return {
        success: false,
        description: `Error: ${error.message}`
      };
    }
  }

  /**
   * Analyze existing VBA code
   */
  async analyzeVBACode(code: string): Promise<VBAResult> {
    try {
      const request = {
        code,
        context: await this.excelService.getCurrentContext()
      };

      const response = await fetch(`${this.baseUrl}/vba/analyze`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify(request)
      });

      if (!response.ok) {
        throw new Error(`VBA analysis failed: ${response.status}`);
      }

      return await response.json();

    } catch (error) {
      console.error('Error analyzing VBA code:', error);
      return {
        success: false,
        description: `Error: ${error.message}`
      };
    }
  }

  /**
   * Optimize VBA code
   */
  async optimizeVBACode(code: string, goals?: string[]): Promise<VBAResult> {
    try {
      const response = await fetch(`${this.baseUrl}/vba/optimize`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({ 
          code, 
          optimization_goals: goals || ['performance', 'readability', 'best_practices']
        })
      });

      if (!response.ok) {
        throw new Error(`VBA optimization failed: ${response.status}`);
      }

      return await response.json();

    } catch (error) {
      console.error('Error optimizing VBA code:', error);
      return {
        success: false,
        description: `Error: ${error.message}`
      };
    }
  }

  /**
   * Get available Excel functions
   */
  async getExcelFunctions(category?: string): Promise<any> {
    try {
      const url = category 
        ? `${this.baseUrl}/formula/functions?category=${category}`
        : `${this.baseUrl}/formula/functions`;
        
      const response = await fetch(url);
      
      if (!response.ok) {
        throw new Error(`Failed to get functions: ${response.status}`);
      }

      return await response.json();

    } catch (error) {
      console.error('Error getting Excel functions:', error);
      return { success: false, error: error.message };
    }
  }

  /**
   * Get VBA templates
   */
  async getVBATemplates(category?: string): Promise<any> {
    try {
      const url = category 
        ? `${this.baseUrl}/vba/templates?category=${category}`
        : `${this.baseUrl}/vba/templates`;
        
      const response = await fetch(url);
      
      if (!response.ok) {
        throw new Error(`Failed to get VBA templates: ${response.status}`);
      }

      return await response.json();

    } catch (error) {
      console.error('Error getting VBA templates:', error);
      return { success: false, error: error.message };
    }
  }

  /**
   * Apply formula to specified range
   */
  private async applyFormula(formula: string, targetRange: string): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const range = worksheet.getRange(targetRange);
      
      // Set the formula
      range.formulas = [[formula]];
      
      await context.sync();
    });
  }

  /**
   * Extract data sample from Excel context for AI processing
   */
  private extractDataSample(context: ExcelContext): any {
    if (!context.cell_data) return null;

    try {
      const { values, formulas } = context.cell_data;
      
      // Create a sample of the data (first few rows/columns)
      const sampleSize = 5;
      const sample = {
        values: Array.isArray(values) 
          ? values.slice(0, sampleSize).map((row: any[]) => 
              Array.isArray(row) ? row.slice(0, sampleSize) : row
            )
          : values,
        formulas: Array.isArray(formulas)
          ? formulas.slice(0, sampleSize).map((row: any[]) => 
              Array.isArray(row) ? row.slice(0, sampleSize) : row
            )
          : formulas,
        data_types: this.inferDataTypes(values),
        range_info: {
          address: context.cell_data.address,
          row_count: Array.isArray(values) ? values.length : 1,
          col_count: Array.isArray(values) && Array.isArray(values[0]) ? values[0].length : 1
        }
      };

      return sample;
    } catch (error) {
      console.warn('Failed to extract data sample:', error);
      return null;
    }
  }

  /**
   * Infer data types from values
   */
  private inferDataTypes(values: any): string[] {
    if (!Array.isArray(values)) return ['unknown'];
    
    try {
      const types: string[] = [];
      
      for (let i = 0; i < Math.min(values.length, 5); i++) {
        const row = values[i];
        if (Array.isArray(row)) {
          for (let j = 0; j < Math.min(row.length, 5); j++) {
            const value = row[j];
            if (typeof value === 'number') {
              types.push('number');
            } else if (typeof value === 'string') {
              if (value.match(/^\d{4}-\d{2}-\d{2}/)) {
                types.push('date');
              } else if (value.match(/^\$?\d+\.?\d*$/)) {
                types.push('currency');
              } else {
                types.push('text');
              }
            } else if (value instanceof Date) {
              types.push('date');
            } else {
              types.push('unknown');
            }
          }
        }
      }
      
      return [...new Set(types)]; // Remove duplicates
    } catch (error) {
      return ['unknown'];
    }
  }

  /**
   * Check service health
   */
  async checkHealth(): Promise<any> {
    try {
      const response = await fetch(`${this.baseUrl}/health/engines`);
      return await response.json();
    } catch (error) {
      return { 
        overall_status: 'error', 
        error: error.message,
        timestamp: Date.now() / 1000
      };
    }
  }

  /**
   * Smart formula suggestions based on selected data
   */
  async suggestFormulas(targetRange?: string): Promise<FormulaResult[]> {
    try {
      const context = await this.excelService.getCurrentContext(targetRange);
      const dataSample = this.extractDataSample(context);
      
      if (!dataSample || !dataSample.values) {
        return [{
          success: false,
          explanation: "No data selected to analyze for suggestions"
        }];
      }

      // Analyze data patterns and suggest relevant formulas
      const suggestions: string[] = [];
      
      // Check if numeric data - suggest math functions
      if (dataSample.data_types.includes('number')) {
        suggestions.push('Calculate sum of selected numbers');
        suggestions.push('Calculate average of selected numbers');
        suggestions.push('Find maximum value in selection');
        suggestions.push('Find minimum value in selection');
      }
      
      // Check if has text - suggest text functions
      if (dataSample.data_types.includes('text')) {
        suggestions.push('Concatenate text values');
        suggestions.push('Convert text to uppercase');
        suggestions.push('Extract first word from text');
      }
      
      // Check if has dates - suggest date functions
      if (dataSample.data_types.includes('date')) {
        suggestions.push('Calculate days between dates');
        suggestions.push('Add business days to date');
        suggestions.push('Extract month from date');
      }

      // Generate formulas for each suggestion
      const results: FormulaResult[] = [];
      for (const suggestion of suggestions.slice(0, 3)) { // Limit to 3 suggestions
        try {
          const result = await this.generateFormula(suggestion, targetRange);
          if (result.success) {
            results.push(result);
          }
        } catch (error) {
          console.warn(`Failed to generate formula for suggestion: ${suggestion}`, error);
        }
      }

      return results;

    } catch (error) {
      console.error('Error getting formula suggestions:', error);
      return [{
        success: false,
        explanation: `Error generating suggestions: ${error.message}`
      }];
    }
  }
}