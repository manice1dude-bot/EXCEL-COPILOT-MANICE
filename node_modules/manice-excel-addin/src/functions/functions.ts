/**
 * Manice Custom Functions
 * Implementation of AI-powered Excel custom functions
 */

import { AIService } from '../services/aiService';
import { ExcelService } from '../services/excelService';

// Initialize services
const aiService = new AIService();
const excelService = new ExcelService();

/**
 * Main Manice function - Natural language AI assistant
 * @param instruction Natural language instruction
 * @param model Optional: Force specific model (large/small)
 * @param range Optional: Target range (default: current selection)
 * @returns AI-generated result
 */
CustomFunctions.associate("MANICE", manice);
async function manice(
  instruction: string, 
  model?: string,
  range?: string
): Promise<any> {
  try {
    // Validate instruction
    if (!instruction || instruction.trim().length === 0) {
      throw new Error("Please provide an instruction for Manice AI");
    }

    // Get Excel context
    const context = await excelService.getCurrentContext(range);
    
    // Call AI service
    const response = await aiService.processInstruction({
      instruction: instruction.trim(),
      context: context,
      force_model: model,
      stream: false
    });

    // Handle different response types
    if (response.action === 'error') {
      throw new Error(response.explanation);
    }

    // Execute Excel operations if any
    if (response.excel_operations && response.excel_operations.length > 0) {
      await excelService.executeOperations(response.excel_operations);
    }

    // Return appropriate result
    switch (response.action) {
      case 'formula':
        return response.parameters?.formula || response.explanation;
      case 'value':
        return response.parameters?.value || response.explanation;
      case 'text_response':
        return response.explanation;
      case 'calculation':
        return response.parameters?.result || response.explanation;
      default:
        return response.explanation;
    }

  } catch (error) {
    console.error('Manice function error:', error);
    return `Error: ${error.message}`;
  }
}

/**
 * Analyze data with AI insights
 * @param range Data range to analyze
 * @param analysisType Type of analysis (summary, trends, insights, forecast)
 * @returns AI analysis results
 */
CustomFunctions.associate("MANICE.ANALYZE", maniceAnalyze);
async function maniceAnalyze(
  range: any[][], 
  analysisType: string = "summary"
): Promise<string> {
  try {
    if (!range || range.length === 0) {
      throw new Error("Please provide data range to analyze");
    }

    const instruction = `Analyze this data and provide ${analysisType}: ${JSON.stringify(range)}`;
    
    const response = await aiService.processInstruction({
      instruction,
      context: {
        cell_data: { range: range },
        analysis_type: analysisType
      },
      force_model: "large", // Use large model for analysis
      stream: false
    });

    return response.explanation || "Analysis completed";

  } catch (error) {
    console.error('Manice analyze error:', error);
    return `Error: ${error.message}`;
  }
}

/**
 * Generate Excel formula from natural language
 * @param description Natural language description of desired formula
 * @param context Optional: Reference cell/range context
 * @returns Generated Excel formula
 */
CustomFunctions.associate("MANICE.FORMULA", maniceFormula);
async function maniceFormula(
  description: string, 
  context?: string
): Promise<string> {
  try {
    if (!description || description.trim().length === 0) {
      throw new Error("Please provide a formula description");
    }

    const instruction = `Generate an Excel formula for: ${description}`;
    
    const response = await aiService.processInstruction({
      instruction,
      context: context ? { reference_context: context } : undefined,
      force_model: "small", // Use small model for formulas
      stream: false
    });

    // Extract formula from response
    if (response.parameters?.formula) {
      return response.parameters.formula;
    }
    
    // Try to extract formula from explanation
    const formulaMatch = response.explanation.match(/=[\w\s\(\),\+\-\*\/\$:]+/);
    if (formulaMatch) {
      return formulaMatch[0];
    }

    return response.explanation;

  } catch (error) {
    console.error('Manice formula error:', error);
    return `Error: ${error.message}`;
  }
}

/**
 * Clean and format data using AI
 * @param range Data range to clean
 * @param cleaningType Cleaning instructions
 * @returns Cleaned data array
 */
CustomFunctions.associate("MANICE.CLEAN", maniceClean);
async function maniceClean(
  range: any[][], 
  cleaningType: string = "basic"
): Promise<any[][]> {
  try {
    if (!range || range.length === 0) {
      throw new Error("Please provide data range to clean");
    }

    const instruction = `Clean this data using ${cleaningType} cleaning: ${JSON.stringify(range)}`;
    
    const response = await aiService.processInstruction({
      instruction,
      context: {
        cell_data: { range: range },
        cleaning_type: cleaningType
      },
      force_model: "small",
      stream: false
    });

    // Return cleaned data if provided
    if (response.parameters?.cleaned_data) {
      return response.parameters.cleaned_data;
    }

    // Fallback: return original data with note
    return range.map(row => [...row, "(cleaned)"]);

  } catch (error) {
    console.error('Manice clean error:', error);
    return [[`Error: ${error.message}`]];
  }
}

/**
 * Predict or forecast values using AI
 * @param historicalData Historical data range for prediction
 * @param periods Number of periods to predict
 * @param method Prediction method
 * @returns Predicted values array
 */
CustomFunctions.associate("MANICE.PREDICT", manicePredict);
async function manicePredict(
  historicalData: any[][], 
  periods: number,
  method: string = "linear"
): Promise<any[]> {
  try {
    if (!historicalData || historicalData.length === 0) {
      throw new Error("Please provide historical data for prediction");
    }

    if (!periods || periods <= 0) {
      throw new Error("Please specify a valid number of periods to predict");
    }

    const instruction = `Predict ${periods} future values using ${method} method from this historical data: ${JSON.stringify(historicalData)}`;
    
    const response = await aiService.processInstruction({
      instruction,
      context: {
        historical_data: historicalData,
        prediction_periods: periods,
        method: method
      },
      force_model: "large", // Use large model for predictions
      stream: false
    });

    // Return predictions if provided
    if (response.parameters?.predictions) {
      return response.parameters.predictions;
    }

    // Fallback: simple linear trend
    if (historicalData.length >= 2) {
      const lastValue = historicalData[historicalData.length - 1][0];
      const secondLastValue = historicalData[historicalData.length - 2][0];
      const trend = lastValue - secondLastValue;
      
      const predictions = [];
      for (let i = 1; i <= periods; i++) {
        predictions.push(lastValue + (trend * i));
      }
      return predictions;
    }

    return [0]; // Fallback

  } catch (error) {
    console.error('Manice predict error:', error);
    return [`Error: ${error.message}`];
  }
}

/**
 * Generate business insights from data
 * @param range Data range to analyze for insights
 * @param context Business context
 * @returns Business insights and recommendations
 */
CustomFunctions.associate("MANICE.INSIGHTS", maniceInsights);
async function maniceInsights(
  range: any[][], 
  context: string = "general"
): Promise<string> {
  try {
    if (!range || range.length === 0) {
      throw new Error("Please provide data range for insights");
    }

    const instruction = `Generate business insights for ${context} context from this data: ${JSON.stringify(range)}`;
    
    const response = await aiService.processInstruction({
      instruction,
      context: {
        cell_data: { range: range },
        business_context: context
      },
      force_model: "large", // Use large model for insights
      stream: false
    });

    return response.explanation || "Insights generated successfully";

  } catch (error) {
    console.error('Manice insights error:', error);
    return `Error: ${error.message}`;
  }
}

// Streaming function for real-time responses (Office.js doesn't support this yet, but prepared for future)
/*
async function maniceStream(instruction: string): Promise<string> {
  try {
    const response = await aiService.processInstruction({
      instruction,
      stream: true
    });

    // In the future, this could return a streaming result
    // For now, return the complete response
    return response.explanation;

  } catch (error) {
    console.error('Manice stream error:', error);
    return `Error: ${error.message}`;
  }
}
*/

// Export functions for testing
export {
  manice,
  maniceAnalyze,
  maniceFormula,
  maniceClean,
  manicePredict,
  maniceInsights
};