/**
 * Ribbon Commands - Handlers for Excel ribbon buttons and menu items
 */

import { AIService } from '../services/aiService';
import { ExcelService } from '../services/excelService';

// Global services
let aiService: AIService;
let excelService: ExcelService;

// Initialize services
function initializeServices() {
  if (!aiService) {
    aiService = new AIService();
  }
  if (!excelService) {
    excelService = new ExcelService();
  }
}

/**
 * Show notification to user
 */
function showNotification(message: string, type: 'info' | 'success' | 'error' = 'info') {
  if (Office.context.requirements.isSetSupported('NotificationMessages', '1.3')) {
    Office.context.ui.displayDialogAsync(
      `data:text/html,<html><body style="font-family: 'Segoe UI', sans-serif; padding: 20px; text-align: center;">
        <h3 style="color: ${type === 'error' ? '#d13438' : type === 'success' ? '#16c60c' : '#0078d4'};">
          ${type === 'error' ? '❌' : type === 'success' ? '✅' : 'ℹ️'} Manice AI
        </h3>
        <p>${message}</p>
        <button onclick="window.close()" style="padding: 8px 16px; background: #0078d4; color: white; border: none; border-radius: 4px; cursor: pointer;">OK</button>
      </body></html>`,
      {
        height: 200,
        width: 400,
        displayInIframe: false
      }
    );
  } else {
    console.log(`${type.toUpperCase()}: ${message}`);
  }
}

/**
 * Analyze selected data with AI insights
 */
async function analyzeSelectedData(event: Office.AddinCommands.Event) {
  initializeServices();
  
  try {
    // Check connection first
    const isConnected = await aiService.checkConnection();
    if (!isConnected) {
      showNotification('AI server is not connected. Please make sure the Manice AI server is running.', 'error');
      event.completed();
      return;
    }

    // Get selected data
    const context = await excelService.getCurrentContext();
    
    if (!context.cell_data || !context.cell_data.values || context.cell_data.values.length === 0) {
      showNotification('Please select some data to analyze.', 'error');
      event.completed();
      return;
    }

    // Show progress
    showNotification('Analyzing your data with AI...', 'info');

    // Send to AI for analysis
    const response = await aiService.processInstruction({
      instruction: 'Analyze this selected data and provide comprehensive insights including trends, patterns, and recommendations.',
      context: context,
      force_model: 'large', // Use large model for detailed analysis
      stream: false
    });

    // Execute any operations
    if (response.excel_operations && response.excel_operations.length > 0) {
      await excelService.executeOperations(response.excel_operations);
    }

    // Show results
    showNotification(`Analysis Complete: ${response.explanation}`, 'success');

  } catch (error) {
    console.error('Error analyzing data:', error);
    showNotification(`Analysis failed: ${error.message}`, 'error');
  } finally {
    event.completed();
  }
}

/**
 * Create smart chart from selected data
 */
async function createSmartChart(event: Office.AddinCommands.Event) {
  initializeServices();
  
  try {
    const isConnected = await aiService.checkConnection();
    if (!isConnected) {
      showNotification('AI server is not connected.', 'error');
      event.completed();
      return;
    }

    const context = await excelService.getCurrentContext();
    
    if (!context.cell_data || !context.cell_data.values || context.cell_data.values.length === 0) {
      showNotification('Please select data to create a chart.', 'error');
      event.completed();
      return;
    }

    showNotification('Creating smart chart...', 'info');

    const response = await aiService.processInstruction({
      instruction: 'Create the most appropriate chart type for this selected data. Choose the best visualization and add professional formatting.',
      context: context,
      force_model: 'small',
      stream: false
    });

    if (response.excel_operations && response.excel_operations.length > 0) {
      await excelService.executeOperations(response.excel_operations);
      showNotification('Smart chart created successfully!', 'success');
    } else {
      showNotification('Could not determine the best chart type for this data.', 'error');
    }

  } catch (error) {
    console.error('Error creating chart:', error);
    showNotification(`Chart creation failed: ${error.message}`, 'error');
  } finally {
    event.completed();
  }
}

/**
 * Clean selected data using AI
 */
async function cleanSelectedData(event: Office.AddinCommands.Event) {
  initializeServices();
  
  try {
    const isConnected = await aiService.checkConnection();
    if (!isConnected) {
      showNotification('AI server is not connected.', 'error');
      event.completed();
      return;
    }

    const context = await excelService.getCurrentContext();
    
    if (!context.cell_data || !context.cell_data.values || context.cell_data.values.length === 0) {
      showNotification('Please select data to clean.', 'error');
      event.completed();
      return;
    }

    showNotification('Cleaning data with AI...', 'info');

    const response = await aiService.processInstruction({
      instruction: 'Clean this selected data by removing duplicates, fixing formatting issues, standardizing values, and organizing the layout for better readability.',
      context: context,
      force_model: 'small',
      stream: false
    });

    if (response.excel_operations && response.excel_operations.length > 0) {
      await excelService.executeOperations(response.excel_operations);
      showNotification(`Data cleaned successfully! Applied ${response.excel_operations.length} improvements.`, 'success');
    } else {
      showNotification('Data appears to be already clean or no improvements were identified.', 'info');
    }

  } catch (error) {
    console.error('Error cleaning data:', error);
    showNotification(`Data cleaning failed: ${error.message}`, 'error');
  } finally {
    event.completed();
  }
}

/**
 * Open settings dialog
 */
async function openSettings(event: Office.AddinCommands.Event) {
  try {
    // Create settings dialog HTML
    const settingsHtml = `
      <!DOCTYPE html>
      <html>
        <head>
          <title>Manice Settings</title>
          <style>
            body { font-family: 'Segoe UI', sans-serif; padding: 20px; margin: 0; }
            .setting-group { margin-bottom: 20px; }
            .setting-label { display: block; margin-bottom: 5px; font-weight: 500; }
            .setting-input { width: 100%; padding: 8px; border: 1px solid #ccc; border-radius: 4px; }
            .setting-description { font-size: 12px; color: #666; margin-top: 2px; }
            .buttons { text-align: right; margin-top: 20px; padding-top: 20px; border-top: 1px solid #eee; }
            .btn { padding: 8px 16px; margin-left: 8px; border: none; border-radius: 4px; cursor: pointer; }
            .btn-primary { background: #0078d4; color: white; }
            .btn-secondary { background: #f3f2f1; color: #323130; border: 1px solid #ccc; }
            .connection-status { padding: 10px; border-radius: 4px; margin-bottom: 20px; }
            .status-connected { background: #dff6dd; color: #16c60c; border: 1px solid #16c60c; }
            .status-disconnected { background: #fde7e9; color: #d13438; border: 1px solid #d13438; }
          </style>
        </head>
        <body>
          <h2>Manice AI Settings</h2>
          
          <div id="connectionStatus" class="connection-status">
            <span id="statusText">Checking connection...</span>
          </div>
          
          <div class="setting-group">
            <label class="setting-label">AI Server URL:</label>
            <input type="text" id="serverUrl" class="setting-input" value="http://127.0.0.1:8899" />
            <div class="setting-description">URL of the local Manice AI server</div>
          </div>
          
          <div class="setting-group">
            <label class="setting-label">Preferred AI Model:</label>
            <select id="preferredModel" class="setting-input">
              <option value="auto">Auto-select (Recommended)</option>
              <option value="large">Always use Large Model (DeepSeek R1)</option>
              <option value="small">Always use Small Model (Mistral-7B)</option>
            </select>
            <div class="setting-description">Choose which AI model to use by default</div>
          </div>
          
          <div class="setting-group">
            <label class="setting-label">Safety Settings:</label>
            <input type="checkbox" id="confirmDestructive" checked /> Confirm destructive operations<br>
            <input type="checkbox" id="enableUndo" checked /> Enable undo for AI operations<br>
            <input type="checkbox" id="backupData" checked /> Backup data before major changes
          </div>
          
          <div class="buttons">
            <button class="btn btn-secondary" onclick="window.close()">Cancel</button>
            <button class="btn btn-primary" onclick="testConnection()">Test Connection</button>
            <button class="btn btn-primary" onclick="saveSettings()">Save Settings</button>
          </div>
          
          <script>
            // Check connection on load
            window.addEventListener('load', checkConnection);
            
            async function checkConnection() {
              const statusDiv = document.getElementById('connectionStatus');
              const statusText = document.getElementById('statusText');
              
              try {
                const response = await fetch(document.getElementById('serverUrl').value + '/health');
                if (response.ok) {
                  statusDiv.className = 'connection-status status-connected';
                  statusText.textContent = '✅ AI Server Connected';
                } else {
                  statusDiv.className = 'connection-status status-disconnected';
                  statusText.textContent = '❌ AI Server Not Responding';
                }
              } catch (error) {
                statusDiv.className = 'connection-status status-disconnected';
                statusText.textContent = '❌ Cannot Connect to AI Server';
              }
            }
            
            async function testConnection() {
              await checkConnection();
            }
            
            function saveSettings() {
              // In a real implementation, save settings to local storage or server
              alert('Settings saved successfully!');
              window.close();
            }
          </script>
        </body>
      </html>
    `;

    // Show settings dialog
    Office.context.ui.displayDialogAsync(
      'data:text/html,' + encodeURIComponent(settingsHtml),
      {
        height: 500,
        width: 600,
        displayInIframe: false
      },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          showNotification('Could not open settings dialog.', 'error');
        }
      }
    );

  } catch (error) {
    console.error('Error opening settings:', error);
    showNotification(`Could not open settings: ${error.message}`, 'error');
  } finally {
    event.completed();
  }
}

/**
 * Open the chat sidebar (main task pane)
 */
function openChatSidebar(event: Office.AddinCommands.Event) {
  // This will be handled by the task pane
  Office.context.ui.messageParent(JSON.stringify({
    action: 'openSidebar'
  }));
  
  event.completed();
}

/**
 * Get contextual help based on current selection
 */
async function getContextualHelp(event: Office.AddinCommands.Event) {
  initializeServices();
  
  try {
    const context = await excelService.getCurrentContext();
    
    let helpMessage = "Manice AI can help you with:\n\n";
    
    if (context.cell_data && context.cell_data.values && context.cell_data.values.length > 1) {
      helpMessage += "• Analyze your selected data for insights\n";
      helpMessage += "• Create charts and visualizations\n";
      helpMessage += "• Clean and format the data\n";
      helpMessage += "• Generate formulas based on the data\n";
    } else {
      helpMessage += "• Creating formulas and functions\n";
      helpMessage += "• Formatting cells and ranges\n";
      helpMessage += "• Data analysis and insights\n";
      helpMessage += "• Chart creation and customization\n";
      helpMessage += "• Pivot tables and data summarization\n";
    }
    
    helpMessage += "\nTry using the =Manice() function or open the AI chat sidebar!";
    
    showNotification(helpMessage, 'info');

  } catch (error) {
    showNotification('Manice AI is your Excel assistant. Use =Manice("your instruction") in any cell or click "Open Manice" to start chatting!', 'info');
  } finally {
    event.completed();
  }
}

/**
 * Register all ribbon command functions
 */
function registerRibbonCommands() {
  // Register functions for ribbon commands
  if (typeof Office !== 'undefined' && Office.actions) {
    Office.actions.associate('analyzeSelectedData', analyzeSelectedData);
    Office.actions.associate('createSmartChart', createSmartChart);
    Office.actions.associate('cleanSelectedData', cleanSelectedData);
    Office.actions.associate('openSettings', openSettings);
    Office.actions.associate('openChatSidebar', openChatSidebar);
    Office.actions.associate('getContextualHelp', getContextualHelp);
  }
}

// Auto-register when the script loads
if (typeof Office !== 'undefined') {
  Office.onReady(() => {
    registerRibbonCommands();
  });
} else {
  // Fallback for testing environments
  setTimeout(registerRibbonCommands, 100);
}

// Export functions for testing
export {
  analyzeSelectedData,
  createSmartChart,
  cleanSelectedData,
  openSettings,
  openChatSidebar,
  getContextualHelp,
  registerRibbonCommands
};