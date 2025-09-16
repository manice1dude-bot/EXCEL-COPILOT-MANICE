/**
 * AI Service - Communication with Manice AI Server
 * Handles all AI-related requests and responses
 */

import axios, { AxiosInstance, AxiosResponse } from 'axios';

// Types
interface ManiceRequest {
  instruction: string;
  context?: any;
  force_model?: string;
  stream?: boolean;
}

interface ExcelOperation {
  type: string;
  target: string;
  value?: any;
  options?: any;
}

interface ManiceResponse {
  action: string;
  explanation: string;
  excel_operations: ExcelOperation[];
  parameters?: any;
  model_info?: any;
  timestamp: number;
}

interface ServerHealth {
  status: string;
  version: string;
  uptime: number;
  providers: any;
  models: any;
  timestamp: number;
}

export class AIService {
  private client: AxiosInstance;
  private serverUrl: string;
  private isConnected: boolean = false;
  private connectionRetries: number = 0;
  private maxRetries: number = 3;

  constructor(serverUrl: string = 'http://127.0.0.1:8899') {
    this.serverUrl = serverUrl;
    
    this.client = axios.create({
      baseURL: this.serverUrl,
      timeout: 60000, // 60 second timeout
      headers: {
        'Content-Type': 'application/json',
        'User-Agent': 'Manice-Excel-Addin/1.0.0'
      }
    });

    // Add request interceptor for logging
    this.client.interceptors.request.use(
      (config) => {
        console.log(`AI Service: ${config.method?.toUpperCase()} ${config.url}`);
        return config;
      },
      (error) => {
        console.error('AI Service request error:', error);
        return Promise.reject(error);
      }
    );

    // Add response interceptor for error handling
    this.client.interceptors.response.use(
      (response) => {
        this.isConnected = true;
        this.connectionRetries = 0;
        return response;
      },
      (error) => {
        this.handleConnectionError(error);
        return Promise.reject(error);
      }
    );

    // Check connection on initialization
    this.checkConnection();
  }

  /**
   * Process an AI instruction
   * @param request The Manice request object
   * @returns Promise resolving to Manice response
   */
  async processInstruction(request: ManiceRequest): Promise<ManiceResponse> {
    try {
      // Ensure connection is available
      await this.ensureConnection();

      // Make the request
      const response: AxiosResponse<ManiceResponse> = await this.client.post('/manice', request);
      
      if (!response.data) {
        throw new Error('Empty response from AI server');
      }

      return response.data;

    } catch (error) {
      console.error('AI Service: Error processing instruction:', error);
      
      // Return fallback response
      return this.createFallbackResponse(error.message, request.instruction);
    }
  }

  /**
   * Process streaming AI instruction (for future use)
   * @param request The Manice request object
   * @returns Promise resolving to streaming response
   */
  async processStreamingInstruction(request: ManiceRequest): Promise<AsyncIterable<string>> {
    try {
      await this.ensureConnection();

      // Note: This is prepared for future streaming support
      // Excel Custom Functions don't currently support streaming
      const response = await this.client.post('/manice/stream', {
        ...request,
        stream: true
      }, {
        responseType: 'stream'
      });

      // Convert stream to async iterable
      return this.parseServerSentEvents(response.data);

    } catch (error) {
      console.error('AI Service: Error processing streaming instruction:', error);
      throw error;
    }
  }

  /**
   * Check AI server health
   * @returns Promise resolving to health status
   */
  async checkHealth(): Promise<ServerHealth> {
    try {
      const response: AxiosResponse<ServerHealth> = await this.client.get('/health');
      return response.data;
    } catch (error) {
      console.error('AI Service: Health check failed:', error);
      throw new Error(`AI server health check failed: ${error.message}`);
    }
  }

  /**
   * Get server statistics
   * @returns Promise resolving to server stats
   */
  async getServerStats(): Promise<any> {
    try {
      const response = await this.client.get('/stats');
      return response.data;
    } catch (error) {
      console.error('AI Service: Failed to get server stats:', error);
      throw error;
    }
  }

  /**
   * Check if AI server is connected and available
   * @returns Promise resolving to connection status
   */
  async checkConnection(): Promise<boolean> {
    try {
      const response = await this.client.get('/', { timeout: 5000 });
      this.isConnected = response.status === 200;
      return this.isConnected;
    } catch (error) {
      this.isConnected = false;
      console.warn('AI Service: Connection check failed:', error.message);
      return false;
    }
  }

  /**
   * Get current connection status
   * @returns Boolean indicating if connected to AI server
   */
  isServerConnected(): boolean {
    return this.isConnected;
  }

  /**
   * Ensure connection to AI server exists
   * @private
   */
  private async ensureConnection(): Promise<void> {
    if (!this.isConnected) {
      const connected = await this.checkConnection();
      
      if (!connected) {
        if (this.connectionRetries < this.maxRetries) {
          this.connectionRetries++;
          console.warn(`AI Service: Retrying connection (${this.connectionRetries}/${this.maxRetries})`);
          
          // Wait before retry
          await new Promise(resolve => setTimeout(resolve, 1000 * this.connectionRetries));
          
          return this.ensureConnection();
        } else {
          throw new Error('Cannot connect to Manice AI server. Please ensure the server is running.');
        }
      }
    }
  }

  /**
   * Handle connection errors
   * @private
   */
  private handleConnectionError(error: any): void {
    if (error.code === 'ECONNREFUSED' || error.code === 'ETIMEDOUT') {
      this.isConnected = false;
      console.error('AI Service: Connection to server lost');
    }
  }

  /**
   * Create fallback response when AI server is unavailable
   * @private
   */
  private createFallbackResponse(errorMessage: string, instruction: string): ManiceResponse {
    return {
      action: 'error',
      explanation: `AI server unavailable: ${errorMessage}. Please check if the Manice AI server is running.`,
      excel_operations: [],
      parameters: {
        fallback: true,
        original_instruction: instruction
      },
      model_info: {
        error: true,
        fallback: true
      },
      timestamp: Date.now()
    };
  }

  /**
   * Parse Server-Sent Events stream
   * @private
   */
  private async* parseServerSentEvents(stream: any): AsyncIterable<string> {
    // This would be implemented for streaming support
    // Currently not used as Excel Custom Functions don't support streaming
    
    const reader = stream.getReader();
    const decoder = new TextDecoder();
    let buffer = '';

    try {
      while (true) {
        const { done, value } = await reader.read();
        
        if (done) break;
        
        buffer += decoder.decode(value, { stream: true });
        const lines = buffer.split('\n');
        buffer = lines.pop() || '';

        for (const line of lines) {
          if (line.startsWith('data: ')) {
            const data = line.slice(6);
            if (data === '[DONE]') return;
            
            try {
              const parsed = JSON.parse(data);
              if (parsed.chunk) {
                yield parsed.chunk;
              }
            } catch (e) {
              console.warn('Failed to parse SSE data:', data);
            }
          }
        }
      }
    } finally {
      reader.releaseLock();
    }
  }

  /**
   * Update server URL
   * @param newUrl New server URL
   */
  updateServerUrl(newUrl: string): void {
    this.serverUrl = newUrl;
    this.client.defaults.baseURL = newUrl;
    this.isConnected = false;
    this.connectionRetries = 0;
    
    // Check new connection
    this.checkConnection();
  }

  /**
   * Get current server URL
   * @returns Current server URL
   */
  getServerUrl(): string {
    return this.serverUrl;
  }
}