/**
 * Chat Sidebar - React component for conversational AI interaction
 */

import React, { useState, useEffect, useRef } from 'react';
import {
  Stack,
  TextField,
  DefaultButton,
  PrimaryButton,
  Text,
  Spinner,
  MessageBar,
  MessageBarType,
  IconButton,
  Panel,
  PanelType,
  Separator,
  TooltipHost,
  getId
} from '@fluentui/react';
import { AIService } from '../services/aiService';
import { ExcelService } from '../services/excelService';
import './ChatSidebar.css';

interface ChatMessage {
  id: string;
  type: 'user' | 'assistant' | 'system';
  content: string;
  timestamp: Date;
  isLoading?: boolean;
  operations?: any[];
  model_info?: any;
}

interface ChatSidebarProps {
  isOpen: boolean;
  onDismiss: () => void;
}

export const ChatSidebar: React.FC<ChatSidebarProps> = ({ isOpen, onDismiss }) => {
  const [messages, setMessages] = useState<ChatMessage[]>([]);
  const [inputText, setInputText] = useState('');
  const [isLoading, setIsLoading] = useState(false);
  const [connectionStatus, setConnectionStatus] = useState<'connected' | 'disconnected' | 'checking'>('checking');
  const [aiService] = useState(new AIService());
  const [excelService] = useState(new ExcelService());
  const messagesEndRef = useRef<HTMLDivElement>(null);
  const inputRef = useRef<HTMLInputElement>(null);

  // Auto-scroll to bottom when new messages arrive
  useEffect(() => {
    messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' });
  }, [messages]);

  // Check AI server connection on mount
  useEffect(() => {
    checkConnection();
    
    // Welcome message
    if (messages.length === 0) {
      addSystemMessage("Welcome to Manice! I'm your Excel AI assistant. Ask me anything about your spreadsheet!");
    }
  }, []);

  // Focus input when sidebar opens
  useEffect(() => {
    if (isOpen && inputRef.current) {
      setTimeout(() => inputRef.current?.focus(), 100);
    }
  }, [isOpen]);

  const checkConnection = async () => {
    setConnectionStatus('checking');
    try {
      const isConnected = await aiService.checkConnection();
      setConnectionStatus(isConnected ? 'connected' : 'disconnected');
    } catch (error) {
      setConnectionStatus('disconnected');
    }
  };

  const addMessage = (message: Omit<ChatMessage, 'id' | 'timestamp'>) => {
    const newMessage: ChatMessage = {
      ...message,
      id: Date.now().toString() + Math.random().toString(36).substr(2, 9),
      timestamp: new Date()
    };
    setMessages(prev => [...prev, newMessage]);
    return newMessage.id;
  };

  const addSystemMessage = (content: string) => {
    addMessage({
      type: 'system',
      content
    });
  };

  const updateMessage = (messageId: string, updates: Partial<ChatMessage>) => {
    setMessages(prev => 
      prev.map(msg => 
        msg.id === messageId ? { ...msg, ...updates } : msg
      )
    );
  };

  const handleSendMessage = async () => {
    if (!inputText.trim() || isLoading) return;

    const userMessage = inputText.trim();
    setInputText('');
    setIsLoading(true);

    // Add user message
    addMessage({
      type: 'user',
      content: userMessage
    });

    // Add loading assistant message
    const assistantMessageId = addMessage({
      type: 'assistant',
      content: 'Thinking...',
      isLoading: true
    });

    try {
      // Get current Excel context
      const context = await excelService.getCurrentContext();

      // Send to AI service
      const response = await aiService.processInstruction({
        instruction: userMessage,
        context: context,
        stream: false
      });

      // Update assistant message with response
      updateMessage(assistantMessageId, {
        content: response.explanation,
        isLoading: false,
        operations: response.excel_operations,
        model_info: response.model_info
      });

      // Execute Excel operations if any
      if (response.excel_operations && response.excel_operations.length > 0) {
        try {
          await excelService.executeOperations(response.excel_operations);
          addSystemMessage(`✅ Applied ${response.excel_operations.length} changes to your spreadsheet.`);
        } catch (error) {
          addSystemMessage(`⚠️ Some changes couldn't be applied: ${error.message}`);
        }
      }

    } catch (error) {
      updateMessage(assistantMessageId, {
        content: `Sorry, I encountered an error: ${error.message}`,
        isLoading: false
      });
    } finally {
      setIsLoading(false);
    }
  };

  const handleQuickAction = async (action: string, description: string) => {
    setInputText(description);
    await handleSendMessage();
  };

  const handleClearChat = () => {
    setMessages([]);
    addSystemMessage("Chat cleared. How can I help you with your spreadsheet?");
  };

  const getConnectionIcon = () => {
    switch (connectionStatus) {
      case 'connected':
        return { iconName: 'PlugConnected', color: 'green' };
      case 'disconnected':
        return { iconName: 'PlugDisconnected', color: 'red' };
      case 'checking':
        return { iconName: 'Sync', color: 'orange' };
    }
  };

  const renderMessage = (message: ChatMessage) => {
    const messageClass = `chat-message chat-message-${message.type}`;
    const timeString = message.timestamp.toLocaleTimeString([], { 
      hour: '2-digit', 
      minute: '2-digit' 
    });

    return (
      <div key={message.id} className={messageClass}>
        <div className="message-header">
          <Text variant="small" className="message-sender">
            {message.type === 'user' ? 'You' : 
             message.type === 'assistant' ? 'Manice AI' : 'System'}
          </Text>
          <Text variant="tiny" className="message-time">
            {timeString}
          </Text>
        </div>
        
        <div className="message-content">
          {message.isLoading ? (
            <div className="loading-container">
              <Spinner size={1} />
              <Text variant="small">{message.content}</Text>
            </div>
          ) : (
            <Text>{message.content}</Text>
          )}
        </div>

        {message.model_info && !message.isLoading && (
          <div className="message-meta">
            <Text variant="tiny" className="model-info">
              {message.model_info.model_used} • {message.model_info.response_time?.toFixed(2)}s
            </Text>
          </div>
        )}

        {message.operations && message.operations.length > 0 && (
          <div className="operations-summary">
            <Text variant="small">
              📝 Applied {message.operations.length} operation(s)
            </Text>
          </div>
        )}
      </div>
    );
  };

  return (
    <Panel
      isOpen={isOpen}
      onDismiss={onDismiss}
      type={PanelType.medium}
      headerText="Manice AI Assistant"
      isHiddenOnDismiss={true}
      className="chat-sidebar"
    >
      <div className="chat-container">
        
        {/* Connection Status */}
        <div className="connection-status">
          <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
            <IconButton
              iconProps={{ 
                iconName: getConnectionIcon().iconName,
                style: { color: getConnectionIcon().color }
              }}
              onClick={checkConnection}
              title="Check AI server connection"
            />
            <Text variant="small">
              {connectionStatus === 'connected' ? 'AI Server Connected' : 
               connectionStatus === 'disconnected' ? 'AI Server Disconnected' : 
               'Checking Connection...'}
            </Text>
            <IconButton
              iconProps={{ iconName: 'Clear' }}
              onClick={handleClearChat}
              title="Clear chat history"
            />
          </Stack>
        </div>

        <Separator />

        {/* Quick Actions */}
        <div className="quick-actions">
          <Text variant="mediumPlus" className="section-title">Quick Actions</Text>
          <Stack tokens={{ childrenGap: 8 }}>
            <DefaultButton 
              text="Analyze Selected Data"
              onClick={() => handleQuickAction('analyze', 'Analyze the selected data and provide insights')}
              disabled={isLoading}
            />
            <DefaultButton 
              text="Create Smart Chart"
              onClick={() => handleQuickAction('chart', 'Create an appropriate chart from the selected data')}
              disabled={isLoading}
            />
            <DefaultButton 
              text="Clean Data"
              onClick={() => handleQuickAction('clean', 'Clean and format the selected data')}
              disabled={isLoading}
            />
            <DefaultButton 
              text="Generate Formula"
              onClick={() => handleQuickAction('formula', 'Help me create a formula')}
              disabled={isLoading}
            />
          </Stack>
        </div>

        <Separator />

        {/* Chat Messages */}
        <div className="chat-messages" role="log" aria-live="polite">
          {messages.map(renderMessage)}
          <div ref={messagesEndRef} />
        </div>

        {/* Connection Warning */}
        {connectionStatus === 'disconnected' && (
          <MessageBar messageBarType={MessageBarType.warning} isMultiline>
            AI server is not connected. Please make sure the Manice AI server is running on port 8899.
            <DefaultButton 
              text="Retry Connection" 
              onClick={checkConnection}
              styles={{ root: { marginLeft: 8 } }}
            />
          </MessageBar>
        )}

        {/* Input Area */}
        <div className="chat-input">
          <Stack horizontal tokens={{ childrenGap: 8 }}>
            <TextField
              ref={inputRef}
              placeholder="Ask me anything about your spreadsheet..."
              value={inputText}
              onChange={(_, newValue) => setInputText(newValue || '')}
              onKeyDown={(e) => {
                if (e.key === 'Enter' && !e.shiftKey) {
                  e.preventDefault();
                  handleSendMessage();
                }
              }}
              multiline
              autoAdjustHeight
              disabled={isLoading || connectionStatus === 'disconnected'}
              styles={{ 
                root: { flexGrow: 1 },
                fieldGroup: { minHeight: 40 }
              }}
            />
            <PrimaryButton
              iconProps={{ iconName: 'Send' }}
              onClick={handleSendMessage}
              disabled={!inputText.trim() || isLoading || connectionStatus === 'disconnected'}
              title="Send message"
            />
          </Stack>
        </div>

        {/* Examples */}
        <div className="chat-examples">
          <Text variant="small" className="examples-title">Try asking:</Text>
          <div className="examples-list">
            <Text variant="tiny" className="example-item">
              "Calculate the total sales for this quarter"
            </Text>
            <Text variant="tiny" className="example-item">
              "Highlight cells where revenue > $1000 in green"
            </Text>
            <Text variant="tiny" className="example-item">
              "Create a pivot table grouped by region"
            </Text>
            <Text variant="tiny" className="example-item">
              "What's the trend in our monthly data?"
            </Text>
          </div>
        </div>

      </div>
    </Panel>
  );
};