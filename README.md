# Manice - Excel AI CoPilot Extension

> **Fully Offline, Real-Time Excel AI Assistant**  
> Powered by DeepSeek R1 + Mistral-7B | 100% Local | Zero External APIs

## 🚀 Overview

Manice is a comprehensive Microsoft Excel AI CoPilot that seamlessly integrates advanced AI capabilities directly into your spreadsheets. Unlike cloud-based solutions, Manice runs entirely offline using local AI models, ensuring your data never leaves your machine.

## ✨ Key Features

### 🔹 Dual AI Model Architecture
- **Large Model**: DeepSeek R1 for complex reasoning, analysis, and heavy computations
- **Small Model**: Mistral-7B for quick formulas, formatting, and lightweight tasks
- **Smart Routing**: Automatically selects the optimal model for each request

### 🔹 Real-Time Excel Integration
- **Formula Function**: `=Manice("your instruction")` for direct cell integration
- **Interactive Sidebar**: Conversational chat UI for complex workflows
- **Live Sheet Control**: Read, modify, and format spreadsheets in real-time
- **Undo/Redo Support**: Safe execution with rollback capabilities

### 🔹 Advanced Capabilities
- **Data Manipulation**: Add/delete rows/columns, apply formulas, auto-fill patterns
- **Smart Formatting**: Conditional formatting, styling, cell merging
- **Interactive Charts**: Dynamic visualizations with filters and slicers  
- **Business Intelligence**: Forecasting, KPI tracking, automated reporting
- **Formula AI**: Generate, debug, and explain complex Excel formulas
- **VBA Automation**: Convert natural language to Excel macros

## 🏗️ Architecture

```
Excel Add-in ←→ AI Server ←→ [DeepSeek R1 | Mistral-7B]
     ↓              ↓                    ↓
Ribbon UI      Model Router         Local Models
Sidebar UI     API Endpoints        (Ollama/LM Studio)
=Manice()      Request Handler      
```

## 📁 Project Structure

```
Manice-Excel-AI-Copilot/
├── ai-server/          # Local AI server with model routing
├── excel-addin/        # Excel add-in (Ribbon + Sidebar + Function)  
├── models-config/      # AI model configuration and setup
├── shared-utils/       # Common utilities and types
├── tests/             # Comprehensive test suite
└── docs/              # Documentation and guides
```

## 🛠️ Installation

### Prerequisites
- Microsoft Excel 2016+ (Windows)
- Node.js 18+
- Python 3.8+
- Ollama, LM Studio, or Jan for local AI models

### Quick Start
```bash
# 1. Clone and setup
git clone <repository-url>
cd Manice-Excel-AI-Copilot

# 2. Install dependencies
npm install
pip install -r requirements.txt

# 3. Configure AI models
cd models-config
./setup-models.bat

# 4. Start AI server
cd ../ai-server
python server.py

# 5. Install Excel add-in
cd ../excel-addin
npm run build
npm run install-addin
```

## 🎯 Usage Examples

### Formula Function
```excel
=Manice("Calculate compound interest for 5 years at 7% on $10,000")
=Manice("Highlight rows where sales > average in green")
=Manice("Create pivot table from A1:E100 grouped by region")
```

### Sidebar Commands
- "Insert a new column between B and C named 'Profit Margin'"
- "Apply formula =A2*0.1 down column D for rows with data"
- "Create interactive chart for revenue trends with month filter"
- "Delete all empty rows and format as professional table"

## 🔧 Development

### Running Tests
```bash
npm test                # Unit tests
npm run test:integration  # Integration tests  
npm run test:excel      # Excel-specific tests
```

### Building
```bash
npm run build:server    # AI server
npm run build:addin     # Excel add-in
npm run build:all       # Complete build
```

## 📊 Performance Benchmarks

- **Small Model (Mistral-7B)**: ~500ms response time
- **Large Model (DeepSeek R1)**: ~2-5s response time  
- **Excel Operations**: <100ms for most sheet modifications
- **Memory Usage**: ~2-4GB (models loaded)

## 🛡️ Security & Privacy

- **100% Offline**: No data transmitted to external servers
- **Local Processing**: All AI inference happens on your machine
- **Secure Integration**: Uses official Excel APIs with proper authentication
- **Safe Execution**: Undo/redo support with confirmation for destructive operations

## 📚 Documentation

- [Installation Guide](docs/installation.md)
- [User Manual](docs/user-guide.md)
- [API Reference](docs/api-reference.md)
- [Developer Guide](docs/developer-guide.md)
- [Troubleshooting](docs/troubleshooting.md)

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## 📜 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- DeepSeek for the R1 reasoning model
- Mistral AI for the efficient 7B model
- Microsoft for Excel API and development tools
- The open-source AI community

---

**Made with ❤️ for Excel power users who value privacy and performance**