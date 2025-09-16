# Manice Installation Guide

Complete setup instructions for the Manice Excel AI CoPilot.

## 🎯 Quick Start

### Prerequisites

Ensure you have these installed:

- **Microsoft Excel 2016+** (Windows)
- **Node.js 18+** ([Download](https://nodejs.org/))
- **Python 3.8+** ([Download](https://python.org/))
- **Git** ([Download](https://git-scm.com/))

### 1. Clone and Setup Project

```bash
# Clone the repository
git clone <repository-url>
cd Manice-Excel-AI-Copilot

# Install root dependencies
npm install

# Install all project dependencies
npm run install:all
```

### 2. Setup AI Models

```bash
# Run the interactive model setup
cd models-config
setup-models.bat
```

Choose option 1 to install Ollama and download the AI models automatically.

**Note**: Model downloads are large (4-20GB each) and may take 30-60 minutes.

### 3. Start AI Server

```bash
# Start the Manice AI server
cd ai-server
python server.py
```

The server will start on `http://127.0.0.1:8899`

### 4. Build and Install Excel Add-in

```bash
# Build the add-in
cd excel-addin
npm run build

# Install in Excel
npm run install-addin
```

### 5. Open Excel

1. Open Microsoft Excel
2. Look for the **"Manice AI"** tab in the ribbon
3. Click **"Open Manice"** to start the sidebar
4. Try typing `=Manice("your instruction")` in any cell!

## 📋 Detailed Setup Instructions

### AI Model Setup Options

#### Option A: Ollama (Recommended)

Ollama is the easiest way to run local AI models:

```bash
# Download and install Ollama
# Visit: https://ollama.ai

# Or use our setup script
cd models-config
setup-models.bat
# Choose option 1
```

After installation:
```bash
# Download DeepSeek R1 (Large model)
ollama pull deepseek-r1

# Download Mistral-7B (Small model) 
ollama pull mistral:7b

# Verify installation
ollama list
```

#### Option B: LM Studio

For a GUI-based model management:

1. Download LM Studio from [lmstudio.ai](https://lmstudio.ai)
2. Install and launch LM Studio
3. Search and download:
   - `deepseek-ai/deepseek-r1`
   - `mistralai/mistral-7b-instruct`
4. Start the local server (usually port 1234)
5. Update configuration:
   ```python
   # In ai-server/config.py
   preferred_provider = "lm_studio"
   lm_studio_url = "http://127.0.0.1:1234"
   ```

#### Option C: Jan

Alternative local model runner:

1. Download Jan from [jan.ai](https://jan.ai)
2. Install and configure models
3. Start server on port 1337
4. Update config to use Jan

### Development Setup

For developers wanting to modify Manice:

```bash
# Install development dependencies
npm install

# Start development servers
npm run start:dev
# This starts both AI server and Excel add-in dev server

# In separate terminals:

# Terminal 1: AI Server
cd ai-server
python server.py

# Terminal 2: Excel Add-in
cd excel-addin
npm start
```

### Excel Add-in Development

```bash
cd excel-addin

# Development mode with hot reload
npm start

# Build for production
npm run build

# Validate manifest
npm run validate

# Install/uninstall for testing
npm run install-addin
npm run uninstall-addin
```

## 🔧 Configuration

### AI Server Configuration

Edit `ai-server/config.py` to customize:

```python
# Server settings
host = "127.0.0.1"
port = 8899

# Model provider
preferred_provider = "ollama"  # or "lm_studio", "jan"

# Model routing
complexity_threshold = 0.7
large_model_keywords = ["analyze", "forecast", "complex"]
small_model_keywords = ["format", "calculate", "quick"]
```

### Excel Add-in Configuration

The add-in automatically detects the AI server. You can configure:

- Server URL (default: `http://127.0.0.1:8899`)
- Model preferences (auto, large, small)
- Safety settings (confirmations, undo)

Access settings via the **"Settings"** button in the Manice ribbon.

## 🧪 Testing Installation

### 1. Test AI Models

```bash
# Test Ollama
curl http://localhost:11434/api/tags

# Test model response
ollama run mistral:7b "Hello, test message"
```

### 2. Test AI Server

```bash
# Test server health
curl http://localhost:8899/health

# Test AI endpoint
curl -X POST http://localhost:8899/manice \
  -H "Content-Type: application/json" \
  -d '{"instruction": "Test connection", "context": null}'
```

### 3. Test Excel Integration

1. Open Excel
2. Look for **"Manice AI"** tab
3. Try these:
   - Click **"Open Manice"** → Should open sidebar
   - Type `=Manice("Hello")` in cell → Should return AI response
   - Select data and click **"Analyze Data"** → Should analyze selection

## 🛠️ Troubleshooting

### Common Issues

#### "AI server is not connected"

**Solution:**
```bash
# Check if server is running
curl http://localhost:8899/health

# Start the server
cd ai-server
python server.py
```

#### "Models not found" 

**Solution:**
```bash
# Check if models are downloaded
ollama list

# Download missing models
ollama pull deepseek-r1
ollama pull mistral:7b
```

#### "Excel add-in not appearing"

**Solutions:**
1. Check if add-in is installed:
   - Excel → File → Options → Add-ins → Manage: COM Add-ins
2. Reinstall the add-in:
   ```bash
   cd excel-addin
   npm run install-addin
   ```
3. Check manifest validation:
   ```bash
   npm run validate
   ```

#### "Custom functions not working"

**Solutions:**
1. Enable custom functions in Excel:
   - File → Options → Trust Center → Trust Center Settings → Add-ins
   - Enable "Require Application Add-ins to be signed by Trusted Publisher"
2. Clear Excel cache:
   - Close Excel completely
   - Clear cache: `%localappdata%\Microsoft\Office\16.0\Wef\`
3. Reinstall add-in

#### Performance Issues

**Memory:**
- Large model needs 8-16GB RAM
- Close other applications
- Use quantized models if needed

**Response Time:**
- First request is slower (model loading)
- Subsequent requests should be faster
- Use small model for quick operations

### Debug Mode

Enable debug mode for detailed logging:

**AI Server:**
```bash
# Set environment variable
MANICE_ENV=development python server.py
```

**Excel Add-in:**
```bash
# Development build
npm run build:dev
```

### Log Locations

- **AI Server**: Console output and server logs
- **Excel Add-in**: Browser console (F12 in task pane)
- **Office.js**: Excel Developer Console

## 🚀 Production Deployment

### Build for Production

```bash
# Build AI server
cd ai-server
pip install -r requirements.txt

# Build Excel add-in
cd excel-addin
npm run build

# The built files will be in excel-addin/dist/
```

### Deploy to Server

For enterprise deployment:

1. **Host the add-in files** on your web server
2. **Update manifest.xml** with production URLs
3. **Deploy via Office Admin Center** or SharePoint
4. **Configure AI server** on internal infrastructure

### Self-Signed Certificates

For HTTPS development:

```bash
cd excel-addin
npm run generate-certs
```

Or use mkcert:
```bash
npm install -g mkcert
mkcert create-ca
mkcert create-cert
```

## 🔒 Security & Privacy

### Data Privacy
- **100% Local**: All AI processing happens on your machine
- **No Cloud**: No data sent to external servers
- **Excel Data**: Never leaves your computer

### Security Best Practices
- Keep AI models updated
- Run server on localhost only
- Use HTTPS in production
- Regular security updates

## 📞 Support

### Getting Help

1. **Documentation**: Check the [User Guide](user-guide.md)
2. **Issues**: [GitHub Issues](https://github.com/manice-ai/excel-copilot/issues)
3. **Community**: [Discussions](https://github.com/manice-ai/excel-copilot/discussions)

### Before Reporting Issues

Please include:
- Operating System version
- Excel version
- Node.js version
- Python version
- Error messages
- Steps to reproduce

### Logs to Include

```bash
# AI Server logs
cd ai-server
python server.py 2>&1 | tee server.log

# Excel add-in logs (from browser console)
# Open F12 in Excel task pane
```

---

## ✅ Installation Checklist

- [ ] Node.js 18+ installed
- [ ] Python 3.8+ installed  
- [ ] Project cloned and dependencies installed
- [ ] AI models downloaded (Ollama or LM Studio)
- [ ] AI server starts successfully
- [ ] Excel add-in built and installed
- [ ] "Manice AI" tab appears in Excel
- [ ] =Manice() function works in cells
- [ ] Sidebar opens and connects to AI server
- [ ] Quick actions work from ribbon

**🎉 Congratulations! Manice is ready to supercharge your Excel experience!**