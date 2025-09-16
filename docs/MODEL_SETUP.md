# AI Model Setup Guide

This guide will help you set up local AI models for the Manice Excel AI Copilot.

## Supported Providers

### 1. Ollama (Recommended)
- **Easy installation and management**
- **Automatic model downloads**
- **Good performance**
- **Cross-platform support**

### 2. LM Studio
- **User-friendly GUI**
- **Wide model selection**
- **Good for beginners**
- **Manual model management**

## Quick Setup

### Automated Setup
Run the automated setup script:

```bash
# Setup both providers with default models
python scripts/setup_models.py

# Setup only Ollama with specific models
python scripts/setup_models.py --provider ollama --models deepseek-r1 mistral-7b

# Setup only LM Studio
python scripts/setup_models.py --provider lmstudio

# Test existing setup
python scripts/setup_models.py --test
```

### Manual Setup

#### Option 1: Ollama Setup

1. **Install Ollama**
   ```bash
   # Windows: Download from https://ollama.com/download
   # macOS:
   brew install ollama
   # Linux:
   curl -fsSL https://ollama.com/install.sh | sh
   ```

2. **Start Ollama**
   ```bash
   ollama serve
   ```

3. **Install Models**
   ```bash
   # Install DeepSeek R1 (for complex reasoning)
   ollama pull deepseek-r1:latest
   
   # Install Mistral 7B (for quick tasks)
   ollama pull mistral:7b-instruct
   ```

4. **Test Installation**
   ```bash
   ollama list
   ollama run deepseek-r1:latest "What is 2+2?"
   ```

#### Option 2: LM Studio Setup

1. **Download and Install**
   - Visit https://lmstudio.ai/
   - Download for your platform
   - Install and launch

2. **Enable Local Server**
   - Go to Settings/Preferences
   - Enable "Local Server"
   - Set port to 1234 (default)

3. **Download Models**
   - Click "Discover" tab
   - Search for:
     - `deepseek-r1` (reasoning tasks)
     - `Mistral-7B-Instruct` (general tasks)
   - Download desired models

4. **Start Local Server**
   - Go to "Local Server" tab
   - Load a model
   - Start the server

## Recommended Models

### For Complex Tasks (DeepSeek R1)
- **Best for:** Formula generation, VBA code, complex analysis
- **Size:** ~14GB
- **Memory:** 16GB+ RAM recommended

### For Simple Tasks (Mistral 7B)
- **Best for:** Quick responses, formatting, simple queries
- **Size:** ~4GB
- **Memory:** 8GB+ RAM recommended

## Configuration

After setup, the system will automatically create a configuration file at:
`ai-server/config/models.json`

### Manual Configuration

You can manually edit the configuration:

```json
{
  "providers": {
    "ollama": {
      "url": "http://localhost:11434",
      "available": true,
      "models": ["deepseek-r1:latest", "mistral:7b-instruct"]
    }
  },
  "routing": {
    "complex_tasks": ["deepseek-r1:latest"],
    "simple_tasks": ["mistral:7b-instruct"],
    "fallback": "mistral:7b-instruct"
  }
}
```

## Starting Services

### Windows
Run `scripts/start_models.bat`

### macOS/Linux
Run `scripts/start_models.sh`

### Manual Start
```bash
# Start Ollama
ollama serve

# Start AI Server
cd ai-server
python main.py
```

## Troubleshooting

### Ollama Issues

**Service Not Starting:**
```bash
# Kill existing processes
pkill ollama

# Restart service
ollama serve
```

**Model Download Fails:**
```bash
# Check internet connection
# Try specific model versions
ollama pull mistral:7b-instruct-v0.2
```

**Permission Issues (Linux/macOS):**
```bash
sudo chown -R $USER ~/.ollama
```

### LM Studio Issues

**Server Not Responding:**
1. Check if server is started in LM Studio
2. Verify port 1234 is not blocked
3. Restart LM Studio

**Model Loading Fails:**
1. Check available disk space
2. Verify RAM requirements
3. Try smaller models first

### General Issues

**High Memory Usage:**
- Use smaller models (7B instead of 13B+)
- Close other applications
- Increase virtual memory

**Slow Responses:**
- Check CPU usage
- Use GPU acceleration if available
- Try different models

**Connection Errors:**
```bash
# Check if services are running
curl http://localhost:11434/api/tags  # Ollama
curl http://localhost:1234/v1/models  # LM Studio
```

## Performance Optimization

### System Requirements

**Minimum:**
- 8GB RAM
- 4-core CPU
- 20GB free disk space

**Recommended:**
- 16GB+ RAM
- 8-core CPU
- 50GB+ free disk space
- GPU with 8GB+ VRAM (optional)

### GPU Acceleration

**Ollama with GPU:**
```bash
# NVIDIA GPU
ollama run deepseek-r1 --gpu

# Check GPU usage
nvidia-smi
```

**LM Studio with GPU:**
- Enable GPU acceleration in settings
- Select appropriate GPU layers

### Memory Management

**Reduce Memory Usage:**
```bash
# Use quantized models
ollama pull mistral:7b-instruct-q4_0

# Limit concurrent requests in config
"max_concurrent_requests": 2
```

## Testing Your Setup

Run the comprehensive test:

```bash
python scripts/setup_models.py --test
```

Or test manually:

```bash
# Test Ollama
curl -X POST http://localhost:11434/api/generate \
  -H "Content-Type: application/json" \
  -d '{"model": "mistral:7b-instruct", "prompt": "Hello!", "stream": false}'

# Test LM Studio
curl -X POST http://localhost:1234/v1/chat/completions \
  -H "Content-Type: application/json" \
  -d '{"model": "local-model", "messages": [{"role": "user", "content": "Hello!"}]}'
```

## Security Considerations

### Local Network Only
- Models run locally on your machine
- No data sent to external servers
- Network access limited to localhost

### Firewall Configuration
```bash
# Allow only local connections
# Block external access to ports 11434, 1234
```

### Model Verification
- Download models from official sources
- Verify checksums when available
- Use trusted model repositories

## Next Steps

After successful setup:

1. **Test the Excel Add-in:** Load the add-in and try the =Manice() function
2. **Explore Features:** Try formula generation, VBA creation, and chat
3. **Optimize Performance:** Adjust model routing based on your usage
4. **Monitor Resources:** Keep an eye on CPU and memory usage

## Support

If you encounter issues:

1. Check the logs in `ai-server/logs/`
2. Run diagnostics: `python scripts/setup_models.py --test`
3. Consult the troubleshooting section above
4. Open an issue with detailed error information

## Model Alternatives

### Other Compatible Models

**Ollama:**
- `llama2:13b-chat` - General purpose
- `codellama:7b-instruct` - Code-focused
- `phi:latest` - Lightweight option

**LM Studio:**
- Search for "Mistral", "Llama", "CodeLlama"
- Filter by size and capabilities
- Read model descriptions for use cases