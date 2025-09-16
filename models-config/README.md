# Manice AI Model Configuration

This directory contains configuration and setup scripts for the AI models used by Manice Excel AI CoPilot.

## 🧠 AI Models

Manice uses a dual-model architecture for optimal performance:

### Large Model: DeepSeek R1
- **Purpose**: Complex reasoning, detailed analysis, business intelligence
- **Size**: ~20GB (quantized versions available)
- **Use Cases**: 
  - Data analysis and insights
  - Business forecasting
  - Complex formula generation
  - Trend analysis
  - Report generation

### Small Model: Mistral-7B
- **Purpose**: Quick operations, simple tasks, fast responses
- **Size**: ~4GB
- **Use Cases**:
  - Cell formatting
  - Simple calculations
  - Quick formulas
  - Data cleaning
  - Basic chart creation

## 🚀 Quick Setup

### Option 1: Ollama (Recommended)

1. Run the setup script:
   ```cmd
   setup-models.bat
   ```

2. Choose option 1 to install Ollama and download models

3. Wait for models to download (this may take 30-60 minutes)

### Option 2: Manual Ollama Setup

1. Install Ollama:
   - Download from: https://ollama.ai
   - Run the installer
   - Restart your terminal

2. Download models:
   ```bash
   ollama pull deepseek-r1
   ollama pull mistral:7b
   ```

3. Verify installation:
   ```bash
   ollama list
   ```

### Option 3: LM Studio

1. Download LM Studio from: https://lmstudio.ai

2. Install and launch LM Studio

3. Search and download these models:
   - `deepseek-ai/deepseek-r1` (or similar)
   - `mistralai/mistral-7b-instruct`

4. Start the local server (usually port 1234)

5. Update Manice configuration:
   ```python
   # In ai-server/config.py
   preferred_provider = "lm_studio"
   lm_studio_url = "http://127.0.0.1:1234"
   ```

## ⚙️ Configuration

### Model Provider Priority

Manice checks providers in this order:

1. **Ollama** (port 11434) - Default
2. **LM Studio** (port 1234) 
3. **Jan** (port 1337)

### Smart Model Selection

The AI server automatically selects the appropriate model based on:

- **Task Complexity**: Keywords and request length
- **Data Size**: Large datasets use the large model
- **User Preference**: Can be overridden in settings

Example routing rules:
```python
# Large Model Keywords
"analyze", "forecast", "explain", "why", "insights", "trends"

# Small Model Keywords  
"format", "calculate", "sum", "highlight", "quick", "simple"
```

### Performance Settings

```python
# Model response timeouts
large_model.timeout = 60  # seconds
small_model.timeout = 15  # seconds

# Memory usage
large_model.memory_gb = 12.0
small_model.memory_gb = 8.0

# Caching
model_cache_size = 3  # responses
```

## 🛠️ Troubleshooting

### Models Not Loading

1. **Check Ollama Service**:
   ```bash
   ollama serve
   ```

2. **Verify Models Exist**:
   ```bash
   ollama list
   ```

3. **Test Model Response**:
   ```bash
   ollama run mistral:7b "Hello, are you working?"
   ```

### Performance Issues

1. **Check Available RAM**: Large model needs 8-16GB
2. **Use GPU**: Enable GPU acceleration if available
3. **Try Quantized Models**: Use Q4 or Q8 quantized versions

### Alternative Models

If the recommended models don't work, try these alternatives:

**Large Model Alternatives**:
- `deepseek-coder:7b` (coding focused)
- `llama2:13b` (general purpose)
- `codellama:13b` (code generation)

**Small Model Alternatives**:
- `llama2:7b` (general purpose)
- `codellama:7b` (code generation)
- `phi:2.7b` (lightweight)

## 📊 Model Comparison

| Feature | DeepSeek R1 | Mistral-7B |
|---------|-------------|------------|
| Size | ~20GB | ~4GB |
| Response Time | 2-5s | 0.5-1s |
| Memory Usage | 8-16GB | 4-8GB |
| Best For | Analysis, Reasoning | Quick Tasks |
| Accuracy | Very High | High |

## 🔧 Advanced Configuration

### Custom Model Paths

Update `ai-server/config.py`:

```python
large_model = AIModelConfig(
    name="your-large-model",
    model_path="your-org/your-large-model",
    # ... other settings
)

small_model = AIModelConfig(
    name="your-small-model", 
    model_path="your-org/your-small-model",
    # ... other settings
)
```

### Environment Variables

```bash
# Override model provider
MANICE_PREFERRED_PROVIDER=ollama

# Override model URLs
MANICE_OLLAMA_URL=http://localhost:11434
MANICE_LM_STUDIO_URL=http://localhost:1234

# Override model selection
MANICE_FORCE_MODEL=large  # or 'small'
```

### Multiple Model Providers

You can run multiple providers simultaneously:

1. **Ollama**: Run `ollama serve` (port 11434)
2. **LM Studio**: Start server (port 1234)
3. **Jan**: Configure custom port (port 1337)

Manice will automatically use the best available provider.

## 🎯 Performance Tips

1. **Use SSD Storage**: Store models on fast storage
2. **Adequate RAM**: 16GB+ recommended for large models
3. **GPU Acceleration**: Enable CUDA if available
4. **Model Quantization**: Use Q4/Q8 for faster inference
5. **Batch Requests**: Group multiple operations together

## 📝 Model Updates

To update models:

```bash
# Update all models
ollama pull deepseek-r1
ollama pull mistral:7b

# Or use the setup script
setup-models.bat
```

Choose option 3 to download models only.

## 🔒 Privacy & Security

- **100% Local**: All models run on your machine
- **No Internet**: No data sent to external servers
- **Private Data**: Your Excel data never leaves your computer
- **Open Source**: All model code is open source

## 📞 Support

If you encounter issues:

1. Check the [Troubleshooting](#troubleshooting) section
2. Run the model test: `setup-models.bat` → Option 4
3. Check the main project README
4. Open an issue on GitHub

---

**Note**: Model downloads are large (4-20GB each). Ensure you have sufficient disk space and a stable internet connection.