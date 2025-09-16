#!/usr/bin/env python3
"""
Manice AI Models Setup - Professional Edition
Advanced model setup with custom directory support and top-class UX
Author: Manice Development Team
Version: 2.0
"""

import os
import sys
from pathlib import Path
"""
Model Configuration Setup Scripts for Manice Excel AI Copilot
Automated setup for DeepSeek R1 and Mistral-7B with Ollama and LM Studio
"""

import os
import sys
import json
import subprocess
import requests
import time
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from loguru import logger
import argparse
import platform

class ModelProvider:
    """Base class for model providers"""
    
    def __init__(self, name: str, base_url: str):
        self.name = name
        self.base_url = base_url
        
    def is_available(self) -> bool:
        """Check if provider is available"""
        raise NotImplementedError
        
    def install_model(self, model_name: str) -> bool:
        """Install model"""
        raise NotImplementedError
        
    def list_models(self) -> List[str]:
        """List available models"""
        raise NotImplementedError

class OllamaProvider(ModelProvider):
    """Ollama model provider"""
    
    def __init__(self):
        super().__init__("Ollama", "http://localhost:11434")
        
    def is_available(self) -> bool:
        """Check if Ollama is running"""
        try:
            response = requests.get(f"{self.base_url}/api/tags", timeout=5)
            return response.status_code == 200
        except Exception:
            return False
    
    def install_ollama(self) -> bool:
        """Install Ollama"""
        logger.info("Installing Ollama...")
        
        system = platform.system()
        try:
            if system == "Windows":
                # Download and run Ollama installer for Windows
                installer_url = "https://ollama.com/download/OllamaSetup.exe"
                logger.info("Please download and install Ollama from: https://ollama.com/download")
                logger.info("After installation, restart this script.")
                return False
            elif system == "Darwin":  # macOS
                # Use Homebrew if available
                subprocess.run(["brew", "install", "ollama"], check=True)
            elif system == "Linux":
                # Use the official install script
                subprocess.run([
                    "curl", "-fsSL", "https://ollama.com/install.sh", "|", "sh"
                ], shell=True, check=True)
            
            logger.info("Ollama installed successfully!")
            return True
        except subprocess.CalledProcessError as e:
            logger.error(f"Failed to install Ollama: {e}")
            return False
    
    def start_ollama(self) -> bool:
        """Start Ollama service"""
        try:
            if platform.system() == "Windows":
                # On Windows, Ollama typically runs as a service
                subprocess.Popen(["ollama", "serve"], shell=True)
            else:
                # On Unix systems
                subprocess.Popen(["ollama", "serve"])
            
            # Wait for service to start
            for _ in range(30):
                if self.is_available():
                    return True
                time.sleep(1)
            
            return False
        except Exception as e:
            logger.error(f"Failed to start Ollama: {e}")
            return False
    
    def install_model(self, model_name: str) -> bool:
        """Install model via Ollama"""
        try:
            logger.info(f"Installing {model_name} via Ollama...")
            
            # Map model names to Ollama names
            ollama_models = {
                "deepseek-r1": "deepseek-r1:latest",
                "mistral-7b": "mistral:7b-instruct",
                "mistral": "mistral:latest"
            }
            
            ollama_name = ollama_models.get(model_name.lower(), model_name)
            
            result = subprocess.run([
                "ollama", "pull", ollama_name
            ], capture_output=True, text=True, timeout=3600)  # 1 hour timeout
            
            if result.returncode == 0:
                logger.info(f"Successfully installed {model_name}")
                return True
            else:
                logger.error(f"Failed to install {model_name}: {result.stderr}")
                return False
                
        except subprocess.TimeoutExpired:
            logger.error(f"Timeout while installing {model_name}")
            return False
        except Exception as e:
            logger.error(f"Error installing {model_name}: {e}")
            return False
    
    def list_models(self) -> List[str]:
        """List installed models"""
        try:
            response = requests.get(f"{self.base_url}/api/tags")
            if response.status_code == 200:
                data = response.json()
                return [model["name"] for model in data.get("models", [])]
            return []
        except Exception:
            return []
    
    def test_model(self, model_name: str) -> bool:
        """Test if model is working"""
        try:
            test_data = {
                "model": model_name,
                "prompt": "Test prompt: What is 2+2?",
                "stream": False
            }
            
            response = requests.post(
                f"{self.base_url}/api/generate",
                json=test_data,
                timeout=30
            )
            
            return response.status_code == 200
        except Exception as e:
            logger.error(f"Model test failed: {e}")
            return False

class LMStudioProvider(ModelProvider):
    """LM Studio model provider"""
    
    def __init__(self):
        super().__init__("LM Studio", "http://localhost:1234")
        
    def is_available(self) -> bool:
        """Check if LM Studio is running"""
        try:
            response = requests.get(f"{self.base_url}/v1/models", timeout=5)
            return response.status_code == 200
        except Exception:
            return False
    
    def install_lmstudio(self) -> bool:
        """Guide user to install LM Studio"""
        logger.info("LM Studio Installation Guide:")
        logger.info("1. Download LM Studio from: https://lmstudio.ai/")
        logger.info("2. Install and launch LM Studio")
        logger.info("3. Enable the local server in LM Studio settings")
        logger.info("4. Start the local server (usually on port 1234)")
        logger.info("5. Re-run this script")
        return False
    
    def install_model(self, model_name: str) -> bool:
        """Guide for model installation in LM Studio"""
        model_suggestions = {
            "deepseek-r1": "Search for 'deepseek-r1' in LM Studio's model library",
            "mistral-7b": "Search for 'Mistral-7B-Instruct' in LM Studio's model library",
            "mistral": "Search for 'Mistral' in LM Studio's model library"
        }
        
        suggestion = model_suggestions.get(model_name.lower(), f"Search for '{model_name}' in LM Studio")
        
        logger.info(f"To install {model_name} in LM Studio:")
        logger.info("1. Open LM Studio")
        logger.info("2. Go to the 'Discover' or 'Models' tab")
        logger.info(f"3. {suggestion}")
        logger.info("4. Download the model")
        logger.info("5. Load the model in the 'Chat' or 'Local Server' tab")
        
        return False  # Manual process
    
    def list_models(self) -> List[str]:
        """List loaded models"""
        try:
            response = requests.get(f"{self.base_url}/v1/models")
            if response.status_code == 200:
                data = response.json()
                return [model["id"] for model in data.get("data", [])]
            return []
        except Exception:
            return []
    
    def test_model(self, model_name: str = None) -> bool:
        """Test if LM Studio is working"""
        try:
            test_data = {
                "model": model_name or "local-model",
                "messages": [{"role": "user", "content": "Test: What is 2+2?"}],
                "max_tokens": 50
            }
            
            response = requests.post(
                f"{self.base_url}/v1/chat/completions",
                json=test_data,
                timeout=30
            )
            
            return response.status_code == 200
        except Exception as e:
            logger.error(f"LM Studio test failed: {e}")
            return False

class ModelConfigurator:
    """Main class for model configuration"""
    
    def __init__(self):
        self.providers = {
            "ollama": OllamaProvider(),
            "lmstudio": LMStudioProvider()
        }
        self.config_path = Path(__file__).parent.parent / "ai-server" / "config" / "models.json"
        self.config_path.parent.mkdir(parents=True, exist_ok=True)
        
    def detect_providers(self) -> Dict[str, bool]:
        """Detect available providers"""
        status = {}
        for name, provider in self.providers.items():
            status[name] = provider.is_available()
            logger.info(f"{name}: {'✓ Available' if status[name] else '✗ Not available'}")
        return status
    
    def setup_ollama(self, install_models: List[str] = None) -> bool:
        """Setup Ollama with specified models"""
        ollama = self.providers["ollama"]
        
        # Check if Ollama is available
        if not ollama.is_available():
            logger.info("Ollama not detected. Attempting to install...")
            
            if not ollama.install_ollama():
                logger.error("Failed to install Ollama automatically")
                return False
            
            # Try to start Ollama
            if not ollama.start_ollama():
                logger.error("Failed to start Ollama")
                return False
        
        # Install requested models
        if install_models:
            for model in install_models:
                if not ollama.install_model(model):
                    logger.warning(f"Failed to install {model}")
                else:
                    # Test the model
                    ollama_name = self._get_ollama_model_name(model)
                    if ollama.test_model(ollama_name):
                        logger.info(f"✓ {model} is working correctly")
                    else:
                        logger.warning(f"⚠ {model} installed but not responding correctly")
        
        return True
    
    def setup_lmstudio(self) -> bool:
        """Setup LM Studio"""
        lmstudio = self.providers["lmstudio"]
        
        if not lmstudio.is_available():
            logger.info("LM Studio not detected.")
            lmstudio.install_lmstudio()
            return False
        
        # Test LM Studio
        if lmstudio.test_model():
            logger.info("✓ LM Studio is working correctly")
            return True
        else:
            logger.warning("⚠ LM Studio is running but not responding correctly")
            return False
    
    def _get_ollama_model_name(self, model: str) -> str:
        """Get Ollama model name"""
        mapping = {
            "deepseek-r1": "deepseek-r1:latest",
            "mistral-7b": "mistral:7b-instruct",
            "mistral": "mistral:latest"
        }
        return mapping.get(model.lower(), model)
    
    def generate_config(self) -> Dict:
        """Generate model configuration"""
        config = {
            "providers": {},
            "models": {},
            "routing": {
                "complex_tasks": [],
                "simple_tasks": [],
                "fallback": "none"
            }
        }
        
        # Check providers and their models
        for name, provider in self.providers.items():
            if provider.is_available():
                models = provider.list_models()
                config["providers"][name] = {
                    "url": provider.base_url,
                    "available": True,
                    "models": models
                }
                
                # Add models to config
                for model in models:
                    config["models"][model] = {
                        "provider": name,
                        "url": provider.base_url,
                        "type": self._classify_model(model)
                    }
                    
                    # Set up routing
                    if self._is_complex_model(model):
                        config["routing"]["complex_tasks"].append(model)
                    else:
                        config["routing"]["simple_tasks"].append(model)
            else:
                config["providers"][name] = {
                    "url": provider.base_url,
                    "available": False,
                    "models": []
                }
        
        # Set fallback
        if config["routing"]["simple_tasks"]:
            config["routing"]["fallback"] = config["routing"]["simple_tasks"][0]
        elif config["routing"]["complex_tasks"]:
            config["routing"]["fallback"] = config["routing"]["complex_tasks"][0]
        
        return config
    
    def _classify_model(self, model_name: str) -> str:
        """Classify model type"""
        model_lower = model_name.lower()
        
        if "deepseek" in model_lower and "r1" in model_lower:
            return "reasoning"
        elif "mistral" in model_lower:
            return "instruct"
        elif "llama" in model_lower:
            return "instruct"
        elif "qwen" in model_lower:
            return "instruct"
        else:
            return "unknown"
    
    def _is_complex_model(self, model_name: str) -> bool:
        """Determine if model is suitable for complex tasks"""
        model_lower = model_name.lower()
        
        # DeepSeek R1 is designed for complex reasoning
        if "deepseek" in model_lower and "r1" in model_lower:
            return True
        
        # Large models are generally better for complex tasks
        size_indicators = ["70b", "65b", "30b", "34b"]
        if any(size in model_lower for size in size_indicators):
            return True
        
        return False
    
    def save_config(self, config: Dict):
        """Save configuration to file"""
        try:
            with open(self.config_path, 'w') as f:
                json.dump(config, f, indent=2)
            logger.info(f"Configuration saved to {self.config_path}")
        except Exception as e:
            logger.error(f"Failed to save config: {e}")
    
    def load_config(self) -> Optional[Dict]:
        """Load existing configuration"""
        try:
            if self.config_path.exists():
                with open(self.config_path, 'r') as f:
                    return json.load(f)
        except Exception as e:
            logger.error(f"Failed to load config: {e}")
        return None
    
    def create_startup_scripts(self):
        """Create startup scripts for different platforms"""
        scripts_dir = Path(__file__).parent
        
        # Windows batch script
        windows_script = scripts_dir / "start_models.bat"
        with open(windows_script, 'w') as f:
            f.write("""@echo off
echo Starting Manice AI Models...

REM Start Ollama if available
where ollama >nul 2>nul
if %ERRORLEVEL% EQU 0 (
    echo Starting Ollama...
    start /B ollama serve
    timeout /t 5 >nul
)

REM Check for LM Studio (manual start required)
echo Please ensure LM Studio is running if you plan to use it.
echo LM Studio should be available at http://localhost:1234

echo Model services startup complete.
pause
""")
        
        # Unix shell script
        unix_script = scripts_dir / "start_models.sh"
        with open(unix_script, 'w') as f:
            f.write("""#!/bin/bash

echo "Starting Manice AI Models..."

# Start Ollama if available
if command -v ollama &> /dev/null; then
    echo "Starting Ollama..."
    ollama serve &
    sleep 5
else
    echo "Ollama not found. Please install Ollama to use local models."
fi

# Check for LM Studio (manual start required)
echo "Please ensure LM Studio is running if you plan to use it."
echo "LM Studio should be available at http://localhost:1234"

echo "Model services startup complete."
""")
        
        # Make Unix script executable
        unix_script.chmod(0o755)
        
        logger.info("Startup scripts created:")
        logger.info(f"  Windows: {windows_script}")
        logger.info(f"  Unix/Linux/macOS: {unix_script}")

def main():
    """Main setup function"""
    parser = argparse.ArgumentParser(description="Setup AI models for Manice Excel AI Copilot")
    parser.add_argument("--provider", choices=["ollama", "lmstudio", "both"], 
                       default="both", help="Which provider to setup")
    parser.add_argument("--models", nargs="+", 
                       default=["deepseek-r1", "mistral-7b"],
                       help="Models to install (for Ollama)")
    parser.add_argument("--config-only", action="store_true",
                       help="Only generate configuration, don't install")
    parser.add_argument("--test", action="store_true",
                       help="Test existing setup")
    
    args = parser.parse_args()
    
    configurator = ModelConfigurator()
    
    logger.info("🤖 Manice Excel AI Copilot - Model Setup")
    logger.info("=" * 50)
    
    if args.test:
        # Test existing setup
        logger.info("Testing existing setup...")
        status = configurator.detect_providers()
        
        for name, available in status.items():
            if available:
                provider = configurator.providers[name]
                models = provider.list_models()
                logger.info(f"{name} models: {models}")
                
                # Test a model
                if models:
                    if provider.test_model(models[0]):
                        logger.info(f"✓ {name} is working correctly")
                    else:
                        logger.warning(f"⚠ {name} has issues")
        return
    
    if not args.config_only:
        # Setup providers
        if args.provider in ["ollama", "both"]:
            logger.info("Setting up Ollama...")
            if not configurator.setup_ollama(args.models):
                logger.warning("Ollama setup had issues")
        
        if args.provider in ["lmstudio", "both"]:
            logger.info("Setting up LM Studio...")
            if not configurator.setup_lmstudio():
                logger.warning("LM Studio setup incomplete - manual steps required")
    
    # Generate configuration
    logger.info("Generating configuration...")
    config = configurator.generate_config()
    configurator.save_config(config)
    
    # Create startup scripts
    configurator.create_startup_scripts()
    
    # Summary
    logger.info("\n" + "=" * 50)
    logger.info("Setup Summary:")
    logger.info("=" * 50)
    
    for provider_name, provider_config in config["providers"].items():
        status = "✓ Available" if provider_config["available"] else "✗ Not available"
        logger.info(f"{provider_name}: {status}")
        
        if provider_config["available"] and provider_config["models"]:
            for model in provider_config["models"]:
                logger.info(f"  - {model}")
    
    logger.info(f"\nRouting Configuration:")
    logger.info(f"  Complex tasks: {config['routing']['complex_tasks']}")
    logger.info(f"  Simple tasks: {config['routing']['simple_tasks']}")
    logger.info(f"  Fallback: {config['routing']['fallback']}")
    
    if config["routing"]["fallback"] == "none":
        logger.warning("\n⚠ No models available! Please install at least one model.")
        logger.info("For Ollama: Run this script again without --config-only")
        logger.info("For LM Studio: Follow the manual installation steps")
    else:
        logger.info("\n✅ Setup complete! Your AI models are ready.")
        logger.info("You can now start the AI server with: python ai-server/main.py")

if __name__ == "__main__":
    # Configure logging
    logger.remove()
    logger.add(
        sys.stderr,
        format="<green>{time:YYYY-MM-DD HH:mm:ss}</green> | <level>{level: <8}</level> | {message}",
        level="INFO"
    )
    
    main()