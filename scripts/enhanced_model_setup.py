#!/usr/bin/env python3
"""
Manice AI Models Setup - Professional Edition
Advanced model setup with custom directory support and top-class UX
Author: Manice Development Team
Version: 2.0
"""

import os
import sys
import subprocess
import time
import json
import requests
import shutil
from pathlib import Path
from typing import List, Dict, Optional, Tuple
from dataclasses import dataclass
import logging
from datetime import datetime
import threading
import queue

# Rich library for beautiful console output
try:
    from rich.console import Console
    from rich.progress import Progress, SpinnerColumn, BarColumn, TextColumn, TimeElapsedColumn
    from rich.table import Table
    from rich.panel import Panel
    from rich.text import Text
    from rich.prompt import Confirm, Prompt
    RICH_AVAILABLE = True
except ImportError:
    RICH_AVAILABLE = False
    print("Installing rich library for better UI...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "rich"])
    from rich.console import Console
    from rich.progress import Progress, SpinnerColumn, BarColumn, TextColumn, TimeElapsedColumn
    from rich.table import Table
    from rich.panel import Panel
    from rich.text import Text
    from rich.prompt import Confirm, Prompt

console = Console()

# Configuration
MODELS_BASE_DIR = Path("D:/Open_Source_AI_Models")
OLLAMA_MODELS_DIR = MODELS_BASE_DIR / "Ollama"
LOGS_DIR = MODELS_BASE_DIR / "Logs"
BACKUP_DIR = MODELS_BASE_DIR / "Backups"

# Create directories
for dir_path in [MODELS_BASE_DIR, OLLAMA_MODELS_DIR, LOGS_DIR, BACKUP_DIR]:
    dir_path.mkdir(parents=True, exist_ok=True)

# Configure logging
log_file = LOGS_DIR / f"model_setup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_file),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

@dataclass
class ModelInfo:
    name: str
    size_gb: float
    description: str
    priority: int
    ollama_name: str
    category: str
    features: List[str]
    recommended_ram_gb: int

# Enhanced model catalog with categories
LARGE_MODELS = [
    ModelInfo("DeepSeek R1", 20.0, "State-of-the-art reasoning model with advanced capabilities", 
              1, "deepseek-r1", "reasoning", ["complex_analysis", "code_generation", "math"], 16),
    ModelInfo("CodeLlama 34B", 19.0, "Specialized large model for code generation and programming", 
              2, "codellama:34b", "programming", ["code_completion", "debugging", "refactoring"], 16),
    ModelInfo("Llama2 70B", 38.0, "Very large general-purpose model with exceptional performance", 
              3, "llama2:70b", "general", ["conversation", "analysis", "reasoning"], 32),
    ModelInfo("Mixtral 8x7B", 26.0, "Mixture of experts model for versatile applications", 
              4, "mixtral:8x7b", "general", ["multilingual", "code", "analysis"], 20),
    ModelInfo("DeepSeek Coder 33B", 18.0, "Advanced coding model with strong programming capabilities", 
              5, "deepseek-coder:33b", "programming", ["code_analysis", "bug_fixing"], 16),
]

SMALL_MODELS = [
    ModelInfo("Mistral 7B", 4.0, "Fast and efficient general-purpose model", 
              1, "mistral:7b", "general", ["quick_responses", "general_tasks"], 8),
    ModelInfo("Llama2 7B Chat", 4.0, "Optimized conversational model", 
              2, "llama2:7b-chat", "conversation", ["dialogue", "assistance"], 8),
    ModelInfo("CodeLlama 7B", 4.0, "Compact code generation model", 
              3, "codellama:7b", "programming", ["code_completion", "small_projects"], 8),
    ModelInfo("Phi-3 Mini", 2.0, "Ultra-compact efficient model", 
              4, "phi3:mini", "efficiency", ["quick_tasks", "low_resource"], 4),
]

UTILITY_MODELS = [
    ModelInfo("Phi-3 Medium", 7.0, "Balanced performance and efficiency", 
              1, "phi3:medium", "utility", ["specialized_tasks", "analysis"], 8),
    ModelInfo("Gemma 7B", 4.0, "Google's instruction-following model", 
              2, "gemma:7b", "instruction", ["task_completion", "structured_output"], 8),
    ModelInfo("Neural Chat 7B", 4.0, "Optimized for conversational AI", 
              3, "neural-chat:7b", "conversation", ["customer_support", "dialogue"], 8),
]

class ManiceModelSetup:
    def __init__(self):
        self.ollama_url = "http://127.0.0.1:11434"
        self.max_retries = 3
        self.retry_delay = 5
        self.total_downloaded_gb = 0.0
        self.max_total_gb = 35.0
        
        # Set environment variables
        os.environ["OLLAMA_MODELS"] = str(OLLAMA_MODELS_DIR)
        
        # Display banner
        self.display_banner()
    
    def display_banner(self):
        """Display professional banner"""
        banner = """
[bold blue]╔══════════════════════════════════════════════════════════════╗
║                    🤖 MANICE AI SETUP PRO                    ║
║                Professional Model Management                  ║
╚══════════════════════════════════════════════════════════════╝[/bold blue]

[cyan]📁 Models Directory:[/cyan] {models_dir}
[cyan]🗂️  Ollama Models:[/cyan] {ollama_dir}
[cyan]📋 Log Files:[/cyan] {logs_dir}
        """.format(
            models_dir=MODELS_BASE_DIR,
            ollama_dir=OLLAMA_MODELS_DIR,
            logs_dir=LOGS_DIR
        )
        
        console.print(Panel(banner, title="[bold green]Setup Information[/bold green]", expand=False))
    
    def check_system_requirements(self) -> bool:
        """Enhanced system requirements check"""
        console.print("\n[bold yellow]🔍 Checking System Requirements...[/bold yellow]")
        
        requirements = []
        
        with Progress(SpinnerColumn(), TextColumn("[progress.description]{task.description}"), console=console) as progress:
            # Check Ollama installation
            task1 = progress.add_task("Checking Ollama installation...", total=None)
            ollama_ok = self._check_ollama()
            progress.update(task1, description="✅ Ollama check complete" if ollama_ok else "❌ Ollama not found")
            requirements.append(("Ollama Installation", ollama_ok))
            
            # Check disk space
            task2 = progress.add_task("Checking disk space...", total=None)
            disk_ok, free_space = self._check_disk_space()
            progress.update(task2, description=f"✅ Disk space: {free_space:.1f} GB free" if disk_ok else f"❌ Insufficient space: {free_space:.1f} GB")
            requirements.append(("Disk Space (40GB+)", disk_ok))
            
            # Check RAM
            task3 = progress.add_task("Checking system RAM...", total=None)
            ram_ok, total_ram = self._check_ram()
            progress.update(task3, description=f"✅ RAM: {total_ram:.1f} GB" if ram_ok else f"⚠️  Limited RAM: {total_ram:.1f} GB")
            requirements.append(("System RAM (8GB+)", ram_ok))
            
            # Check internet
            task4 = progress.add_task("Checking internet connection...", total=None)
            net_ok = self._check_internet()
            progress.update(task4, description="✅ Internet connected" if net_ok else "❌ No internet connection")
            requirements.append(("Internet Connection", net_ok))
        
        # Display requirements table
        table = Table(title="System Requirements Check")
        table.add_column("Requirement", style="cyan")
        table.add_column("Status", style="bold")
        
        for req_name, status in requirements:
            status_text = "[green]✅ Pass[/green]" if status else "[red]❌ Fail[/red]"
            table.add_row(req_name, status_text)
        
        console.print(table)
        
        all_critical = all(req[1] for req in requirements[:2])  # Ollama and disk space are critical
        
        if not all_critical:
            console.print("\n[bold red]❌ Critical requirements not met![/bold red]")
            return False
        
        console.print("\n[bold green]✅ System ready for model installation![/bold green]")
        return True
    
    def _check_ollama(self) -> bool:
        """Check if Ollama is installed and running"""
        try:
            result = subprocess.run(['ollama', 'version'], capture_output=True, text=True, timeout=10)
            if result.returncode == 0:
                # Check if service is running
                try:
                    response = requests.get(f"{self.ollama_url}/api/tags", timeout=5)
                    return response.status_code == 200
                except:
                    # Try to start Ollama service
                    try:
                        subprocess.Popen(['ollama', 'serve'], creationflags=subprocess.CREATE_NEW_CONSOLE if sys.platform == "win32" else 0)
                        time.sleep(3)
                        response = requests.get(f"{self.ollama_url}/api/tags", timeout=5)
                        return response.status_code == 200
                    except:
                        return False
            return False
        except:
            return False
    
    def _check_disk_space(self) -> Tuple[bool, float]:
        """Check available disk space"""
        try:
            free_space_gb = shutil.disk_usage(str(MODELS_BASE_DIR.parent))[2] / (1024**3)
            return free_space_gb >= 40, free_space_gb
        except:
            return False, 0.0
    
    def _check_ram(self) -> Tuple[bool, float]:
        """Check system RAM"""
        try:
            import psutil
            total_ram_gb = psutil.virtual_memory().total / (1024**3)
            return total_ram_gb >= 8, total_ram_gb
        except ImportError:
            # Install psutil if not available
            subprocess.check_call([sys.executable, "-m", "pip", "install", "psutil"])
            import psutil
            total_ram_gb = psutil.virtual_memory().total / (1024**3)
            return total_ram_gb >= 8, total_ram_gb
        except:
            return True, 0.0  # Assume OK if can't check
    
    def _check_internet(self) -> bool:
        """Check internet connectivity"""
        try:
            response = requests.get("https://ollama.ai", timeout=10)
            return response.status_code == 200
        except:
            return False
    
    def interactive_model_selection(self) -> Tuple[List[ModelInfo], float]:
        """Interactive model selection with advanced UI"""
        console.print("\n[bold cyan]📋 Interactive Model Selection[/bold cyan]")
        
        selected_models = []
        estimated_size = 0.0
        
        # Large model selection
        console.print("\n[bold yellow]🧠 Select Large Model (Required):[/bold yellow]")
        large_table = Table(title="Large Models Available")
        large_table.add_column("#", style="bold")
        large_table.add_column("Model", style="cyan")
        large_table.add_column("Size", style="green")
        large_table.add_column("Description", style="white")
        large_table.add_column("Features", style="magenta")
        
        for i, model in enumerate(LARGE_MODELS, 1):
            features_str = ", ".join(model.features[:2])
            large_table.add_row(str(i), model.name, f"{model.size_gb} GB", 
                              model.description[:50] + "...", features_str)
        
        console.print(large_table)
        
        while True:
            try:
                choice = int(Prompt.ask("Choose large model", choices=[str(i) for i in range(1, len(LARGE_MODELS)+1)]))
                large_model = LARGE_MODELS[choice-1]
                selected_models.append(large_model)
                estimated_size += large_model.size_gb
                console.print(f"✅ Selected: [bold]{large_model.name}[/bold] ({large_model.size_gb} GB)")
                break
            except (ValueError, IndexError):
                console.print("[red]Invalid choice. Please try again.[/red]")
        
        # Small model selection
        console.print("\n[bold yellow]⚡ Select Small Model (Required):[/bold yellow]")
        small_table = Table(title="Small Models Available")
        small_table.add_column("#", style="bold")
        small_table.add_column("Model", style="cyan")
        small_table.add_column("Size", style="green")
        small_table.add_column("Description", style="white")
        
        for i, model in enumerate(SMALL_MODELS, 1):
            small_table.add_row(str(i), model.name, f"{model.size_gb} GB", model.description[:50] + "...")
        
        console.print(small_table)
        
        while True:
            try:
                choice = int(Prompt.ask("Choose small model", choices=[str(i) for i in range(1, len(SMALL_MODELS)+1)]))
                small_model = SMALL_MODELS[choice-1]
                selected_models.append(small_model)
                estimated_size += small_model.size_gb
                console.print(f"✅ Selected: [bold]{small_model.name}[/bold] ({small_model.size_gb} GB)")
                break
            except (ValueError, IndexError):
                console.print("[red]Invalid choice. Please try again.[/red]")
        
        # Optional utility models
        remaining_space = self.max_total_gb - estimated_size
        if remaining_space > 5:
            if Confirm.ask(f"\n🔧 Add utility model? (Remaining space: {remaining_space:.1f} GB)"):
                console.print("\n[bold yellow]🔧 Select Utility Model (Optional):[/bold yellow]")
                utility_table = Table(title="Utility Models Available")
                utility_table.add_column("#", style="bold")
                utility_table.add_column("Model", style="cyan")
                utility_table.add_column("Size", style="green")
                utility_table.add_column("Category", style="yellow")
                
                available_utility = [m for m in UTILITY_MODELS if m.size_gb <= remaining_space]
                for i, model in enumerate(available_utility, 1):
                    utility_table.add_row(str(i), model.name, f"{model.size_gb} GB", model.category)
                
                if available_utility:
                    console.print(utility_table)
                    choices = [str(i) for i in range(1, len(available_utility)+1)] + ["0"]
                    choice = int(Prompt.ask("Choose utility model (0 to skip)", choices=choices))
                    
                    if choice > 0:
                        utility_model = available_utility[choice-1]
                        selected_models.append(utility_model)
                        estimated_size += utility_model.size_gb
                        console.print(f"✅ Selected: [bold]{utility_model.name}[/bold] ({utility_model.size_gb} GB)")
        
        return selected_models, estimated_size
    
    def download_models_with_progress(self, models: List[ModelInfo]) -> bool:
        """Download models with beautiful progress tracking"""
        console.print(f"\n[bold green]🚀 Starting Download Process[/bold green]")
        console.print(f"[cyan]Total models: {len(models)}[/cyan]")
        console.print(f"[cyan]Estimated size: {sum(m.size_gb for m in models):.1f} GB[/cyan]")
        
        success_count = 0
        
        with Progress(
            SpinnerColumn(),
            TextColumn("[progress.description]{task.description}"),
            BarColumn(),
            TextColumn("[progress.percentage]{task.percentage:>3.0f}%"),
            TimeElapsedColumn(),
            console=console
        ) as progress:
            
            main_task = progress.add_task("[bold blue]Overall Progress", total=len(models))
            
            for i, model in enumerate(models, 1):
                model_task = progress.add_task(f"[cyan]{model.name}[/cyan]", total=100)
                
                console.print(f"\n[bold yellow]📥 Downloading {model.name}[/bold yellow]")
                console.print(f"[dim]Size: {model.size_gb} GB | Category: {model.category}[/dim]")
                
                if self._download_single_model_with_progress(model, progress, model_task):
                    progress.update(model_task, completed=100, description=f"[green]✅ {model.name} - Complete[/green]")
                    success_count += 1
                    console.print(f"[bold green]✅ {model.name} downloaded successfully![/bold green]")
                else:
                    progress.update(model_task, description=f"[red]❌ {model.name} - Failed[/red]")
                    console.print(f"[bold red]❌ Failed to download {model.name}[/bold red]")
                
                progress.update(main_task, advance=1)
        
        # Summary
        success_rate = (success_count / len(models)) * 100
        console.print(f"\n[bold blue]📊 Download Summary[/bold blue]")
        console.print(f"[green]✅ Successful: {success_count}/{len(models)} ({success_rate:.1f}%)[/green]")
        
        if success_count == len(models):
            console.print("[bold green]🎉 All models downloaded successfully![/bold green]")
        elif success_count >= 2:  # At least large + small
            console.print("[bold yellow]⚠️ Partial success - Core models available[/bold yellow]")
        else:
            console.print("[bold red]❌ Download failed - Please check logs[/bold red]")
        
        return success_count >= 2
    
    def _download_single_model_with_progress(self, model: ModelInfo, progress, task_id) -> bool:
        """Download single model with progress tracking"""
        for attempt in range(self.max_retries):
            try:
                process = subprocess.Popen(
                    ['ollama', 'pull', model.ollama_name],
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    text=True,
                    universal_newlines=True
                )
                
                # Simulate progress (Ollama doesn't provide detailed progress)
                progress_value = 0
                while process.poll() is None:
                    if progress_value < 90:
                        progress_value += 2
                        progress.update(task_id, completed=progress_value)
                    time.sleep(1)
                
                if process.returncode == 0:
                    progress.update(task_id, completed=100)
                    return True
                else:
                    if attempt < self.max_retries - 1:
                        console.print(f"[yellow]⏳ Retry attempt {attempt + 2}/{self.max_retries}[/yellow]")
                        time.sleep(self.retry_delay)
                    
            except Exception as e:
                logger.error(f"Error downloading {model.name}: {e}")
                if attempt < self.max_retries - 1:
                    time.sleep(self.retry_delay)
        
        return False
    
    def verify_installation(self) -> None:
        """Verify installed models with testing"""
        console.print("\n[bold cyan]🔍 Verifying Installation[/bold cyan]")
        
        try:
            response = requests.get(f"{self.ollama_url}/api/tags", timeout=10)
            if response.status_code == 200:
                models = response.json().get("models", [])
                
                if models:
                    table = Table(title="Installed Models")
                    table.add_column("Model Name", style="cyan")
                    table.add_column("Size", style="green")
                    table.add_column("Status", style="bold")
                    
                    for model in models:
                        name = model.get("name", "Unknown")
                        size = model.get("size", 0)
                        size_gb = size / (1024**3) if size else 0
                        
                        # Test model
                        if self._test_model_quick(name):
                            status = "[green]✅ Working[/green]"
                        else:
                            status = "[yellow]⚠️ Unknown[/yellow]"
                        
                        table.add_row(name, f"{size_gb:.1f} GB", status)
                    
                    console.print(table)
                else:
                    console.print("[red]❌ No models found![/red]")
            else:
                console.print("[red]❌ Cannot connect to Ollama service[/red]")
                
        except Exception as e:
            console.print(f"[red]❌ Verification failed: {e}[/red]")
    
    def _test_model_quick(self, model_name: str) -> bool:
        """Quick model functionality test"""
        try:
            payload = {
                "model": model_name,
                "prompt": "Hello",
                "stream": False
            }
            
            response = requests.post(
                f"{self.ollama_url}/api/generate",
                json=payload,
                timeout=15
            )
            
            return response.status_code == 200
        except:
            return False
    
    def create_shortcuts_and_config(self) -> None:
        """Create helpful shortcuts and configuration files"""
        console.print("\n[bold cyan]🔧 Creating Configuration Files[/bold cyan]")
        
        # Create config file for Manice
        config_content = {
            "models_directory": str(MODELS_BASE_DIR),
            "ollama_models_path": str(OLLAMA_MODELS_DIR),
            "setup_date": datetime.now().isoformat(),
            "version": "2.0"
        }
        
        config_file = MODELS_BASE_DIR / "manice_config.json"
        with open(config_file, 'w') as f:
            json.dump(config_content, f, indent=2)
        
        console.print(f"✅ Configuration saved: [cyan]{config_file}[/cyan]")
        
        # Create batch file for easy Ollama start
        if sys.platform == "win32":
            batch_content = f'''@echo off
title Ollama Server for Manice AI
echo Starting Ollama Server for Manice AI...
echo Models Directory: {OLLAMA_MODELS_DIR}
echo.
set OLLAMA_MODELS={OLLAMA_MODELS_DIR}
ollama serve
pause
'''
            batch_file = MODELS_BASE_DIR / "start_ollama.bat"
            with open(batch_file, 'w') as f:
                f.write(batch_content)
            
            console.print(f"✅ Startup script: [cyan]{batch_file}[/cyan]")

def main():
    """Main setup function with enhanced UX"""
    try:
        setup = ManiceModelSetup()
        
        # Check system requirements
        if not setup.check_system_requirements():
            console.print("\n[bold red]Setup aborted due to system requirements.[/bold red]")
            return False
        
        # Interactive model selection
        models, total_size = setup.interactive_model_selection()
        
        # Confirmation
        console.print(f"\n[bold blue]📋 Setup Summary[/bold blue]")
        summary_table = Table()
        summary_table.add_column("Model", style="cyan")
        summary_table.add_column("Size", style="green")
        summary_table.add_column("Category", style="yellow")
        
        for model in models:
            summary_table.add_row(model.name, f"{model.size_gb} GB", model.category)
        
        summary_table.add_row("[bold]Total", f"[bold]{total_size:.1f} GB", "[bold]---")
        console.print(summary_table)
        
        if not Confirm.ask(f"\n🚀 Proceed with download?"):
            console.print("[yellow]Setup cancelled by user.[/yellow]")
            return False
        
        # Download models
        success = setup.download_models_with_progress(models)
        
        if success:
            # Verify installation
            setup.verify_installation()
            
            # Create config files
            setup.create_shortcuts_and_config()
            
            # Success message
            success_panel = Panel(
                f"""[bold green]🎉 Setup Completed Successfully![/bold green]

[cyan]📁 Models Location:[/cyan] {OLLAMA_MODELS_DIR}
[cyan]📋 Log Files:[/cyan] {LOGS_DIR}
[cyan]🔧 Configuration:[/cyan] {MODELS_BASE_DIR}/manice_config.json

[bold yellow]Next Steps:[/bold yellow]
1. Start Manice AI Server: [cyan]cd ai-server && python server.py[/cyan]
2. Install Excel Add-in: [cyan]cd excel-addin && npm run install-addin[/cyan]
3. Open Excel and look for the "Manice AI" tab!

[dim]Total installation size: ~{total_size:.1f} GB[/dim]""",
                title="[bold green]Installation Complete[/bold green]",
                expand=False
            )
            console.print(success_panel)
            return True
        else:
            console.print("\n[bold red]❌ Setup failed. Please check the logs and try again.[/bold red]")
            return False
            
    except KeyboardInterrupt:
        console.print("\n\n[yellow]⏹️ Setup interrupted by user.[/yellow]")
        return False
    except Exception as e:
        console.print(f"\n[bold red]❌ Setup failed with error: {e}[/bold red]")
        logger.error(f"Setup error: {e}")
        return False

if __name__ == "__main__":
    success = main()
    input("\nPress Enter to exit...")
    sys.exit(0 if success else 1)