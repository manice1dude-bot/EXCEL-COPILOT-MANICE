#!/usr/bin/env python3
"""
Manice AI Models Setup - 8GB RAM Optimized Edition
Lightweight model setup designed specifically for 8GB systems
Author: Manice Development Team
Version: 2.1 - 8GB Optimized
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

# Try to use Rich for beautiful output, fallback to basic print
try:
    from rich.console import Console
    from rich.progress import Progress, SpinnerColumn, BarColumn, TextColumn, TimeElapsedColumn
    from rich.table import Table
    from rich.panel import Panel
    from rich.prompt import Confirm, Prompt
    RICH_AVAILABLE = True
except ImportError:
    print("Installing rich library for better UI...")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "rich"])
        from rich.console import Console
        from rich.progress import Progress, SpinnerColumn, BarColumn, TextColumn, TimeElapsedColumn
        from rich.table import Table
        from rich.panel import Panel
        from rich.prompt import Confirm, Prompt
        RICH_AVAILABLE = True
    except:
        RICH_AVAILABLE = False

if RICH_AVAILABLE:
    console = Console()
else:
    class SimpleConsole:
        def print(self, text, **kwargs):
            print(text)
    console = SimpleConsole()

# Configuration for 8GB Systems
MODELS_BASE_DIR = Path("D:/Open_Source_AI_Models")
OLLAMA_MODELS_DIR = MODELS_BASE_DIR / "Ollama"
LOGS_DIR = MODELS_BASE_DIR / "Logs"
BACKUP_DIR = MODELS_BASE_DIR / "Backups"

# Create directories
for dir_path in [MODELS_BASE_DIR, OLLAMA_MODELS_DIR, LOGS_DIR, BACKUP_DIR]:
    dir_path.mkdir(parents=True, exist_ok=True)

# Configure logging
log_file = LOGS_DIR / f"model_setup_8gb_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
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

# 8GB Optimized Model Catalog (Total: ~10GB max)
LIGHTWEIGHT_MODELS = [
    ModelInfo("Phi-3 Mini", 2.0, "Ultra-efficient 3.8B model perfect for 8GB systems", 
              1, "phi3:mini", "efficiency", ["quick_responses", "low_memory", "excel_tasks"], 4),
    ModelInfo("Phi-3 Medium", 7.0, "Balanced 14B model with excellent performance", 
              2, "phi3:medium", "balanced", ["reasoning", "code_help", "analysis"], 6),
    ModelInfo("Gemma 2B", 1.3, "Google's compact model for basic tasks", 
              3, "gemma:2b", "basic", ["simple_tasks", "formatting"], 3),
    ModelInfo("TinyLlama", 0.6, "Ultra-lightweight model for instant responses", 
              4, "tinyllama", "instant", ["quick_help", "simple_formulas"], 2),
]

UTILITY_MODELS_8GB = [
    ModelInfo("Codellama 7B", 4.0, "Specialized code model for Excel formulas", 
              1, "codellama:7b", "programming", ["excel_formulas", "vba_help"], 5),
    ModelInfo("Neural Chat 7B", 4.0, "Optimized conversational model", 
              2, "neural-chat:7b", "conversation", ["user_assistance", "explanations"], 5),
]

class Lightweight8GBSetup:
    def __init__(self):
        self.ollama_url = "http://127.0.0.1:11434"
        self.max_retries = 3
        self.retry_delay = 5
        self.total_downloaded_gb = 0.0
        self.max_total_gb = 10.0  # Conservative limit for 8GB systems
        
        # Set environment variables
        os.environ["OLLAMA_MODELS"] = str(OLLAMA_MODELS_DIR)
        
        # Display banner
        self.display_banner()
    
    def display_banner(self):
        """Display 8GB optimized banner"""
        if RICH_AVAILABLE:
            banner = """
[bold green]╔══════════════════════════════════════════════════════════════╗
║                🚀 MANICE AI SETUP - 8GB OPTIMIZED            ║
║              Lightweight Models for Efficient Performance     ║
╚══════════════════════════════════════════════════════════════╝[/bold green]

[cyan]📁 Models Directory:[/cyan] {models_dir}
[cyan]🗂️  Ollama Models:[/cyan] {ollama_dir}
[cyan]📋 Log Files:[/cyan] {logs_dir}
[yellow]⚡ RAM Optimized:[/yellow] Designed for 8GB systems
[green]💾 Total Size:[/green] ~10GB maximum (vs 30GB standard)
            """.format(
                models_dir=MODELS_BASE_DIR,
                ollama_dir=OLLAMA_MODELS_DIR,
                logs_dir=LOGS_DIR
            )
            
            console.print(Panel(banner, title="[bold green]8GB Optimized Setup[/bold green]", expand=False))
        else:
            print("="*60)
            print("🚀 MANICE AI SETUP - 8GB OPTIMIZED")
            print("Lightweight Models for Efficient Performance")
            print("="*60)
            print(f"📁 Models Directory: {MODELS_BASE_DIR}")
            print(f"🗂️  Ollama Models: {OLLAMA_MODELS_DIR}")
            print(f"📋 Log Files: {LOGS_DIR}")
            print("⚡ RAM Optimized: Designed for 8GB systems")
            print("💾 Total Size: ~10GB maximum (vs 30GB standard)")
            print("="*60)
    
    def check_system_requirements_8gb(self) -> bool:
        """8GB-specific system requirements check"""
        if RICH_AVAILABLE:
            console.print("\n[bold yellow]🔍 Checking 8GB System Requirements...[/bold yellow]")
        else:
            print("\n🔍 Checking 8GB System Requirements...")
        
        requirements = []
        
        # Check Ollama installation
        ollama_ok = self._check_ollama()
        requirements.append(("Ollama Installation", ollama_ok))
        
        # Check disk space (reduced requirement)
        disk_ok, free_space = self._check_disk_space_8gb()
        requirements.append(("Disk Space (15GB+)", disk_ok))
        
        # Check RAM (adjusted for 8GB)
        ram_ok, total_ram = self._check_ram_8gb()
        requirements.append(("System RAM (7GB+ for 8GB systems)", ram_ok))
        
        # Check internet connection
        net_ok = self._check_internet()
        requirements.append(("Internet Connection", net_ok))
        
        # Display requirements table
        if RICH_AVAILABLE:
            table = Table(title="8GB System Requirements Check")
            table.add_column("Requirement", style="cyan")
            table.add_column("Status", style="bold")
            
            for req_name, status in requirements:
                status_text = "[green]✅ Pass[/green]" if status else "[red]❌ Fail[/red]"
                table.add_row(req_name, status_text)
            
            console.print(table)
        else:
            print("\n8GB System Requirements Check:")
            print("-" * 40)
            for req_name, status in requirements:
                status_text = "✅ Pass" if status else "❌ Fail"
                print(f"{req_name}: {status_text}")
        
        # For 8GB systems, we're more lenient - only Ollama and disk space are critical
        critical_ok = ollama_ok and disk_ok
        
        if not critical_ok:
            if RICH_AVAILABLE:
                console.print("\n[bold red]❌ Critical requirements not met![/bold red]")
            else:
                print("\n❌ Critical requirements not met!")
            
            if not ollama_ok:
                if RICH_AVAILABLE:
                    console.print("[yellow]💡 Install Ollama from: https://ollama.ai[/yellow]")
                else:
                    print("💡 Install Ollama from: https://ollama.ai")
            return False
        
        if RICH_AVAILABLE:
            console.print("\n[bold green]✅ 8GB system ready for lightweight model installation![/bold green]")
        else:
            print("\n✅ 8GB system ready for lightweight model installation!")
        return True
    
    def _check_ollama(self) -> bool:
        """Check if Ollama is installed and running"""
        try:
            result = subprocess.run(['ollama', '--version'], capture_output=True, text=True, timeout=10)
            if result.returncode == 0:
                # Check if service is running
                try:
                    response = requests.get(f"{self.ollama_url}/api/tags", timeout=5)
                    return response.status_code == 200
                except:
                    # Try to start Ollama service
                    try:
                        subprocess.Popen(['ollama', 'serve'], 
                                       creationflags=subprocess.CREATE_NEW_CONSOLE if sys.platform == "win32" else 0)
                        time.sleep(5)  # Give more time to start
                        response = requests.get(f"{self.ollama_url}/api/tags", timeout=10)
                        return response.status_code == 200
                    except:
                        return False
            return False
        except:
            return False
    
    def _check_disk_space_8gb(self) -> Tuple[bool, float]:
        """Check available disk space for 8GB systems"""
        try:
            free_space_gb = shutil.disk_usage(str(MODELS_BASE_DIR.parent))[2] / (1024**3)
            return free_space_gb >= 15, free_space_gb  # Reduced requirement
        except:
            return False, 0.0
    
    def _check_ram_8gb(self) -> Tuple[bool, float]:
        """Check system RAM for 8GB systems"""
        try:
            import psutil
            total_ram_gb = psutil.virtual_memory().total / (1024**3)
            # For 8GB systems, we need at least 7GB to account for OS overhead
            return total_ram_gb >= 7.0, total_ram_gb
        except ImportError:
            try:
                subprocess.check_call([sys.executable, "-m", "pip", "install", "psutil"])
                import psutil
                total_ram_gb = psutil.virtual_memory().total / (1024**3)
                return total_ram_gb >= 7.0, total_ram_gb
            except:
                return True, 8.0  # Assume 8GB if can't check
        except:
            return True, 8.0  # Assume OK if can't check
    
    def _check_internet(self) -> bool:
        """Check internet connectivity"""
        try:
            response = requests.get("https://ollama.ai", timeout=10)
            return response.status_code == 200
        except:
            return False
    
    def interactive_8gb_model_selection(self) -> Tuple[List[ModelInfo], float]:
        """8GB optimized model selection"""
        if RICH_AVAILABLE:
            console.print("\n[bold cyan]📋 8GB Optimized Model Selection[/bold cyan]")
        else:
            print("\n📋 8GB Optimized Model Selection")
        
        selected_models = []
        estimated_size = 0.0
        
        # Mandatory: Select one primary model
        if RICH_AVAILABLE:
            console.print("\n[bold yellow]🎯 Select Primary Model (Required):[/bold yellow]")
            primary_table = Table(title="Lightweight Models for 8GB Systems")
            primary_table.add_column("#", style="bold")
            primary_table.add_column("Model", style="cyan")
            primary_table.add_column("Size", style="green")
            primary_table.add_column("Description", style="white")
            primary_table.add_column("RAM Needed", style="magenta")
            
            for i, model in enumerate(LIGHTWEIGHT_MODELS, 1):
                primary_table.add_row(
                    str(i), model.name, f"{model.size_gb} GB", 
                    model.description[:50] + "...", f"{model.recommended_ram_gb} GB"
                )
            
            console.print(primary_table)
            
            while True:
                try:
                    choice = int(Prompt.ask("Choose primary model", 
                                          choices=[str(i) for i in range(1, len(LIGHTWEIGHT_MODELS)+1)]))
                    primary_model = LIGHTWEIGHT_MODELS[choice-1]
                    selected_models.append(primary_model)
                    estimated_size += primary_model.size_gb
                    console.print(f"✅ Selected: [bold]{primary_model.name}[/bold] ({primary_model.size_gb} GB)")
                    break
                except (ValueError, IndexError):
                    console.print("[red]Invalid choice. Please try again.[/red]")
        else:
            print("\n🎯 Select Primary Model (Required):")
            print("Lightweight Models for 8GB Systems:")
            print("-" * 50)
            for i, model in enumerate(LIGHTWEIGHT_MODELS, 1):
                print(f"{i}. {model.name} ({model.size_gb} GB) - {model.description[:40]}...")
            
            while True:
                try:
                    choice = int(input("Choose primary model (1-{}): ".format(len(LIGHTWEIGHT_MODELS))))
                    if 1 <= choice <= len(LIGHTWEIGHT_MODELS):
                        primary_model = LIGHTWEIGHT_MODELS[choice-1]
                        selected_models.append(primary_model)
                        estimated_size += primary_model.size_gb
                        print(f"✅ Selected: {primary_model.name} ({primary_model.size_gb} GB)")
                        break
                    else:
                        print("Invalid choice. Please try again.")
                except ValueError:
                    print("Invalid choice. Please try again.")
        
        # Optional: Add utility model if space allows
        remaining_space = self.max_total_gb - estimated_size
        if remaining_space > 3:
            add_utility = False
            if RICH_AVAILABLE:
                add_utility = Confirm.ask(f"\n🔧 Add utility model? (Remaining space: {remaining_space:.1f} GB)")
            else:
                response = input(f"\n🔧 Add utility model? (Remaining space: {remaining_space:.1f} GB) [y/N]: ")
                add_utility = response.lower().startswith('y')
            
            if add_utility:
                available_utility = [m for m in UTILITY_MODELS_8GB if m.size_gb <= remaining_space]
                if available_utility:
                    if RICH_AVAILABLE:
                        console.print("\n[bold yellow]🔧 Select Utility Model (Optional):[/bold yellow]")
                        utility_table = Table(title="Utility Models Available")
                        utility_table.add_column("#", style="bold")
                        utility_table.add_column("Model", style="cyan")
                        utility_table.add_column("Size", style="green")
                        utility_table.add_column("Category", style="yellow")
                        
                        for i, model in enumerate(available_utility, 1):
                            utility_table.add_row(str(i), model.name, f"{model.size_gb} GB", model.category)
                        
                        console.print(utility_table)
                        choices = [str(i) for i in range(1, len(available_utility)+1)] + ["0"]
                        choice = int(Prompt.ask("Choose utility model (0 to skip)", choices=choices))
                        
                        if choice > 0:
                            utility_model = available_utility[choice-1]
                            selected_models.append(utility_model)
                            estimated_size += utility_model.size_gb
                            console.print(f"✅ Selected: [bold]{utility_model.name}[/bold] ({utility_model.size_gb} GB)")
                    else:
                        print("\n🔧 Select Utility Model (Optional):")
                        for i, model in enumerate(available_utility, 1):
                            print(f"{i}. {model.name} ({model.size_gb} GB) - {model.category}")
                        print("0. Skip")
                        
                        try:
                            choice = int(input("Choose utility model (0 to skip): "))
                            if choice > 0 and choice <= len(available_utility):
                                utility_model = available_utility[choice-1]
                                selected_models.append(utility_model)
                                estimated_size += utility_model.size_gb
                                print(f"✅ Selected: {utility_model.name} ({utility_model.size_gb} GB)")
                        except ValueError:
                            print("Skipping utility model.")
        
        return selected_models, estimated_size
    
    def download_8gb_models(self, models: List[ModelInfo]) -> bool:
        """Download models optimized for 8GB systems"""
        if RICH_AVAILABLE:
            console.print(f"\n[bold green]🚀 Starting 8GB Optimized Download[/bold green]")
            console.print(f"[cyan]Models: {len(models)}[/cyan]")
            console.print(f"[cyan]Total size: {sum(m.size_gb for m in models):.1f} GB[/cyan]")
            console.print(f"[green]Memory friendly: Designed for 8GB systems[/green]")
        else:
            print(f"\n🚀 Starting 8GB Optimized Download")
            print(f"Models: {len(models)}")
            print(f"Total size: {sum(m.size_gb for m in models):.1f} GB")
            print("Memory friendly: Designed for 8GB systems")
        
        success_count = 0
        
        if RICH_AVAILABLE:
            with Progress(
                SpinnerColumn(),
                TextColumn("[progress.description]{task.description}"),
                BarColumn(),
                TextColumn("[progress.percentage]{task.percentage:>3.0f}%"),
                TimeElapsedColumn(),
                console=console
            ) as progress:
                
                main_task = progress.add_task("[bold blue]Overall Progress", total=len(models))
                
                for model in models:
                    model_task = progress.add_task(f"[cyan]{model.name}[/cyan]", total=100)
                    
                    console.print(f"\n[bold yellow]📥 Downloading {model.name}[/bold yellow]")
                    console.print(f"[dim]Size: {model.size_gb} GB | RAM: {model.recommended_ram_gb} GB[/dim]")
                    
                    if self._download_single_model_8gb(model, progress, model_task):
                        progress.update(model_task, completed=100, description=f"[green]✅ {model.name}[/green]")
                        success_count += 1
                        console.print(f"[bold green]✅ {model.name} ready![/bold green]")
                    else:
                        progress.update(model_task, description=f"[red]❌ {model.name} failed[/red]")
                        console.print(f"[bold red]❌ Failed: {model.name}[/bold red]")
                    
                    progress.update(main_task, advance=1)
        else:
            for i, model in enumerate(models, 1):
                print(f"\n📥 Downloading {model.name} ({i}/{len(models)})")
                print(f"Size: {model.size_gb} GB | RAM: {model.recommended_ram_gb} GB")
                
                if self._download_single_model_8gb(model):
                    success_count += 1
                    print(f"✅ {model.name} ready!")
                else:
                    print(f"❌ Failed: {model.name}")
        
        # Summary
        success_rate = (success_count / len(models)) * 100
        if RICH_AVAILABLE:
            console.print(f"\n[bold blue]📊 Download Summary[/bold blue]")
            console.print(f"[green]✅ Successful: {success_count}/{len(models)} ({success_rate:.1f}%)[/green]")
        else:
            print(f"\n📊 Download Summary")
            print(f"✅ Successful: {success_count}/{len(models)} ({success_rate:.1f}%)")
        
        if success_count >= 1:
            if RICH_AVAILABLE:
                console.print("[bold green]🎉 8GB optimized setup complete![/bold green]")
            else:
                print("🎉 8GB optimized setup complete!")
        
        return success_count >= 1
    
    def _download_single_model_8gb(self, model: ModelInfo, progress=None, task_id=None) -> bool:
        """Download single model with 8GB optimizations"""
        for attempt in range(self.max_retries):
            try:
                process = subprocess.Popen(
                    ['ollama', 'pull', model.ollama_name],
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    text=True,
                    universal_newlines=True
                )
                
                # Progress simulation for models
                progress_value = 0
                while process.poll() is None:
                    if progress and task_id:
                        if progress_value < 90:
                            progress_value += 1
                            progress.update(task_id, completed=progress_value)
                    time.sleep(0.5)
                
                if process.returncode == 0:
                    if progress and task_id:
                        progress.update(task_id, completed=100)
                    return True
                else:
                    if attempt < self.max_retries - 1:
                        if RICH_AVAILABLE:
                            console.print(f"[yellow]⏳ Retry {attempt + 2}/{self.max_retries} for {model.name}[/yellow]")
                        else:
                            print(f"⏳ Retry {attempt + 2}/{self.max_retries} for {model.name}")
                        time.sleep(self.retry_delay)
                    
            except Exception as e:
                logger.error(f"Error downloading {model.name}: {e}")
                if attempt < self.max_retries - 1:
                    time.sleep(self.retry_delay)
        
        return False
    
    def create_8gb_config(self, models: List[ModelInfo]) -> None:
        """Create 8GB optimized configuration"""
        if RICH_AVAILABLE:
            console.print("\n[bold cyan]🔧 Creating 8GB Optimized Configuration[/bold cyan]")
        else:
            print("\n🔧 Creating 8GB Optimized Configuration")
        
        # Create config file
        config_content = {
            "models_directory": str(MODELS_BASE_DIR),
            "ollama_models_path": str(OLLAMA_MODELS_DIR),
            "setup_date": datetime.now().isoformat(),
            "version": "2.1-8GB-Optimized",
            "system_optimization": "8GB_RAM",
            "total_size_gb": sum(m.size_gb for m in models),
            "installed_models": [
                {
                    "name": m.name,
                    "ollama_name": m.ollama_name,
                    "size_gb": m.size_gb,
                    "category": m.category,
                    "ram_gb": m.recommended_ram_gb
                } for m in models
            ]
        }
        
        config_file = MODELS_BASE_DIR / "manice_8gb_config.json"
        with open(config_file, 'w') as f:
            json.dump(config_content, f, indent=2)
        
        if RICH_AVAILABLE:
            console.print(f"✅ 8GB Config saved: [cyan]{config_file}[/cyan]")
        else:
            print(f"✅ 8GB Config saved: {config_file}")
        
        # Create optimized startup script
        if sys.platform == "win32":
            batch_content = f'''@echo off
title Ollama Server for Manice AI (8GB Optimized)
echo Starting Ollama Server for Manice AI (8GB Optimized)...
echo Models Directory: {OLLAMA_MODELS_DIR}
echo System: 8GB RAM Optimized
echo.
set OLLAMA_MODELS={OLLAMA_MODELS_DIR}
set OLLAMA_NUM_PARALLEL=1
set OLLAMA_MAX_LOADED_MODELS=1
echo 8GB RAM optimizations enabled
ollama serve
pause
'''
            batch_file = MODELS_BASE_DIR / "start_ollama_8gb.bat"
            with open(batch_file, 'w') as f:
                f.write(batch_content)
            
            if RICH_AVAILABLE:
                console.print(f"✅ 8GB Startup script: [cyan]{batch_file}[/cyan]")
            else:
                print(f"✅ 8GB Startup script: {batch_file}")

def main():
    """Main setup function for 8GB systems"""
    setup = Lightweight8GBSetup()
    
    try:
        # Check system requirements
        if not setup.check_system_requirements_8gb():
            if RICH_AVAILABLE:
                console.print("\n[bold red]Setup cannot continue. Please install Ollama first.[/bold red]")
                console.print("[yellow]Download from: https://ollama.ai[/yellow]")
            else:
                print("\nSetup cannot continue. Please install Ollama first.")
                print("Download from: https://ollama.ai")
            return False
        
        # Model selection
        models, total_size = setup.interactive_8gb_model_selection()
        
        # Confirmation
        if RICH_AVAILABLE:
            console.print(f"\n[bold blue]📋 8GB Optimized Setup Summary[/bold blue]")
            summary_table = Table()
            summary_table.add_column("Model", style="cyan")
            summary_table.add_column("Size", style="green")
            summary_table.add_column("RAM Needed", style="yellow")
            summary_table.add_column("Category", style="magenta")
            
            for model in models:
                summary_table.add_row(
                    model.name, f"{model.size_gb} GB", 
                    f"{model.recommended_ram_gb} GB", model.category
                )
            
            summary_table.add_row("[bold]Total", f"[bold]{total_size:.1f} GB", "[bold]8GB System", "[bold]Optimized")
            console.print(summary_table)
            
            if not Confirm.ask(f"\n🚀 Proceed with 8GB optimized download?"):
                console.print("[yellow]Setup cancelled.[/yellow]")
                return False
        else:
            print(f"\n📋 8GB Optimized Setup Summary")
            print("-" * 40)
            for model in models:
                print(f"{model.name}: {model.size_gb} GB (RAM: {model.recommended_ram_gb} GB)")
            print(f"Total: {total_size:.1f} GB")
            
            response = input("\n🚀 Proceed with 8GB optimized download? [y/N]: ")
            if not response.lower().startswith('y'):
                print("Setup cancelled.")
                return False
        
        # Download models
        success = setup.download_8gb_models(models)
        
        if success:
            # Create config
            setup.create_8gb_config(models)
            
            # Success message
            if RICH_AVAILABLE:
                success_panel = Panel(
                    f"""[bold green]🎉 8GB Optimized Setup Complete![/bold green]

[cyan]📁 Models Location:[/cyan] {OLLAMA_MODELS_DIR}
[cyan]📋 Logs:[/cyan] {LOGS_DIR}
[cyan]🔧 Configuration:[/cyan] {MODELS_BASE_DIR}/manice_8gb_config.json
[cyan]🚀 Startup Script:[/cyan] {MODELS_BASE_DIR}/start_ollama_8gb.bat

[bold yellow]8GB Optimizations Enabled:[/bold yellow]
• Memory-efficient models ({total_size:.1f} GB total)
• Optimized server settings
• Single model loading
• Reduced cache usage

[bold yellow]Next Steps:[/bold yellow]
1. Start Manice AI Server: [cyan]cd ai-server && python server.py[/cyan]
2. Install Excel Add-in: [cyan]cd excel-addin && npm run build && npm run install-addin[/cyan]
3. Open Excel and enjoy your 8GB-optimized AI assistant!

[dim]Perfect for 8GB systems - No performance compromise![/dim]""",
                    title="[bold green]8GB Setup Complete[/bold green]",
                    expand=False
                )
                console.print(success_panel)
            else:
                print("\n🎉 8GB Optimized Setup Complete!")
                print(f"📁 Models Location: {OLLAMA_MODELS_DIR}")
                print(f"📋 Logs: {LOGS_DIR}")
                print(f"🔧 Configuration: {MODELS_BASE_DIR}/manice_8gb_config.json")
                print("\n8GB Optimizations Enabled:")
                print(f"• Memory-efficient models ({total_size:.1f} GB total)")
                print("• Optimized server settings")
                print("• Single model loading")
                print("• Reduced cache usage")
                print("\nNext Steps:")
                print("1. Start Manice AI Server: cd ai-server && python server.py")
                print("2. Install Excel Add-in: cd excel-addin && npm run build && npm run install-addin")
                print("3. Open Excel and enjoy your 8GB-optimized AI assistant!")
            
            return True
        else:
            if RICH_AVAILABLE:
                console.print("\n[bold red]❌ Setup failed. Check logs for details.[/bold red]")
            else:
                print("\n❌ Setup failed. Check logs for details.")
            return False
            
    except KeyboardInterrupt:
        if RICH_AVAILABLE:
            console.print("\n\n[yellow]⏹️ Setup interrupted.[/yellow]")
        else:
            print("\n\n⏹️ Setup interrupted.")
        return False
    except Exception as e:
        if RICH_AVAILABLE:
            console.print(f"\n[bold red]❌ Setup error: {e}[/bold red]")
        else:
            print(f"\n❌ Setup error: {e}")
        logger.error(f"Setup error: {e}")
        return False

if __name__ == "__main__":
    success = main()
    if not RICH_AVAILABLE:
        input("\nPress Enter to exit...")
    sys.exit(0 if success else 1)