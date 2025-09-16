"""
Configuration management for Manice AI Server
Handles model routing, server settings, and Excel integration parameters
"""

import os
from typing import Dict, List, Optional, Literal
from pydantic import Field
from pydantic_settings import BaseSettings
from pathlib import Path


class AIModelConfig(BaseSettings):
    """Configuration for individual AI models"""
    name: str
    type: Literal["large", "small"] 
    model_path: str
    api_endpoint: Optional[str] = None
    max_tokens: int = 4096
    temperature: float = 0.7
    timeout: int = 30
    memory_gb: float = 4.0
    

class ServerConfig(BaseSettings):
    """Main server configuration"""
    
    # Server Settings
    host: str = "127.0.0.1"
    port: int = 8899
    debug: bool = False
    log_level: str = "INFO"
    # Optimized Model Configurations for 8GB RAM Systems
    large_model: AIModelConfig = AIModelConfig(
        name="phi3-medium", 
        type="large",
        model_path="phi3:medium",  # 7GB model, more efficient than DeepSeek R1
        max_tokens=4096,
        temperature=0.3,
        timeout=45,
        memory_gb=4.0  # Reduced memory allocation for 8GB systems
    )
    
    small_model: AIModelConfig = AIModelConfig(
        name="phi3-mini",
        type="small", 
        model_path="phi3:mini",  # 2GB model, very efficient
        max_tokens=2048,
        temperature=0.5,
        timeout=15,
        memory_gb=2.0  # Minimal memory usage
    )
    # Model Routing Rules
    complexity_threshold: float = 0.7  # 0.0-1.0 scale
    large_model_keywords: List[str] = [
        "analyze", "forecast", "complex", "reasoning", "explain", "why",
        "prediction", "trend", "correlation", "statistical", "machine learning",
        "business intelligence", "dashboard", "report", "insights", "optimization"
    ]
    
    small_model_keywords: List[str] = [
        "formula", "calculate", "sum", "format", "color", "highlight", 
        "insert", "delete", "copy", "paste", "sort", "filter", "quick",
        "simple", "basic", "cell", "row", "column"
    ]
    
    # Excel Integration
    excel_com_timeout: int = 10
    max_cells_batch: int = 10000
    enable_undo_tracking: bool = True
    confirm_destructive_ops: bool = True
    
    # Performance Settings Optimized for 8GB RAM
    model_cache_size: int = 2  # Reduced cache to save memory
    concurrent_requests: int = 2  # Limit concurrent requests for stability
    request_queue_size: int = 20  # Smaller queue to reduce memory usage
    
    # Local Model Providers
    ollama_url: str = "http://127.0.0.1:11434"
    lm_studio_url: str = "http://127.0.0.1:1234"  
    jan_url: str = "http://127.0.0.1:1337"
    preferred_provider: Literal["ollama", "lm_studio", "jan"] = "ollama"
    
    class Config:
        env_prefix = "MANICE_"
        env_file = ".env"


class ExcelOperationConfig(BaseSettings):
    """Configuration for Excel-specific operations"""
    
    # Safety Settings
    max_rows_delete: int = 1000
    max_columns_delete: int = 50
    backup_before_destructive: bool = True
    
    # Chart Settings
    default_chart_type: str = "column"
    max_chart_data_points: int = 1000
    enable_interactive_charts: bool = True
    
    # Formula Settings
    max_formula_complexity: int = 10  # Nested function depth
    enable_vba_generation: bool = True
    vba_security_check: bool = True
    
    # Formatting
    default_number_format: str = "General"
    max_conditional_formats: int = 64  # Excel limit
    
    class Config:
        env_prefix = "MANICE_EXCEL_"


# Global configuration instances
server_config = ServerConfig()
excel_config = ExcelOperationConfig()


def get_model_for_task(task_description: str, context_size: int = 0) -> Literal["large", "small"]:
    """
    Determine which AI model to use based on task complexity
    
    Args:
        task_description: Natural language description of the task
        context_size: Size of context/data being processed
    
    Returns:
        "large" for DeepSeek R1 or "small" for Mistral-7B
    """
    task_lower = task_description.lower()
    
    # Check for explicit large model keywords
    large_score = sum(1 for keyword in server_config.large_model_keywords 
                     if keyword in task_lower)
    
    # Check for explicit small model keywords  
    small_score = sum(1 for keyword in server_config.small_model_keywords
                     if keyword in task_lower)
    
    # Context size influence
    if context_size > 5000:  # Large datasets need large model
        large_score += 2
    
    # Length and complexity heuristics
    if len(task_description) > 200:  # Detailed requests
        large_score += 1
    
    if any(word in task_lower for word in ["because", "explain", "why", "how", "analyze"]):
        large_score += 1
        
    # Final decision
    if large_score > small_score:
        return "large"
    elif small_score > large_score:
        return "small"
    else:
        # Default to small for balanced cases (faster)
        return "small"


def validate_model_availability() -> Dict[str, bool]:
    """Check if configured AI models are available"""
    availability = {
        "ollama": False,
        "lm_studio": False, 
        "jan": False,
        "deepseek_r1": False,
        "mistral_7b": False
    }
    
    # This will be implemented to actually check model availability
    # For now, assume available for development
    availability.update({
        "ollama": True,
        "deepseek_r1": True,
        "mistral_7b": True
    })
    
    return availability


# Environment-specific overrides
if os.getenv("MANICE_ENV") == "development":
    server_config.debug = True
    server_config.log_level = "DEBUG"
    server_config.large_model.timeout = 120  # More time for development
    
elif os.getenv("MANICE_ENV") == "production":
    server_config.debug = False
    server_config.log_level = "WARNING"
    excel_config.backup_before_destructive = True