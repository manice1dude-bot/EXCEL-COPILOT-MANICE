"""
AI Model Interface for Manice Excel CoPilot
Handles communication with DeepSeek R1 and Mistral-7B models via local providers
"""

import asyncio
import httpx
import json
import time
from typing import Dict, List, Optional, Union, Any, AsyncGenerator
from dataclasses import dataclass
from enum import Enum
from loguru import logger

try:
    from ..config import server_config, AIModelConfig
except ImportError:
    from config import server_config, AIModelConfig


class ModelProvider(Enum):
    """Supported local AI model providers"""
    OLLAMA = "ollama"
    LM_STUDIO = "lm_studio" 
    JAN = "jan"


@dataclass
class ModelResponse:
    """Response from AI model"""
    content: str
    model_used: str
    provider: ModelProvider
    tokens_used: int
    response_time: float
    metadata: Dict[str, Any]


@dataclass
class ExcelContext:
    """Excel context for AI requests"""
    sheet_name: Optional[str] = None
    selected_range: Optional[str] = None
    cell_data: Optional[Dict[str, Any]] = None
    workbook_info: Optional[Dict[str, Any]] = None
    previous_operations: Optional[List[str]] = None


class AIModelInterface:
    """Main interface for AI model communication"""
    
    def __init__(self):
        self.client = httpx.AsyncClient(timeout=60.0)
        self.model_cache: Dict[str, ModelResponse] = {}
        self.request_history: List[str] = []
        
    async def __aenter__(self):
        return self
        
    async def __aexit__(self, exc_type, exc_val, exc_tb):
        await self.client.aclose()
    
    async def generate_response(
        self, 
        prompt: str,
        model_type: Optional[str] = None,
        excel_context: Optional[ExcelContext] = None,
        stream: bool = False
    ) -> Union[ModelResponse, AsyncGenerator[str, None]]:
        """
        Generate AI response using appropriate model
        
        Args:
            prompt: User input/instruction
            model_type: Force specific model ("large" or "small")  
            excel_context: Current Excel context
            stream: Whether to stream response
            
        Returns:
            ModelResponse or streaming generator
        """
        start_time = time.time()
        
        # Determine model to use
        if not model_type:
            try:
                from ..config import get_model_for_task
            except ImportError:
                from config import get_model_for_task
            context_size = self._calculate_context_size(excel_context)
            model_type = get_model_for_task(prompt, context_size)
        
        # Get model config
        model_config = (server_config.large_model 
                       if model_type == "large" 
                       else server_config.small_model)
        
        # Build enhanced prompt with Excel context
        enhanced_prompt = self._build_enhanced_prompt(prompt, excel_context)
        
        # Check cache first
        cache_key = f"{model_type}:{hash(enhanced_prompt)}"
        if cache_key in self.model_cache:
            cached = self.model_cache[cache_key]
            logger.info(f"Returning cached response for {model_type} model")
            return cached
        
        # Generate response
        try:
            if server_config.preferred_provider == "ollama":
                response = await self._query_ollama(
                    enhanced_prompt, model_config, stream
                )
            elif server_config.preferred_provider == "lm_studio":
                response = await self._query_lm_studio(
                    enhanced_prompt, model_config, stream
                )
            elif server_config.preferred_provider == "jan":
                response = await self._query_jan(
                    enhanced_prompt, model_config, stream
                )
            else:
                raise ValueError(f"Unknown provider: {server_config.preferred_provider}")
            
            if not stream:
                # Cache successful responses
                self.model_cache[cache_key] = response
                self._cleanup_cache()
                
            return response
            
        except Exception as e:
            logger.error(f"Error generating response: {e}")
            # Fallback to alternative provider
            return await self._fallback_response(enhanced_prompt, model_config)
    
    async def _query_ollama(
        self, 
        prompt: str, 
        model_config: AIModelConfig,
        stream: bool = False
    ) -> ModelResponse:
        """Query Ollama API"""
        
        payload = {
            "model": model_config.model_path,
            "prompt": prompt,
            "options": {
                "temperature": model_config.temperature,
                "num_predict": model_config.max_tokens
            },
            "stream": stream
        }
        
        start_time = time.time()
        
        try:
            response = await self.client.post(
                f"{server_config.ollama_url}/api/generate",
                json=payload,
                timeout=model_config.timeout
            )
            response.raise_for_status()
            
            if stream:
                return self._handle_ollama_stream(response)
            else:
                result = response.json()
                
                return ModelResponse(
                    content=result.get("response", ""),
                    model_used=model_config.name,
                    provider=ModelProvider.OLLAMA,
                    tokens_used=result.get("eval_count", 0),
                    response_time=time.time() - start_time,
                    metadata={
                        "total_duration": result.get("total_duration", 0),
                        "load_duration": result.get("load_duration", 0),
                        "eval_duration": result.get("eval_duration", 0)
                    }
                )
                
        except httpx.RequestError as e:
            logger.error(f"Ollama request failed: {e}")
            raise
        except httpx.HTTPStatusError as e:
            logger.error(f"Ollama HTTP error: {e}")
            raise
    
    async def _query_lm_studio(
        self, 
        prompt: str, 
        model_config: AIModelConfig,
        stream: bool = False
    ) -> ModelResponse:
        """Query LM Studio API"""
        
        payload = {
            "model": model_config.model_path,
            "messages": [
                {
                    "role": "system",
                    "content": "You are Manice, an Excel AI assistant."
                },
                {
                    "role": "user", 
                    "content": prompt
                }
            ],
            "temperature": model_config.temperature,
            "max_tokens": model_config.max_tokens,
            "stream": stream
        }
        
        start_time = time.time()
        
        try:
            response = await self.client.post(
                f"{server_config.lm_studio_url}/v1/chat/completions",
                json=payload,
                timeout=model_config.timeout
            )
            response.raise_for_status()
            
            if stream:
                return self._handle_lm_studio_stream(response)
            else:
                result = response.json()
                message = result["choices"][0]["message"]["content"]
                
                return ModelResponse(
                    content=message,
                    model_used=model_config.name,
                    provider=ModelProvider.LM_STUDIO,
                    tokens_used=result.get("usage", {}).get("total_tokens", 0),
                    response_time=time.time() - start_time,
                    metadata=result.get("usage", {})
                )
                
        except Exception as e:
            logger.error(f"LM Studio request failed: {e}")
            raise
    
    async def _query_jan(
        self, 
        prompt: str, 
        model_config: AIModelConfig,
        stream: bool = False
    ) -> ModelResponse:
        """Query Jan API"""
        
        # Jan uses OpenAI-compatible API
        payload = {
            "model": model_config.model_path,
            "messages": [
                {
                    "role": "system",
                    "content": "You are Manice, a helpful Excel AI assistant."
                },
                {
                    "role": "user",
                    "content": prompt
                }
            ],
            "temperature": model_config.temperature,
            "max_tokens": model_config.max_tokens,
            "stream": stream
        }
        
        start_time = time.time()
        
        try:
            response = await self.client.post(
                f"{server_config.jan_url}/v1/chat/completions",
                json=payload,
                timeout=model_config.timeout
            )
            response.raise_for_status()
            
            result = response.json()
            message = result["choices"][0]["message"]["content"]
            
            return ModelResponse(
                content=message,
                model_used=model_config.name,
                provider=ModelProvider.JAN,
                tokens_used=result.get("usage", {}).get("total_tokens", 0),
                response_time=time.time() - start_time,
                metadata=result.get("usage", {})
            )
            
        except Exception as e:
            logger.error(f"Jan request failed: {e}")
            raise
    
    def _build_enhanced_prompt(self, prompt: str, context: Optional[ExcelContext]) -> str:
        """Build enhanced prompt with Excel context"""
        
        enhanced = f"""You are Manice, an advanced Excel AI CoPilot assistant. You can:

1. Read and analyze Excel data
2. Modify cells, rows, columns, and formatting in real-time
3. Create formulas and explain existing ones
4. Generate charts and visualizations  
5. Perform complex data analysis and business intelligence
6. Convert natural language to Excel actions

Current task: {prompt}
"""
        
        if context:
            enhanced += f"\n## Excel Context:\n"
            
            if context.sheet_name:
                enhanced += f"- Active Sheet: {context.sheet_name}\n"
                
            if context.selected_range:
                enhanced += f"- Selected Range: {context.selected_range}\n"
                
            if context.cell_data:
                enhanced += f"- Cell Data: {json.dumps(context.cell_data, indent=2)}\n"
                
            if context.workbook_info:
                enhanced += f"- Workbook Info: {json.dumps(context.workbook_info, indent=2)}\n"
                
            if context.previous_operations:
                enhanced += f"- Previous Operations: {context.previous_operations}\n"
        
        enhanced += """\n
## Response Format:
Provide a JSON response with these fields:
{
  "action": "excel_operation_type", 
  "parameters": {...},
  "explanation": "Human-readable explanation",
  "excel_operations": [
    {
      "type": "cell_edit|formula|format|chart|etc",
      "target": "A1:B10", 
      "value": "...",
      "options": {...}
    }
  ]
}

Be precise, actionable, and always consider Excel's capabilities and limitations.
"""
        
        return enhanced
    
    def _calculate_context_size(self, context: Optional[ExcelContext]) -> int:
        """Calculate approximate context size for model selection"""
        if not context:
            return 0
            
        size = 0
        
        if context.cell_data:
            size += len(str(context.cell_data))
            
        if context.workbook_info:
            size += len(str(context.workbook_info))
            
        if context.previous_operations:
            size += sum(len(op) for op in context.previous_operations)
            
        return size
    
    async def _fallback_response(self, prompt: str, model_config: AIModelConfig) -> ModelResponse:
        """Fallback response when primary provider fails"""
        
        logger.warning("Using fallback response generation")
        
        return ModelResponse(
            content=json.dumps({
                "action": "error",
                "explanation": "AI model temporarily unavailable. Please try again.",
                "excel_operations": []
            }),
            model_used="fallback",
            provider=ModelProvider.OLLAMA,
            tokens_used=0,
            response_time=0.1,
            metadata={"fallback": True}
        )
    
    def _cleanup_cache(self):
        """Clean up model response cache to prevent memory issues"""
        if len(self.model_cache) > server_config.model_cache_size:
            # Remove oldest entries
            keys_to_remove = list(self.model_cache.keys())[:-server_config.model_cache_size]
            for key in keys_to_remove:
                del self.model_cache[key]
    
    async def health_check(self) -> Dict[str, Any]:
        """Check health of AI models and providers"""
        health = {
            "status": "healthy",
            "providers": {},
            "models": {},
            "timestamp": time.time()
        }
        
        # Check each provider
        for provider in ["ollama", "lm_studio", "jan"]:
            try:
                if provider == "ollama":
                    url = f"{server_config.ollama_url}/api/tags"
                else:
                    url = f"{getattr(server_config, f'{provider}_url')}/v1/models"
                
                response = await self.client.get(url, timeout=5.0)
                health["providers"][provider] = {
                    "available": response.status_code == 200,
                    "response_time": response.elapsed.total_seconds()
                }
                
            except Exception as e:
                health["providers"][provider] = {
                    "available": False,
                    "error": str(e)
                }
        
        return health


# Singleton instance
_ai_interface = None

async def get_ai_interface() -> AIModelInterface:
    """Get or create AI interface singleton"""
    global _ai_interface
    if _ai_interface is None:
        _ai_interface = AIModelInterface()
    return _ai_interface