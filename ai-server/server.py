"""
Manice AI Server
Main FastAPI application for Excel AI CoPilot with real-time integration
"""

import asyncio
import json
import time
import traceback
from datetime import datetime
from typing import Dict, List, Optional, Any

from fastapi import FastAPI, HTTPException, Depends, BackgroundTasks
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse, StreamingResponse
from pydantic import BaseModel, Field
from loguru import logger
import uvicorn

try:
    from config import server_config, excel_config
    from models import get_ai_interface, ExcelContext, ModelResponse
    from routes.formula_routes import router as formula_router
except ImportError:
    from .config import server_config, excel_config
    from .models import get_ai_interface, ExcelContext, ModelResponse
    from .routes.formula_routes import router as formula_router


# Pydantic Models for API
class ManiceRequest(BaseModel):
    """Request model for Manice AI operations"""
    instruction: str = Field(..., description="Natural language instruction")
    context: Optional[Dict[str, Any]] = Field(None, description="Excel context data")
    force_model: Optional[str] = Field(None, description="Force specific model (large/small)")
    stream: bool = Field(False, description="Stream response")


class ExcelOperation(BaseModel):
    """Model for Excel operations to be executed"""
    type: str = Field(..., description="Operation type")
    target: str = Field(..., description="Target range/cell")
    value: Optional[Any] = Field(None, description="Value to set")
    options: Optional[Dict[str, Any]] = Field(None, description="Additional options")


class ManiceResponse(BaseModel):
    """Response model for Manice AI operations"""
    action: str = Field(..., description="Action type")
    explanation: str = Field(..., description="Human readable explanation")
    excel_operations: List[ExcelOperation] = Field(default=[], description="Excel operations to execute")
    parameters: Optional[Dict[str, Any]] = Field(None, description="Additional parameters")
    model_info: Optional[Dict[str, Any]] = Field(None, description="Model execution info")
    timestamp: float = Field(default_factory=time.time, description="Response timestamp")


class HealthResponse(BaseModel):
    """Health check response model"""
    status: str
    version: str
    uptime: float
    providers: Dict[str, Any]
    models: Dict[str, Any]
    timestamp: float


# Initialize FastAPI app
app = FastAPI(
    title="Manice AI Server",
    description="Offline Excel AI CoPilot with DeepSeek R1 + Mistral-7B",
    version="1.0.0",
    docs_url="/docs" if server_config.debug else None,
    redoc_url="/redoc" if server_config.debug else None
)

# CORS middleware for Excel add-in communication
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # In production, restrict to Excel origins
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Include routers
app.include_router(formula_router)

# Global state
server_start_time = time.time()
request_count = 0
active_requests = 0


@app.middleware("http")
async def logging_middleware(request, call_next):
    """Log all HTTP requests"""
    global request_count, active_requests
    
    start_time = time.time()
    request_count += 1
    active_requests += 1
    
    try:
        response = await call_next(request)
        process_time = time.time() - start_time
        
        logger.info(
            f"{request.method} {request.url.path} - "
            f"Status: {response.status_code} - "
            f"Time: {process_time:.3f}s"
        )
        
        response.headers["X-Process-Time"] = str(process_time)
        response.headers["X-Request-ID"] = str(request_count)
        
        return response
        
    except Exception as e:
        logger.error(f"Request failed: {e}")
        raise
    finally:
        active_requests -= 1


# Dependency to get AI interface
async def get_ai_service():
    """Dependency to get AI interface"""
    return await get_ai_interface()


@app.get("/", response_model=Dict[str, str])
async def root():
    """Root endpoint"""
    return {
        "service": "Manice AI Server",
        "status": "running",
        "version": "1.0.0",
        "docs": "/docs" if server_config.debug else "disabled"
    }


@app.get("/health", response_model=HealthResponse)
async def health_check(ai_service=Depends(get_ai_service)):
    """Comprehensive health check"""
    
    try:
        # Get AI model health
        ai_health = await ai_service.health_check()
        
        return HealthResponse(
            status="healthy",
            version="1.0.0",
            uptime=time.time() - server_start_time,
            providers=ai_health.get("providers", {}),
            models={
                "large_model": server_config.large_model.dict(),
                "small_model": server_config.small_model.dict()
            },
            timestamp=time.time()
        )
        
    except Exception as e:
        logger.error(f"Health check failed: {e}")
        return HealthResponse(
            status="unhealthy",
            version="1.0.0",
            uptime=time.time() - server_start_time,
            providers={},
            models={},
            timestamp=time.time()
        )


@app.post("/manice", response_model=ManiceResponse)
async def process_manice_request(
    request: ManiceRequest,
    ai_service=Depends(get_ai_service)
):
    """
    Main endpoint for processing Manice AI requests
    Handles both formula function and sidebar requests
    """
    
    logger.info(f"Processing request: {request.instruction[:100]}...")
    
    try:
        # Build Excel context from request
        excel_context = None
        if request.context:
            excel_context = ExcelContext(
                sheet_name=request.context.get("sheet_name"),
                selected_range=request.context.get("selected_range"),
                cell_data=request.context.get("cell_data"),
                workbook_info=request.context.get("workbook_info"),
                previous_operations=request.context.get("previous_operations", [])
            )
        
        # Generate AI response
        model_response = await ai_service.generate_response(
            prompt=request.instruction,
            model_type=request.force_model,
            excel_context=excel_context,
            stream=False  # Non-streaming for now
        )
        
        # Parse AI response
        try:
            ai_content = json.loads(model_response.content)
        except json.JSONDecodeError:
            # Fallback if AI doesn't return valid JSON
            ai_content = {
                "action": "text_response",
                "explanation": model_response.content,
                "excel_operations": []
            }
        
        # Build response
        response = ManiceResponse(
            action=ai_content.get("action", "unknown"),
            explanation=ai_content.get("explanation", "Action completed"),
            excel_operations=[
                ExcelOperation(**op) for op in ai_content.get("excel_operations", [])
            ],
            parameters=ai_content.get("parameters"),
            model_info={
                "model_used": model_response.model_used,
                "provider": model_response.provider.value,
                "tokens_used": model_response.tokens_used,
                "response_time": model_response.response_time
            }
        )
        
        logger.info(f"Request completed successfully using {model_response.model_used}")
        return response
        
    except Exception as e:
        logger.error(f"Error processing request: {e}")
        logger.error(traceback.format_exc())
        
        # Return error response
        return ManiceResponse(
            action="error",
            explanation=f"An error occurred: {str(e)}",
            excel_operations=[],
            model_info={"error": True}
        )


@app.post("/manice/stream")
async def stream_manice_request(
    request: ManiceRequest,
    ai_service=Depends(get_ai_service)
):
    """
    Streaming endpoint for real-time AI responses
    """
    
    if not request.stream:
        raise HTTPException(status_code=400, detail="Stream must be enabled for this endpoint")
    
    logger.info(f"Starting stream for: {request.instruction[:100]}...")
    
    async def generate_stream():
        try:
            # Build Excel context
            excel_context = None
            if request.context:
                excel_context = ExcelContext(
                    sheet_name=request.context.get("sheet_name"),
                    selected_range=request.context.get("selected_range"),
                    cell_data=request.context.get("cell_data"),
                    workbook_info=request.context.get("workbook_info"),
                    previous_operations=request.context.get("previous_operations", [])
                )
            
            # Generate streaming response
            async for chunk in await ai_service.generate_response(
                prompt=request.instruction,
                model_type=request.force_model,
                excel_context=excel_context,
                stream=True
            ):
                yield f"data: {json.dumps({'chunk': chunk})}\n\n"
                
        except Exception as e:
            logger.error(f"Streaming error: {e}")
            yield f"data: {json.dumps({'error': str(e)})}\n\n"
        
        yield f"data: {json.dumps({'done': True})}\n\n"
    
    return StreamingResponse(
        generate_stream(),
        media_type="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "Connection": "keep-alive",
        }
    )


@app.post("/excel/operation")
async def execute_excel_operation(
    operations: List[ExcelOperation],
    background_tasks: BackgroundTasks
):
    """
    Execute Excel operations directly
    This endpoint would be called by the Excel add-in to perform actions
    """
    
    logger.info(f"Executing {len(operations)} Excel operations")
    
    results = []
    
    for operation in operations:
        try:
            # This is where we would integrate with Excel COM/JS API
            # For now, we'll simulate the operation
            result = {
                "operation_id": f"op_{int(time.time() * 1000)}",
                "type": operation.type,
                "target": operation.target,
                "status": "simulated",  # Would be "success" or "error" in real implementation
                "message": f"Simulated {operation.type} on {operation.target}"
            }
            
            results.append(result)
            
        except Exception as e:
            logger.error(f"Operation failed: {e}")
            results.append({
                "operation_id": f"op_{int(time.time() * 1000)}",
                "type": operation.type,
                "target": operation.target,
                "status": "error",
                "message": str(e)
            })
    
    return {"operations": results, "timestamp": time.time()}


@app.get("/stats")
async def get_server_stats():
    """Get server statistics"""
    return {
        "uptime": time.time() - server_start_time,
        "requests_total": request_count,
        "active_requests": active_requests,
        "configuration": {
            "debug": server_config.debug,
            "models": {
                "large": server_config.large_model.name,
                "small": server_config.small_model.name
            },
            "provider": server_config.preferred_provider
        },
        "timestamp": time.time()
    }


@app.get("/config")
async def get_configuration():
    """Get current configuration (debug mode only)"""
    if not server_config.debug:
        raise HTTPException(status_code=404, detail="Not found")
    
    return {
        "server": server_config.dict(),
        "excel": excel_config.dict()
    }


# Error handlers
@app.exception_handler(404)
async def not_found_handler(request, exc):
    return JSONResponse(
        status_code=404,
        content={"error": "Endpoint not found", "timestamp": time.time()}
    )


@app.exception_handler(500)
async def internal_error_handler(request, exc):
    logger.error(f"Internal server error: {exc}")
    return JSONResponse(
        status_code=500,
        content={"error": "Internal server error", "timestamp": time.time()}
    )


# Startup and shutdown events
@app.on_event("startup")
async def startup_event():
    """Initialize services on startup"""
    logger.info("Starting Manice AI Server...")
    logger.info(f"Using {server_config.preferred_provider} provider")
    logger.info(f"Large model: {server_config.large_model.name}")
    logger.info(f"Small model: {server_config.small_model.name}")
    
    # Initialize AI interface
    ai_service = await get_ai_interface()
    health = await ai_service.health_check()
    logger.info(f"AI service health: {health['status']}")


@app.on_event("shutdown")
async def shutdown_event():
    """Cleanup on shutdown"""
    logger.info("Shutting down Manice AI Server...")


def main():
    """Run the server"""
    logger.info(f"Starting server on {server_config.host}:{server_config.port}")
    
    uvicorn.run(
        "server:app",
        host=server_config.host,
        port=server_config.port,
        reload=server_config.debug,
        log_level=server_config.log_level.lower(),
        access_log=server_config.debug
    )


if __name__ == "__main__":
    main()