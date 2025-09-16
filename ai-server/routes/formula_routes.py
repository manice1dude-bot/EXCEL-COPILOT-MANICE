"""
Formula and VBA API Routes for Manice Excel AI Copilot
Provides endpoints for formula generation, debugging, and VBA macro creation
"""

import json
import time
from typing import Dict, List, Optional, Any
from fastapi import APIRouter, HTTPException, Depends
from pydantic import BaseModel, Field
from loguru import logger

try:
    from ..services.formula_engine import FormulaGenerator, FormulaDebugger, FormulaResult, FormulaError
    from ..services.vba_engine import VBAMacroEngine, VBAMacro, MacroAnalysis
    from ..models import get_ai_interface
except ImportError:
    from services.formula_engine import FormulaGenerator, FormulaDebugger, FormulaResult, FormulaError
    from services.vba_engine import VBAMacroEngine, VBAMacro, MacroAnalysis
    from models import get_ai_interface

# Create router
router = APIRouter(prefix="/api/v1", tags=["Formula & VBA"])

# Pydantic Models
class FormulaRequest(BaseModel):
    """Request model for formula generation"""
    requirement: str = Field(..., description="Natural language requirement for the formula")
    context: Optional[Dict[str, Any]] = Field(None, description="Excel context (sheet data, ranges, etc.)")
    data_sample: Optional[Dict[str, Any]] = Field(None, description="Sample data to work with")

class FormulaDebugRequest(BaseModel):
    """Request model for formula debugging"""
    formula: str = Field(..., description="Excel formula to debug")
    error_message: Optional[str] = Field(None, description="Error message if any")
    cell_data: Optional[Dict[str, Any]] = Field(None, description="Cell data context")

class VBARequest(BaseModel):
    """Request model for VBA macro generation"""
    requirement: str = Field(..., description="Natural language requirement for the macro")
    context: Optional[Dict[str, Any]] = Field(None, description="Excel context")
    constraints: Optional[Dict[str, Any]] = Field(None, description="Security/performance constraints")

class VBAAnalysisRequest(BaseModel):
    """Request model for VBA macro analysis"""
    code: str = Field(..., description="VBA code to analyze")
    context: Optional[Dict[str, Any]] = Field(None, description="Analysis context")

class FormulaResponse(BaseModel):
    """Response model for formula operations"""
    success: bool = Field(..., description="Whether operation was successful")
    formula: Optional[str] = Field(None, description="Generated/debugged formula")
    explanation: Optional[str] = Field(None, description="Human-readable explanation")
    category: Optional[str] = Field(None, description="Formula category")
    complexity: Optional[str] = Field(None, description="Formula complexity level")
    examples: List[str] = Field(default=[], description="Usage examples")
    alternatives: List[Dict[str, str]] = Field(default=[], description="Alternative formulas")
    prerequisites: List[str] = Field(default=[], description="Prerequisites for using the formula")
    confidence: Optional[float] = Field(None, description="Confidence score (0-1)")
    errors: List[Dict[str, Any]] = Field(default=[], description="Identified errors")
    solutions: List[Dict[str, Any]] = Field(default=[], description="Proposed solutions")
    optimizations: List[Dict[str, Any]] = Field(default=[], description="Optimization suggestions")
    timestamp: float = Field(default_factory=time.time)

class VBAResponse(BaseModel):
    """Response model for VBA operations"""
    success: bool = Field(..., description="Whether operation was successful")
    name: Optional[str] = Field(None, description="Macro name")
    code: Optional[str] = Field(None, description="Generated VBA code")
    description: Optional[str] = Field(None, description="Macro description")
    category: Optional[str] = Field(None, description="Macro category")
    complexity: Optional[str] = Field(None, description="Macro complexity level")
    security_level: Optional[str] = Field(None, description="Security risk level")
    dependencies: List[str] = Field(default=[], description="Required dependencies")
    parameters: List[Dict[str, str]] = Field(default=[], description="Macro parameters")
    usage_examples: List[str] = Field(default=[], description="Usage examples")
    warnings: List[str] = Field(default=[], description="Security/usage warnings")
    performance_notes: List[str] = Field(default=[], description="Performance considerations")
    analysis: Optional[Dict[str, Any]] = Field(None, description="Code analysis results")
    timestamp: float = Field(default_factory=time.time)

# Dependencies
async def get_formula_generator():
    """Get formula generator instance"""
    ai_service = await get_ai_interface()
    return FormulaGenerator(ai_service)

async def get_formula_debugger():
    """Get formula debugger instance"""
    ai_service = await get_ai_interface()
    return FormulaDebugger(ai_service)

async def get_vba_engine():
    """Get VBA engine instance"""
    ai_service = await get_ai_interface()
    return VBAMacroEngine(ai_service)

# Formula Generation Endpoints
@router.post("/formula/generate", response_model=FormulaResponse)
async def generate_formula(
    request: FormulaRequest,
    generator: FormulaGenerator = Depends(get_formula_generator)
):
    """
    Generate Excel formula from natural language requirement
    """
    logger.info(f"Generating formula for: {request.requirement[:100]}...")
    
    try:
        result = await generator.generate_formula(
            requirement=request.requirement,
            context=request.context,
            data_sample=request.data_sample
        )
        
        return FormulaResponse(
            success=True,
            formula=result.formula,
            explanation=result.explanation,
            category=result.category.value,
            complexity=result.complexity.value,
            examples=result.examples,
            alternatives=result.alternatives,
            prerequisites=result.prerequisites,
            confidence=result.confidence
        )
        
    except Exception as e:
        logger.error(f"Formula generation failed: {e}")
        return FormulaResponse(
            success=False,
            explanation=f"Error generating formula: {str(e)}"
        )

@router.post("/formula/debug", response_model=FormulaResponse)
async def debug_formula(
    request: FormulaDebugRequest,
    debugger: FormulaDebugger = Depends(get_formula_debugger)
):
    """
    Debug Excel formula and provide solutions
    """
    logger.info(f"Debugging formula: {request.formula}")
    
    try:
        result = await debugger.debug_formula(
            formula=request.formula,
            error_message=request.error_message,
            cell_data=request.cell_data
        )
        
        # Convert errors to dict format
        errors = []
        for error in result["errors"]:
            errors.append({
                "type": error.error_type,
                "location": error.location,
                "description": error.description,
                "suggestion": error.suggestion,
                "severity": error.severity
            })
        
        return FormulaResponse(
            success=True,
            formula=result["corrected_formula"],
            explanation=f"Analysis completed. Found {len(errors)} issues.",
            errors=errors,
            solutions=result["solutions"],
            optimizations=result["optimizations"]
        )
        
    except Exception as e:
        logger.error(f"Formula debugging failed: {e}")
        return FormulaResponse(
            success=False,
            explanation=f"Error debugging formula: {str(e)}"
        )

@router.post("/formula/explain", response_model=FormulaResponse)
async def explain_formula(
    formula: str,
    generator: FormulaGenerator = Depends(get_formula_generator)
):
    """
    Explain how an Excel formula works
    """
    logger.info(f"Explaining formula: {formula}")
    
    try:
        # Use the AI service to explain the formula
        explanation = await generator._generate_explanation(formula, f"Explain: {formula}", {})
        examples = await generator._generate_examples(formula)
        
        return FormulaResponse(
            success=True,
            formula=formula,
            explanation=explanation,
            examples=examples
        )
        
    except Exception as e:
        logger.error(f"Formula explanation failed: {e}")
        return FormulaResponse(
            success=False,
            explanation=f"Error explaining formula: {str(e)}"
        )

# VBA Macro Endpoints
@router.post("/vba/generate", response_model=VBAResponse)
async def generate_vba_macro(
    request: VBARequest,
    vba_engine: VBAMacroEngine = Depends(get_vba_engine)
):
    """
    Generate VBA macro from natural language requirement
    """
    logger.info(f"Generating VBA macro for: {request.requirement[:100]}...")
    
    try:
        macro = await vba_engine.generate_macro(
            requirement=request.requirement,
            context=request.context,
            constraints=request.constraints
        )
        
        return VBAResponse(
            success=True,
            name=macro.name,
            code=macro.code,
            description=macro.description,
            category=macro.category.value,
            complexity=macro.complexity.value,
            security_level=macro.security_level.value,
            dependencies=macro.dependencies,
            parameters=macro.parameters,
            usage_examples=macro.usage_examples,
            warnings=macro.warnings,
            performance_notes=macro.performance_notes
        )
        
    except Exception as e:
        logger.error(f"VBA generation failed: {e}")
        return VBAResponse(
            success=False,
            description=f"Error generating VBA macro: {str(e)}"
        )

@router.post("/vba/analyze", response_model=VBAResponse)
async def analyze_vba_macro(
    request: VBAAnalysisRequest,
    vba_engine: VBAMacroEngine = Depends(get_vba_engine)
):
    """
    Analyze existing VBA macro code
    """
    logger.info(f"Analyzing VBA code ({len(request.code)} characters)")
    
    try:
        analysis = await vba_engine.analyze_macro(
            code=request.code,
            context=request.context
        )
        
        analysis_dict = {
            "security_issues": analysis.security_issues,
            "performance_issues": analysis.performance_issues,
            "best_practices": analysis.best_practices,
            "optimization_suggestions": analysis.optimization_suggestions,
            "code_quality_score": analysis.code_quality_score,
            "maintainability_score": analysis.maintainability_score,
            "overall_rating": analysis.overall_rating
        }
        
        return VBAResponse(
            success=True,
            code=request.code,
            description=f"Code analysis complete. Overall rating: {analysis.overall_rating}",
            analysis=analysis_dict,
            warnings=analysis.security_issues if analysis.security_issues else []
        )
        
    except Exception as e:
        logger.error(f"VBA analysis failed: {e}")
        return VBAResponse(
            success=False,
            description=f"Error analyzing VBA code: {str(e)}"
        )

@router.post("/vba/optimize", response_model=VBAResponse)
async def optimize_vba_macro(
    code: str,
    optimization_goals: Optional[List[str]] = None,
    vba_engine: VBAMacroEngine = Depends(get_vba_engine)
):
    """
    Optimize existing VBA macro code
    """
    logger.info(f"Optimizing VBA code ({len(code)} characters)")
    
    try:
        optimized_code = await vba_engine.optimize_macro(
            code=code,
            optimization_goals=optimization_goals
        )
        
        return VBAResponse(
            success=True,
            code=optimized_code,
            description="VBA code has been optimized for better performance and best practices."
        )
        
    except Exception as e:
        logger.error(f"VBA optimization failed: {e}")
        return VBAResponse(
            success=False,
            description=f"Error optimizing VBA code: {str(e)}"
        )

# Utility Endpoints
@router.get("/formula/functions", response_model=Dict[str, Any])
async def get_excel_functions(
    category: Optional[str] = None,
    generator: FormulaGenerator = Depends(get_formula_generator)
):
    """
    Get list of available Excel functions
    """
    try:
        functions = generator.function_library
        
        if category:
            # Filter by category
            filtered = {}
            for func_name, func_info in functions.items():
                if func_info.get("category", "").value.lower() == category.lower():
                    filtered[func_name] = func_info
            functions = filtered
        
        return {
            "success": True,
            "functions": functions,
            "count": len(functions)
        }
        
    except Exception as e:
        logger.error(f"Error retrieving functions: {e}")
        return {
            "success": False,
            "error": str(e)
        }

@router.get("/vba/templates", response_model=Dict[str, Any])
async def get_vba_templates(
    category: Optional[str] = None,
    vba_engine: VBAMacroEngine = Depends(get_vba_engine)
):
    """
    Get available VBA macro templates
    """
    try:
        templates = vba_engine.vba_templates
        
        if category and category in templates:
            templates = {category: templates[category]}
        
        return {
            "success": True,
            "templates": templates,
            "categories": list(templates.keys())
        }
        
    except Exception as e:
        logger.error(f"Error retrieving VBA templates: {e}")
        return {
            "success": False,
            "error": str(e)
        }

@router.get("/health/engines", response_model=Dict[str, Any])
async def check_engines_health():
    """
    Check health of formula and VBA engines
    """
    try:
        ai_service = await get_ai_interface()
        ai_health = await ai_service.health_check()
        
        return {
            "formula_engine": {
                "status": "healthy" if ai_health["status"] == "healthy" else "degraded",
                "ai_service_available": ai_health["status"] == "healthy"
            },
            "vba_engine": {
                "status": "healthy" if ai_health["status"] == "healthy" else "degraded", 
                "ai_service_available": ai_health["status"] == "healthy"
            },
            "overall_status": "healthy" if ai_health["status"] == "healthy" else "degraded",
            "timestamp": time.time()
        }
        
    except Exception as e:
        logger.error(f"Health check failed: {e}")
        return {
            "formula_engine": {"status": "error", "error": str(e)},
            "vba_engine": {"status": "error", "error": str(e)},
            "overall_status": "error",
            "timestamp": time.time()
        }