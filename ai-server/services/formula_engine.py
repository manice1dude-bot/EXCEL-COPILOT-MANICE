"""
Formula Generation Engine for Manice Excel AI Copilot
Provides AI-powered formula generation, debugging, and optimization
"""

from typing import Dict, List, Optional, Any, Union
import re
import json
from dataclasses import dataclass
from enum import Enum
from loguru import logger

class FormulaComplexity(Enum):
    SIMPLE = "simple"
    MODERATE = "moderate"
    COMPLEX = "complex"
    ADVANCED = "advanced"

class FormulaCategory(Enum):
    MATH = "mathematical"
    TEXT = "text_manipulation"
    DATE = "date_time"
    LOOKUP = "lookup_reference"
    LOGICAL = "logical"
    STATISTICAL = "statistical"
    FINANCIAL = "financial"
    ARRAY = "array"
    DATABASE = "database"
    CUSTOM = "custom"

@dataclass
class FormulaResult:
    formula: str
    explanation: str
    category: FormulaCategory
    complexity: FormulaComplexity
    examples: List[str]
    prerequisites: List[str]
    alternatives: List[Dict[str, str]]
    confidence: float

@dataclass
class FormulaError:
    error_type: str
    location: Optional[str]
    description: str
    suggestion: str
    severity: str

class FormulaGenerator:
    """AI-powered Excel formula generation and analysis"""
    
    def __init__(self, ai_service):
        self.ai_service = ai_service
        self.formula_patterns = self._load_formula_patterns()
        self.function_library = self._load_function_library()
        
    def _load_formula_patterns(self) -> Dict:
        """Load common Excel formula patterns and templates"""
        return {
            "lookup": {
                "vlookup": "=VLOOKUP({lookup_value}, {table_array}, {col_index}, {range_lookup})",
                "index_match": "=INDEX({return_array}, MATCH({lookup_value}, {lookup_array}, 0))",
                "xlookup": "=XLOOKUP({lookup_value}, {lookup_array}, {return_array})"
            },
            "conditional": {
                "if_simple": "=IF({condition}, {value_if_true}, {value_if_false})",
                "if_nested": "=IF({condition1}, {value1}, IF({condition2}, {value2}, {value3}))",
                "ifs": "=IFS({condition1}, {value1}, {condition2}, {value2}, TRUE, {default})"
            },
            "aggregation": {
                "sum_if": "=SUMIF({range}, {criteria}, {sum_range})",
                "count_if": "=COUNTIF({range}, {criteria})",
                "average_if": "=AVERAGEIF({range}, {criteria}, {average_range})"
            },
            "text": {
                "concatenate": "=CONCAT({text1}, {text2})",
                "text_join": "=TEXTJOIN({delimiter}, {ignore_empty}, {text1}, {text2})",
                "substitute": "=SUBSTITUTE({text}, {old_text}, {new_text})"
            },
            "date_time": {
                "date_diff": "=DATEDIF({start_date}, {end_date}, {unit})",
                "workday": "=WORKDAY({start_date}, {days}, {holidays})",
                "today": "=TODAY()"
            },
            "array": {
                "unique": "=UNIQUE({array})",
                "sort": "=SORT({array}, {sort_index}, {sort_order})",
                "filter": "=FILTER({array}, {include})"
            }
        }
    
    def _load_function_library(self) -> Dict:
        """Load Excel function library with descriptions"""
        return {
            "SUM": {
                "description": "Adds numbers in a range",
                "syntax": "SUM(number1, [number2], ...)",
                "examples": ["=SUM(A1:A10)", "=SUM(A1, B1, C1)"],
                "category": FormulaCategory.MATH
            },
            "VLOOKUP": {
                "description": "Looks up a value in the first column and returns a value in the same row from another column",
                "syntax": "VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])",
                "examples": ["=VLOOKUP(A2, B:D, 3, FALSE)"],
                "category": FormulaCategory.LOOKUP
            },
            "INDEX": {
                "description": "Returns a value from a table based on row and column numbers",
                "syntax": "INDEX(array, row_num, [column_num])",
                "examples": ["=INDEX(A1:C10, 5, 2)"],
                "category": FormulaCategory.LOOKUP
            },
            "MATCH": {
                "description": "Searches for a specified item and returns its relative position",
                "syntax": "MATCH(lookup_value, lookup_array, [match_type])",
                "examples": ["=MATCH(\"Apple\", A1:A10, 0)"],
                "category": FormulaCategory.LOOKUP
            },
            "IF": {
                "description": "Returns one value if condition is TRUE, another if FALSE",
                "syntax": "IF(logical_test, value_if_true, [value_if_false])",
                "examples": ["=IF(A1>10, \"High\", \"Low\")"],
                "category": FormulaCategory.LOGICAL
            },
            "CONCATENATE": {
                "description": "Joins several text strings into one text string",
                "syntax": "CONCATENATE(text1, [text2], ...)",
                "examples": ["=CONCATENATE(A1, \" \", B1)"],
                "category": FormulaCategory.TEXT
            },
            "COUNTIF": {
                "description": "Counts cells that meet a criteria",
                "syntax": "COUNTIF(range, criteria)",
                "examples": ["=COUNTIF(A1:A10, \">5\")"],
                "category": FormulaCategory.STATISTICAL
            },
            "SUMIF": {
                "description": "Adds cells that meet a criteria",
                "syntax": "SUMIF(range, criteria, [sum_range])",
                "examples": ["=SUMIF(A1:A10, \">5\", B1:B10)"],
                "category": FormulaCategory.MATH
            },
            "AVERAGE": {
                "description": "Returns the average of numbers",
                "syntax": "AVERAGE(number1, [number2], ...)",
                "examples": ["=AVERAGE(A1:A10)"],
                "category": FormulaCategory.STATISTICAL
            },
            "TODAY": {
                "description": "Returns today's date",
                "syntax": "TODAY()",
                "examples": ["=TODAY()"],
                "category": FormulaCategory.DATE
            }
        }
    
    async def generate_formula(self, 
                             requirement: str, 
                             context: Optional[Dict] = None,
                             data_sample: Optional[Dict] = None) -> FormulaResult:
        """Generate Excel formula based on natural language requirement"""
        try:
            # Analyze the requirement
            analysis = await self._analyze_requirement(requirement, context)
            
            # Determine formula category and complexity
            category = self._determine_category(requirement, analysis)
            complexity = self._determine_complexity(requirement, analysis)
            
            # Generate formula using AI
            formula = await self._generate_formula_ai(requirement, analysis, category, complexity, data_sample)
            
            # Validate and optimize formula
            validated_formula = await self._validate_formula(formula)
            
            # Generate explanation and examples
            explanation = await self._generate_explanation(validated_formula, requirement, analysis)
            examples = await self._generate_examples(validated_formula, context)
            
            # Find alternatives
            alternatives = await self._find_alternatives(validated_formula, requirement, category)
            
            # Calculate confidence score
            confidence = await self._calculate_confidence(validated_formula, requirement, analysis)
            
            return FormulaResult(
                formula=validated_formula,
                explanation=explanation,
                category=category,
                complexity=complexity,
                examples=examples,
                prerequisites=self._get_prerequisites(validated_formula),
                alternatives=alternatives,
                confidence=confidence
            )
            
        except Exception as e:
            logger.error(f"Error generating formula: {str(e)}")
            raise
    
    async def _analyze_requirement(self, requirement: str, context: Optional[Dict] = None) -> Dict:
        """Analyze natural language requirement to understand intent"""
        prompt = f"""
        Analyze this Excel formula requirement and extract key information:
        
        Requirement: "{requirement}"
        Context: {json.dumps(context, indent=2) if context else "None"}
        
        Extract:
        1. Main operation (sum, lookup, count, etc.)
        2. Data ranges involved
        3. Conditions/criteria
        4. Expected output type
        5. Key keywords and Excel functions mentioned
        6. Data relationships
        7. Complexity indicators
        
        Return as JSON.
        """
        
        try:
            response = await self.ai_service.generate_response(prompt, max_tokens=500)
            return json.loads(response)
        except Exception as e:
            logger.warning(f"Could not analyze requirement with AI: {str(e)}")
            # Fallback to basic keyword analysis
            return self._basic_requirement_analysis(requirement)
    
    def _basic_requirement_analysis(self, requirement: str) -> Dict:
        """Basic keyword-based requirement analysis"""
        requirement_lower = requirement.lower()
        
        operations = []
        if any(word in requirement_lower for word in ['sum', 'add', 'total']):
            operations.append('sum')
        if any(word in requirement_lower for word in ['count', 'number of']):
            operations.append('count')
        if any(word in requirement_lower for word in ['lookup', 'find', 'search']):
            operations.append('lookup')
        if any(word in requirement_lower for word in ['average', 'mean']):
            operations.append('average')
        if any(word in requirement_lower for word in ['if', 'condition', 'when']):
            operations.append('conditional')
        
        return {
            "main_operation": operations[0] if operations else "unknown",
            "operations": operations,
            "complexity": "simple" if len(operations) <= 1 else "moderate",
            "keywords": re.findall(r'\b\w+\b', requirement_lower),
            "has_conditions": any(word in requirement_lower for word in ['if', 'when', 'where', 'condition']),
            "has_ranges": bool(re.search(r'[A-Z]\d+:[A-Z]\d+|[A-Z]:[A-Z]', requirement))
        }
    
    def _determine_category(self, requirement: str, analysis: Dict) -> FormulaCategory:
        """Determine formula category based on requirement"""
        main_op = analysis.get("main_operation", "").lower()
        
        category_mapping = {
            "sum": FormulaCategory.MATH,
            "count": FormulaCategory.STATISTICAL,
            "average": FormulaCategory.STATISTICAL,
            "lookup": FormulaCategory.LOOKUP,
            "conditional": FormulaCategory.LOGICAL,
            "text": FormulaCategory.TEXT,
            "date": FormulaCategory.DATE,
            "financial": FormulaCategory.FINANCIAL
        }
        
        return category_mapping.get(main_op, FormulaCategory.CUSTOM)
    
    def _determine_complexity(self, requirement: str, analysis: Dict) -> FormulaComplexity:
        """Determine formula complexity"""
        complexity_score = 0
        
        # Check for multiple operations
        operations = analysis.get("operations", [])
        complexity_score += len(operations)
        
        # Check for nested conditions
        if analysis.get("has_conditions") and any(word in requirement.lower() 
                                                for word in ['nested', 'multiple', 'several']):
            complexity_score += 2
        
        # Check for array operations
        if any(word in requirement.lower() for word in ['unique', 'sort', 'filter', 'array']):
            complexity_score += 2
        
        # Check for complex functions
        complex_functions = ['vlookup', 'index', 'match', 'indirect', 'offset']
        if any(func in requirement.lower() for func in complex_functions):
            complexity_score += 1
        
        if complexity_score <= 1:
            return FormulaComplexity.SIMPLE
        elif complexity_score <= 3:
            return FormulaComplexity.MODERATE
        elif complexity_score <= 5:
            return FormulaComplexity.COMPLEX
        else:
            return FormulaComplexity.ADVANCED
    
    async def _generate_formula_ai(self, 
                                 requirement: str, 
                                 analysis: Dict, 
                                 category: FormulaCategory, 
                                 complexity: FormulaComplexity,
                                 data_sample: Optional[Dict] = None) -> str:
        """Generate formula using AI with context"""
        
        # Get relevant function library entries
        relevant_functions = self._get_relevant_functions(category, analysis)
        
        prompt = f"""
        Generate an Excel formula for this requirement:
        
        Requirement: "{requirement}"
        Category: {category.value}
        Complexity: {complexity.value}
        Analysis: {json.dumps(analysis, indent=2)}
        Data Sample: {json.dumps(data_sample, indent=2) if data_sample else "None"}
        
        Relevant Excel Functions:
        {json.dumps(relevant_functions, indent=2)}
        
        Formula Patterns:
        {json.dumps(self._get_relevant_patterns(category), indent=2)}
        
        Requirements:
        1. Generate a valid Excel formula starting with =
        2. Use proper Excel function syntax
        3. Include cell references and ranges as needed
        4. Handle edge cases and errors
        5. Optimize for performance
        6. Make it as simple as possible while meeting requirements
        
        Return only the Excel formula, nothing else.
        """
        
        try:
            formula = await self.ai_service.generate_response(prompt, max_tokens=300)
            formula = formula.strip()
            
            # Ensure formula starts with =
            if not formula.startswith('='):
                formula = '=' + formula
                
            return formula
        except Exception as e:
            logger.error(f"AI formula generation failed: {str(e)}")
            # Fallback to pattern-based generation
            return self._generate_formula_pattern(requirement, analysis, category)
    
    def _get_relevant_functions(self, category: FormulaCategory, analysis: Dict) -> Dict:
        """Get relevant Excel functions for the category"""
        relevant = {}
        operations = analysis.get("operations", [])
        
        for func_name, func_info in self.function_library.items():
            if func_info["category"] == category:
                relevant[func_name] = func_info
            elif any(op in func_name.lower() for op in operations):
                relevant[func_name] = func_info
                
        return relevant
    
    def _get_relevant_patterns(self, category: FormulaCategory) -> Dict:
        """Get relevant formula patterns for the category"""
        pattern_mapping = {
            FormulaCategory.LOOKUP: self.formula_patterns.get("lookup", {}),
            FormulaCategory.LOGICAL: self.formula_patterns.get("conditional", {}),
            FormulaCategory.MATH: self.formula_patterns.get("aggregation", {}),
            FormulaCategory.TEXT: self.formula_patterns.get("text", {}),
            FormulaCategory.DATE: self.formula_patterns.get("date_time", {}),
            FormulaCategory.ARRAY: self.formula_patterns.get("array", {})
        }
        
        return pattern_mapping.get(category, {})
    
    def _generate_formula_pattern(self, requirement: str, analysis: Dict, category: FormulaCategory) -> str:
        """Fallback pattern-based formula generation"""
        main_op = analysis.get("main_operation", "")
        
        # Simple pattern matching
        if main_op == "sum":
            return "=SUM(A:A)"
        elif main_op == "count":
            return "=COUNT(A:A)"
        elif main_op == "average":
            return "=AVERAGE(A:A)"
        elif main_op == "lookup":
            return "=VLOOKUP(A2, B:D, 2, FALSE)"
        elif main_op == "conditional":
            return "=IF(A1>0, \"Positive\", \"Negative\")"
        else:
            return "=A1"
    
    async def _validate_formula(self, formula: str) -> str:
        """Validate and clean up formula"""
        # Basic validation
        if not formula.startswith('='):
            formula = '=' + formula
            
        # Check for balanced parentheses
        if formula.count('(') != formula.count(')'):
            logger.warning("Formula has unbalanced parentheses")
            
        # Check for valid function names
        functions = re.findall(r'[A-Z]+(?=\()', formula)
        invalid_functions = [f for f in functions if f not in self.function_library and f not in [
            'XLOOKUP', 'XMATCH', 'UNIQUE', 'SORT', 'FILTER', 'SEQUENCE', 'RANDARRAY'
        ]]
        
        if invalid_functions:
            logger.warning(f"Formula contains potentially invalid functions: {invalid_functions}")
        
        return formula
    
    async def _generate_explanation(self, formula: str, requirement: str, analysis: Dict) -> str:
        """Generate human-readable explanation of the formula"""
        prompt = f"""
        Explain this Excel formula in simple, clear terms:
        
        Formula: {formula}
        Original Requirement: "{requirement}"
        Analysis: {json.dumps(analysis, indent=2)}
        
        Provide:
        1. What the formula does in plain English
        2. How it works step by step
        3. What each part means
        4. When to use this approach
        
        Keep it concise but comprehensive.
        """
        
        try:
            return await self.ai_service.generate_response(prompt, max_tokens=400)
        except Exception as e:
            logger.warning(f"Could not generate explanation: {str(e)}")
            return f"This formula ({formula}) performs the requested operation based on your requirement."
    
    async def _generate_examples(self, formula: str, context: Optional[Dict] = None) -> List[str]:
        """Generate example uses of the formula"""
        prompt = f"""
        Generate 3 practical examples of how to use this Excel formula:
        
        Formula: {formula}
        Context: {json.dumps(context, indent=2) if context else "General business data"}
        
        For each example, provide:
        - Sample data scenario
        - How the formula would be applied
        - Expected result
        
        Keep examples realistic and business-relevant.
        """
        
        try:
            response = await self.ai_service.generate_response(prompt, max_tokens=300)
            # Parse response into list (simplified)
            examples = [line.strip() for line in response.split('\n') if line.strip() and not line.startswith('Example')]
            return examples[:3]  # Limit to 3 examples
        except Exception as e:
            logger.warning(f"Could not generate examples: {str(e)}")
            return [f"Use {formula} with your data ranges"]
    
    async def _find_alternatives(self, formula: str, requirement: str, category: FormulaCategory) -> List[Dict[str, str]]:
        """Find alternative formulas that achieve the same result"""
        prompt = f"""
        Suggest 2-3 alternative Excel formulas for this requirement:
        
        Current Formula: {formula}
        Requirement: "{requirement}"
        Category: {category.value}
        
        For each alternative, provide:
        - The alternative formula
        - Why it might be better/different
        - When to use it instead
        
        Focus on different approaches or newer Excel functions.
        """
        
        try:
            response = await self.ai_service.generate_response(prompt, max_tokens=400)
            # Parse response into alternatives (simplified)
            alternatives = []
            lines = response.split('\n')
            current_alt = {}
            
            for line in lines:
                line = line.strip()
                if line.startswith('='):
                    if current_alt:
                        alternatives.append(current_alt)
                    current_alt = {"formula": line, "description": ""}
                elif line and current_alt:
                    current_alt["description"] += " " + line
            
            if current_alt:
                alternatives.append(current_alt)
                
            return alternatives[:3]  # Limit to 3 alternatives
        except Exception as e:
            logger.warning(f"Could not find alternatives: {str(e)}")
            return []
    
    def _get_prerequisites(self, formula: str) -> List[str]:
        """Get prerequisites for using the formula"""
        prerequisites = []
        
        # Check for advanced functions
        if 'XLOOKUP' in formula:
            prerequisites.append("Excel 365 or Excel 2021")
        if any(func in formula for func in ['UNIQUE', 'SORT', 'FILTER']):
            prerequisites.append("Excel 365 with dynamic arrays")
        if 'LAMBDA' in formula:
            prerequisites.append("Excel 365 with LAMBDA function support")
        
        # Check for complex features
        if '[]' in formula:
            prerequisites.append("Dynamic array formulas enabled")
        if 'INDIRECT' in formula:
            prerequisites.append("Be cautious with file links and references")
        
        return prerequisites
    
    async def _calculate_confidence(self, formula: str, requirement: str, analysis: Dict) -> float:
        """Calculate confidence score for the generated formula"""
        confidence = 0.8  # Base confidence
        
        # Adjust based on analysis quality
        if analysis.get("main_operation") != "unknown":
            confidence += 0.1
        
        # Adjust based on formula complexity
        if formula.count('(') <= 2:  # Simple formula
            confidence += 0.1
        
        # Adjust based on recognized functions
        functions = re.findall(r'[A-Z]+(?=\()', formula)
        if all(func in self.function_library for func in functions):
            confidence += 0.1
        
        # Cap at 1.0
        return min(confidence, 1.0)

class FormulaDebugger:
    """AI-powered Excel formula debugging system"""
    
    def __init__(self, ai_service):
        self.ai_service = ai_service
        self.error_patterns = self._load_error_patterns()
        
    def _load_error_patterns(self) -> Dict:
        """Load common Excel error patterns and solutions"""
        return {
            "#DIV/0!": {
                "description": "Division by zero error",
                "common_causes": [
                    "Dividing by a cell containing zero",
                    "Dividing by an empty cell",
                    "Result of another formula is zero"
                ],
                "solutions": [
                    "Use IF statement to check for zero: =IF(B1=0, \"\", A1/B1)",
                    "Use IFERROR: =IFERROR(A1/B1, \"Error\")",
                    "Check data source for zero values"
                ]
            },
            "#VALUE!": {
                "description": "Wrong data type error",
                "common_causes": [
                    "Text used in mathematical operations",
                    "Incompatible data types",
                    "Invalid date/time values"
                ],
                "solutions": [
                    "Use VALUE() to convert text to numbers",
                    "Check data formatting",
                    "Use ISNUMBER() to validate data"
                ]
            },
            "#REF!": {
                "description": "Invalid cell reference",
                "common_causes": [
                    "Referenced cells were deleted",
                    "Invalid range references",
                    "Circular references"
                ],
                "solutions": [
                    "Check and update cell references",
                    "Restore deleted cells",
                    "Use INDIRECT for dynamic references"
                ]
            },
            "#NAME?": {
                "description": "Unrecognized function or name",
                "common_causes": [
                    "Misspelled function name",
                    "Missing quotes around text",
                    "Undefined named range"
                ],
                "solutions": [
                    "Check function spelling",
                    "Add quotes around text values",
                    "Define or fix named ranges"
                ]
            },
            "#N/A": {
                "description": "Value not available",
                "common_causes": [
                    "VLOOKUP/MATCH value not found",
                    "Array size mismatch",
                    "Missing data"
                ],
                "solutions": [
                    "Use IFERROR with lookup functions",
                    "Check lookup values exist",
                    "Use approximate match if appropriate"
                ]
            },
            "#NULL!": {
                "description": "Null intersection error",
                "common_causes": [
                    "Incorrect range operator",
                    "Missing comma or colon in range",
                    "Space instead of comma"
                ],
                "solutions": [
                    "Check range operators (: vs space)",
                    "Verify comma placement",
                    "Use proper range syntax"
                ]
            }
        }
    
    async def debug_formula(self, 
                          formula: str, 
                          error_message: Optional[str] = None,
                          cell_data: Optional[Dict] = None) -> Dict:
        """Debug Excel formula and provide solutions"""
        try:
            # Analyze the formula
            analysis = await self._analyze_formula_structure(formula)
            
            # Identify errors
            errors = await self._identify_errors(formula, error_message, analysis, cell_data)
            
            # Generate solutions
            solutions = await self._generate_solutions(formula, errors, analysis)
            
            # Provide optimizations
            optimizations = await self._suggest_optimizations(formula, analysis)
            
            return {
                "formula": formula,
                "analysis": analysis,
                "errors": errors,
                "solutions": solutions,
                "optimizations": optimizations,
                "corrected_formula": solutions[0]["formula"] if solutions else formula
            }
            
        except Exception as e:
            logger.error(f"Error debugging formula: {str(e)}")
            raise
    
    async def _analyze_formula_structure(self, formula: str) -> Dict:
        """Analyze formula structure and components"""
        analysis = {
            "functions": re.findall(r'[A-Z]+(?=\()', formula),
            "cell_references": re.findall(r'[A-Z]+\d+', formula),
            "range_references": re.findall(r'[A-Z]+\d+:[A-Z]+\d+', formula),
            "operators": re.findall(r'[+\-*/^&<>=]', formula),
            "parentheses_count": formula.count('('),
            "parentheses_balanced": formula.count('(') == formula.count(')'),
            "has_nested_functions": len(re.findall(r'[A-Z]+\([^)]*[A-Z]+\(', formula)) > 0,
            "complexity_score": len(re.findall(r'[A-Z]+(?=\()', formula)) + formula.count('(')
        }
        
        # Check for common patterns
        analysis["has_conditions"] = bool(re.search(r'IF\s*\(', formula))
        analysis["has_lookups"] = any(func in formula for func in ['VLOOKUP', 'HLOOKUP', 'INDEX', 'MATCH'])
        analysis["has_arrays"] = bool(re.search(r'\{.*\}', formula))
        
        return analysis
    
    async def _identify_errors(self, 
                             formula: str, 
                             error_message: Optional[str],
                             analysis: Dict,
                             cell_data: Optional[Dict] = None) -> List[FormulaError]:
        """Identify errors in the formula"""
        errors = []
        
        # Check for known error messages
        if error_message and error_message in self.error_patterns:
            error_info = self.error_patterns[error_message]
            errors.append(FormulaError(
                error_type=error_message,
                location=None,
                description=error_info["description"],
                suggestion=error_info["solutions"][0],
                severity="high"
            ))
        
        # Check structural errors
        if not analysis["parentheses_balanced"]:
            errors.append(FormulaError(
                error_type="syntax",
                location="parentheses",
                description="Unbalanced parentheses in formula",
                suggestion="Check that each opening parenthesis has a matching closing parenthesis",
                severity="high"
            ))
        
        # Check for invalid functions
        invalid_functions = [f for f in analysis["functions"] 
                           if f not in ['SUM', 'COUNT', 'AVERAGE', 'IF', 'VLOOKUP', 'INDEX', 'MATCH', 'CONCATENATE']]
        if invalid_functions:
            errors.append(FormulaError(
                error_type="function",
                location=", ".join(invalid_functions),
                description=f"Potentially invalid or unsupported functions: {', '.join(invalid_functions)}",
                suggestion="Verify function names and Excel version compatibility",
                severity="medium"
            ))
        
        # Use AI for deeper analysis
        if not errors:
            ai_errors = await self._ai_error_detection(formula, analysis, cell_data)
            errors.extend(ai_errors)
        
        return errors
    
    async def _ai_error_detection(self, 
                                formula: str, 
                                analysis: Dict,
                                cell_data: Optional[Dict] = None) -> List[FormulaError]:
        """Use AI to detect potential errors"""
        prompt = f"""
        Analyze this Excel formula for potential errors and issues:
        
        Formula: {formula}
        Structure Analysis: {json.dumps(analysis, indent=2)}
        Cell Data: {json.dumps(cell_data, indent=2) if cell_data else "Not available"}
        
        Common Excel Error Types:
        {json.dumps(list(self.error_patterns.keys()))}
        
        Look for:
        1. Syntax errors
        2. Logic errors
        3. Performance issues
        4. Data type mismatches
        5. Reference errors
        6. Edge cases
        
        Return findings as JSON array with error_type, description, suggestion, severity.
        """
        
        try:
            response = await self.ai_service.generate_response(prompt, max_tokens=500)
            ai_errors_data = json.loads(response)
            
            errors = []
            for error_data in ai_errors_data:
                errors.append(FormulaError(
                    error_type=error_data.get("error_type", "unknown"),
                    location=error_data.get("location"),
                    description=error_data.get("description", ""),
                    suggestion=error_data.get("suggestion", ""),
                    severity=error_data.get("severity", "medium")
                ))
            
            return errors
        except Exception as e:
            logger.warning(f"AI error detection failed: {str(e)}")
            return []
    
    async def _generate_solutions(self, 
                                formula: str, 
                                errors: List[FormulaError],
                                analysis: Dict) -> List[Dict]:
        """Generate solutions for identified errors"""
        if not errors:
            return [{"formula": formula, "description": "No errors found", "changes": []}]
        
        solutions = []
        
        for error in errors:
            if error.error_type in self.error_patterns:
                # Use predefined solutions
                error_info = self.error_patterns[error.error_type]
                for i, solution in enumerate(error_info["solutions"][:2]):  # Limit to 2 solutions
                    corrected_formula = await self._apply_solution(formula, error, solution)
                    solutions.append({
                        "formula": corrected_formula,
                        "description": solution,
                        "changes": [f"Fixed {error.error_type} error"],
                        "confidence": 0.8 - (i * 0.1)
                    })
            else:
                # Use AI to generate solution
                ai_solution = await self._ai_generate_solution(formula, error, analysis)
                if ai_solution:
                    solutions.append(ai_solution)
        
        return solutions
    
    async def _apply_solution(self, formula: str, error: FormulaError, solution: str) -> str:
        """Apply a solution to fix the formula"""
        # This is a simplified implementation
        # In practice, you'd need more sophisticated pattern matching and replacement
        
        if error.error_type == "#DIV/0!":
            # Wrap division operations with IFERROR
            division_pattern = r'([A-Z\d:]+)/([A-Z\d:]+)'
            if re.search(division_pattern, formula):
                formula = re.sub(division_pattern, r'IFERROR(\1/\2, "")', formula)
        
        elif error.error_type == "#N/A":
            # Wrap lookup functions with IFERROR
            if 'VLOOKUP' in formula:
                formula = re.sub(r'VLOOKUP\([^)]+\)', r'IFERROR(&, "Not Found")', formula)
        
        return formula
    
    async def _ai_generate_solution(self, 
                                  formula: str, 
                                  error: FormulaError,
                                  analysis: Dict) -> Optional[Dict]:
        """Use AI to generate solution for complex errors"""
        prompt = f"""
        Fix this Excel formula error:
        
        Original Formula: {formula}
        Error: {error.error_type} - {error.description}
        Suggestion: {error.suggestion}
        Analysis: {json.dumps(analysis, indent=2)}
        
        Provide:
        1. Corrected formula
        2. Explanation of changes made
        3. Why this solution works
        
        Return as JSON with formula, description, changes, confidence fields.
        """
        
        try:
            response = await self.ai_service.generate_response(prompt, max_tokens=400)
            solution_data = json.loads(response)
            return solution_data
        except Exception as e:
            logger.warning(f"AI solution generation failed: {str(e)}")
            return None
    
    async def _suggest_optimizations(self, formula: str, analysis: Dict) -> List[Dict]:
        """Suggest formula optimizations"""
        optimizations = []
        
        # Check for performance improvements
        if analysis["complexity_score"] > 5:
            optimizations.append({
                "type": "performance",
                "description": "Formula complexity is high, consider breaking into multiple cells",
                "suggestion": "Split complex operations into intermediate calculations"
            })
        
        # Check for newer function alternatives
        if 'VLOOKUP' in formula:
            optimizations.append({
                "type": "modern_alternative",
                "description": "Consider using XLOOKUP instead of VLOOKUP for better performance",
                "suggestion": "Replace VLOOKUP with XLOOKUP if using Excel 365"
            })
        
        # Use AI for advanced optimizations
        ai_optimizations = await self._ai_suggest_optimizations(formula, analysis)
        optimizations.extend(ai_optimizations)
        
        return optimizations
    
    async def _ai_suggest_optimizations(self, formula: str, analysis: Dict) -> List[Dict]:
        """Use AI to suggest optimizations"""
        prompt = f"""
        Suggest optimizations for this Excel formula:
        
        Formula: {formula}
        Analysis: {json.dumps(analysis, indent=2)}
        
        Look for:
        1. Performance improvements
        2. Modern function alternatives
        3. Simplification opportunities
        4. Better error handling
        5. Readability improvements
        
        Return as JSON array with type, description, suggestion fields.
        """
        
        try:
            response = await self.ai_service.generate_response(prompt, max_tokens=300)
            return json.loads(response)
        except Exception as e:
            logger.warning(f"AI optimization suggestions failed: {str(e)}")
            return []