"""
VBA Macro Automation Engine for Manice Excel AI Copilot
Provides AI-powered VBA macro generation, analysis, and optimization
"""

from typing import Dict, List, Optional, Any, Union
import re
import json
from dataclasses import dataclass
from enum import Enum
from loguru import logger

class MacroComplexity(Enum):
    SIMPLE = "simple"
    MODERATE = "moderate"
    COMPLEX = "complex"
    ADVANCED = "advanced"

class MacroCategory(Enum):
    DATA_MANIPULATION = "data_manipulation"
    FORMATTING = "formatting"
    AUTOMATION = "automation"
    REPORTING = "reporting"
    USER_INTERFACE = "user_interface"
    FILE_OPERATIONS = "file_operations"
    CALCULATIONS = "calculations"
    CHARTS_GRAPHS = "charts_graphs"
    VALIDATION = "validation"
    INTEGRATION = "integration"

class SecurityLevel(Enum):
    SAFE = "safe"
    MODERATE = "moderate"
    HIGH_RISK = "high_risk"
    DANGEROUS = "dangerous"

@dataclass
class VBAMacro:
    name: str
    code: str
    description: str
    category: MacroCategory
    complexity: MacroComplexity
    security_level: SecurityLevel
    dependencies: List[str]
    parameters: List[Dict[str, str]]
    usage_examples: List[str]
    warnings: List[str]
    performance_notes: List[str]

@dataclass
class MacroAnalysis:
    security_issues: List[str]
    performance_issues: List[str]
    best_practices: List[str]
    optimization_suggestions: List[str]
    code_quality_score: float
    maintainability_score: float
    overall_rating: str

class VBAMacroEngine:
    """AI-powered VBA macro generation and analysis"""
    
    def __init__(self, ai_service):
        self.ai_service = ai_service
        self.vba_templates = self._load_vba_templates()
        self.security_patterns = self._load_security_patterns()
        self.best_practices = self._load_best_practices()
        
    def _load_vba_templates(self) -> Dict:
        """Load common VBA macro templates"""
        return {
            "data_manipulation": {
                "sort_data": '''
Sub SortData()
    Dim ws As Worksheet
    Dim dataRange As Range
    
    Set ws = ActiveSheet
    Set dataRange = ws.UsedRange
    
    dataRange.Sort Key1:=dataRange.Columns(1), Order1:=xlAscending, Header:=xlYes
End Sub
                ''',
                "filter_data": '''
Sub FilterData(criteria As String)
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ws.Range("A1").AutoFilter Field:=1, Criteria1:=criteria
End Sub
                ''',
                "remove_duplicates": '''
Sub RemoveDuplicates()
    Dim ws As Worksheet
    Dim dataRange As Range
    
    Set ws = ActiveSheet
    Set dataRange = ws.UsedRange
    
    dataRange.RemoveDuplicates Columns:=1, Header:=xlYes
End Sub
                '''
            },
            "formatting": {
                "format_table": '''
Sub FormatTable()
    Dim ws As Worksheet
    Dim tbl As ListObject
    
    Set ws = ActiveSheet
    Set tbl = ws.ListObjects.Add(xlSrcRange, ws.UsedRange, , xlYes)
    
    tbl.TableStyle = "TableStyleMedium2"
End Sub
                ''',
                "conditional_formatting": '''
Sub ApplyConditionalFormatting()
    Dim ws As Worksheet
    Dim rng As Range
    
    Set ws = ActiveSheet
    Set rng = ws.Range("A1:Z1000")
    
    rng.FormatConditions.AddColorScale ColorScaleType:=3
End Sub
                '''
            },
            "automation": {
                "copy_sheets": '''
Sub CopySheets()
    Dim srcWb As Workbook
    Dim destWb As Workbook
    Dim ws As Worksheet
    
    Set srcWb = ThisWorkbook
    Set destWb = Workbooks.Add
    
    For Each ws In srcWb.Worksheets
        ws.Copy After:=destWb.Sheets(destWb.Sheets.Count)
    Next ws
End Sub
                ''',
                "email_report": '''
Sub EmailReport()
    Dim OutApp As Object
    Dim OutMail As Object
    
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    
    With OutMail
        .To = "recipient@example.com"
        .Subject = "Automated Report"
        .Body = "Please find the attached report."
        .Attachments.Add ThisWorkbook.FullName
        .Send
    End With
End Sub
                '''
            },
            "file_operations": {
                "save_as_pdf": '''
Sub SaveAsPDF()
    Dim ws As Worksheet
    Dim fileName As String
    
    Set ws = ActiveSheet
    fileName = ThisWorkbook.Path & "\\" & ws.Name & ".pdf"
    
    ws.ExportAsFixedFormat Type:=xlTypePDF, fileName:=fileName
End Sub
                ''',
                "import_csv": '''
Sub ImportCSV(filePath As String)
    Dim ws As Worksheet
    Dim qt As QueryTable
    
    Set ws = ActiveSheet
    Set qt = ws.QueryTables.Add(Connection:="TEXT;" & filePath, Destination:=ws.Range("A1"))
    
    With qt
        .TextFileParseType = xlDelimited
        .TextFileCommaDelimiter = True
        .Refresh
    End With
End Sub
                '''
            },
            "reporting": {
                "create_summary": '''
Sub CreateSummary()
    Dim ws As Worksheet
    Dim summaryWs As Worksheet
    Dim dataRange As Range
    
    Set ws = ActiveSheet
    Set summaryWs = ThisWorkbook.Worksheets.Add
    summaryWs.Name = "Summary"
    
    Set dataRange = ws.UsedRange
    
    summaryWs.Range("A1").Value = "Total Records: " & dataRange.Rows.Count - 1
    summaryWs.Range("A2").Value = "Average: " & Application.WorksheetFunction.Average(dataRange.Columns(2))
End Sub
                ''',
                "pivot_table": '''
Sub CreatePivotTable()
    Dim ws As Worksheet
    Dim pivotWs As Worksheet
    Dim pivotCache As PivotCache
    Dim pivotTable As PivotTable
    
    Set ws = ActiveSheet
    Set pivotWs = ThisWorkbook.Worksheets.Add
    
    Set pivotCache = ThisWorkbook.PivotCaches.Create(xlDatabase, ws.UsedRange)
    Set pivotTable = pivotCache.CreatePivotTable(pivotWs.Range("A1"))
End Sub
                '''
            }
        }
    
    def _load_security_patterns(self) -> Dict:
        """Load security patterns and risks"""
        return {
            "dangerous_functions": [
                "Shell", "CreateObject", "GetObject", "Environ", 
                "Dir", "Kill", "RmDir", "MkDir", "ChDir", "ChDrive",
                "FileSystem", "Scripting.FileSystemObject"
            ],
            "risky_patterns": [
                r"Application\.EnableEvents\s*=\s*False",
                r"Application\.ScreenUpdating\s*=\s*False",
                r"\.Execute\s*\(",
                r"\.Run\s*\(",
                r"SendKeys",
                r"DoEvents",
                r"Application\.Wait",
                r"Sleep"
            ],
            "file_operations": [
                r"Open\s+.+For\s+(Input|Output|Append)",
                r"\.OpenText",
                r"\.SaveAs",
                r"\.Delete",
                r"\.Move",
                r"\.Copy"
            ],
            "external_access": [
                r"CreateObject\(.*(Excel|Word|PowerPoint|Outlook).*\)",
                r"CreateObject\(.*(Internet|Http|XML).*\)",
                r"CreateObject\(.*(Shell|WScript).*\)"
            ]
        }
    
    def _load_best_practices(self) -> Dict:
        """Load VBA best practices"""
        return {
            "variable_declaration": [
                "Always use Option Explicit",
                "Declare variables with specific types",
                "Use meaningful variable names",
                "Initialize variables properly"
            ],
            "error_handling": [
                "Always include error handling",
                "Use On Error GoTo for critical sections",
                "Clean up objects and resources",
                "Provide meaningful error messages"
            ],
            "performance": [
                "Turn off screen updating for long operations",
                "Disable automatic calculations when needed",
                "Use arrays for bulk data operations",
                "Avoid selecting ranges unnecessarily"
            ],
            "maintainability": [
                "Break complex procedures into smaller functions",
                "Use comments to explain complex logic",
                "Follow consistent naming conventions",
                "Avoid hard-coded values"
            ]
        }
    
    async def generate_macro(self, 
                           requirement: str, 
                           context: Optional[Dict] = None,
                           constraints: Optional[Dict] = None) -> VBAMacro:
        """Generate VBA macro based on natural language requirement"""
        try:
            # Analyze the requirement
            analysis = await self._analyze_requirement(requirement, context)
            
            # Determine category and complexity
            category = self._determine_category(requirement, analysis)
            complexity = self._determine_complexity(requirement, analysis)
            
            # Generate macro code
            code = await self._generate_macro_code(requirement, analysis, category, complexity, constraints)
            
            # Analyze security implications
            security_level = self._assess_security_level(code)
            
            # Generate metadata
            name = self._generate_macro_name(requirement, analysis)
            description = await self._generate_description(requirement, code, analysis)
            dependencies = self._extract_dependencies(code)
            parameters = self._extract_parameters(code)
            examples = await self._generate_usage_examples(code, requirement)
            warnings = self._generate_warnings(code, security_level)
            performance_notes = self._generate_performance_notes(code, complexity)
            
            return VBAMacro(
                name=name,
                code=code,
                description=description,
                category=category,
                complexity=complexity,
                security_level=security_level,
                dependencies=dependencies,
                parameters=parameters,
                usage_examples=examples,
                warnings=warnings,
                performance_notes=performance_notes
            )
            
        except Exception as e:
            logger.error(f"Error generating macro: {str(e)}")
            raise
    
    async def _analyze_requirement(self, requirement: str, context: Optional[Dict] = None) -> Dict:
        """Analyze natural language requirement for macro generation"""
        prompt = f"""
        Analyze this VBA macro requirement and extract key information:
        
        Requirement: "{requirement}"
        Context: {json.dumps(context, indent=2) if context else "None"}
        
        Extract:
        1. Main action/operation
        2. Data sources and targets
        3. User interface requirements
        4. Automation level needed
        5. Error handling needs
        6. Performance considerations
        7. Security implications
        8. External dependencies
        
        Return as JSON.
        """
        
        try:
            response = await self.ai_service.generate_response(prompt, max_tokens=500)
            return json.loads(response)
        except Exception as e:
            logger.warning(f"Could not analyze requirement with AI: {str(e)}")
            return self._basic_requirement_analysis(requirement)
    
    def _basic_requirement_analysis(self, requirement: str) -> Dict:
        """Basic keyword-based requirement analysis"""
        requirement_lower = requirement.lower()
        
        operations = []
        if any(word in requirement_lower for word in ['sort', 'order', 'arrange']):
            operations.append('sort')
        if any(word in requirement_lower for word in ['filter', 'search', 'find']):
            operations.append('filter')
        if any(word in requirement_lower for word in ['format', 'style', 'color']):
            operations.append('format')
        if any(word in requirement_lower for word in ['copy', 'duplicate', 'clone']):
            operations.append('copy')
        if any(word in requirement_lower for word in ['delete', 'remove', 'clear']):
            operations.append('delete')
        if any(word in requirement_lower for word in ['email', 'send', 'mail']):
            operations.append('email')
        if any(word in requirement_lower for word in ['save', 'export', 'output']):
            operations.append('save')
        if any(word in requirement_lower for word in ['chart', 'graph', 'plot']):
            operations.append('chart')
        
        return {
            "main_operation": operations[0] if operations else "unknown",
            "operations": operations,
            "has_user_input": any(word in requirement_lower for word in ['input', 'prompt', 'ask']),
            "has_file_operations": any(word in requirement_lower for word in ['file', 'save', 'open', 'import', 'export']),
            "has_external_access": any(word in requirement_lower for word in ['email', 'internet', 'web', 'api']),
            "complexity_indicators": len(operations)
        }
    
    def _determine_category(self, requirement: str, analysis: Dict) -> MacroCategory:
        """Determine macro category based on requirement"""
        main_op = analysis.get("main_operation", "").lower()
        
        category_mapping = {
            "sort": MacroCategory.DATA_MANIPULATION,
            "filter": MacroCategory.DATA_MANIPULATION,
            "format": MacroCategory.FORMATTING,
            "copy": MacroCategory.AUTOMATION,
            "delete": MacroCategory.DATA_MANIPULATION,
            "email": MacroCategory.INTEGRATION,
            "save": MacroCategory.FILE_OPERATIONS,
            "chart": MacroCategory.CHARTS_GRAPHS
        }
        
        return category_mapping.get(main_op, MacroCategory.AUTOMATION)
    
    def _determine_complexity(self, requirement: str, analysis: Dict) -> MacroComplexity:
        """Determine macro complexity"""
        complexity_score = 0
        
        # Multiple operations
        operations = analysis.get("operations", [])
        complexity_score += len(operations)
        
        # External dependencies
        if analysis.get("has_external_access"):
            complexity_score += 2
        
        # File operations
        if analysis.get("has_file_operations"):
            complexity_score += 1
        
        # User interaction
        if analysis.get("has_user_input"):
            complexity_score += 1
        
        # Complex keywords
        complex_keywords = ['loop', 'condition', 'if', 'array', 'dictionary', 'class']
        if any(word in requirement.lower() for word in complex_keywords):
            complexity_score += 2
        
        if complexity_score <= 1:
            return MacroComplexity.SIMPLE
        elif complexity_score <= 3:
            return MacroComplexity.MODERATE
        elif complexity_score <= 5:
            return MacroComplexity.COMPLEX
        else:
            return MacroComplexity.ADVANCED
    
    async def _generate_macro_code(self, 
                                 requirement: str, 
                                 analysis: Dict, 
                                 category: MacroCategory, 
                                 complexity: MacroComplexity,
                                 constraints: Optional[Dict] = None) -> str:
        """Generate VBA macro code using AI"""
        
        # Get relevant templates
        relevant_templates = self._get_relevant_templates(category, analysis)
        
        prompt = f"""
        Generate a VBA macro for this requirement:
        
        Requirement: "{requirement}"
        Category: {category.value}
        Complexity: {complexity.value}
        Analysis: {json.dumps(analysis, indent=2)}
        Constraints: {json.dumps(constraints, indent=2) if constraints else "None"}
        
        Relevant Templates:
        {json.dumps(relevant_templates, indent=2)}
        
        Best Practices:
        {json.dumps(self.best_practices, indent=2)}
        
        Requirements:
        1. Generate complete, working VBA code
        2. Include proper error handling
        3. Use meaningful variable names
        4. Add comments for complex logic
        5. Follow VBA best practices
        6. Include Option Explicit
        7. Handle edge cases
        8. Optimize for performance
        
        Return only the VBA code with proper formatting.
        """
        
        try:
            code = await self.ai_service.generate_response(prompt, max_tokens=800)
            code = self._clean_generated_code(code)
            return code
        except Exception as e:
            logger.error(f"AI macro generation failed: {str(e)}")
            return self._generate_fallback_macro(requirement, analysis, category)
    
    def _get_relevant_templates(self, category: MacroCategory, analysis: Dict) -> Dict:
        """Get relevant macro templates for the category"""
        category_map = {
            MacroCategory.DATA_MANIPULATION: "data_manipulation",
            MacroCategory.FORMATTING: "formatting",
            MacroCategory.AUTOMATION: "automation",
            MacroCategory.FILE_OPERATIONS: "file_operations",
            MacroCategory.REPORTING: "reporting"
        }
        
        template_key = category_map.get(category, "automation")
        return self.vba_templates.get(template_key, {})
    
    def _clean_generated_code(self, code: str) -> str:
        """Clean and format generated VBA code"""
        # Remove markdown formatting if present
        code = re.sub(r'```vba\n?', '', code)
        code = re.sub(r'```\n?', '', code)
        
        # Ensure proper indentation
        lines = code.split('\n')
        cleaned_lines = []
        indent_level = 0
        
        for line in lines:
            line = line.strip()
            if not line:
                cleaned_lines.append('')
                continue
            
            # Decrease indent for End statements
            if line.startswith(('End ', 'Next', 'Loop', 'Wend')):
                indent_level = max(0, indent_level - 1)
            
            # Add appropriate indentation
            cleaned_lines.append('    ' * indent_level + line)
            
            # Increase indent for structure statements
            if any(line.startswith(stmt) for stmt in [
                'Sub ', 'Function ', 'If ', 'For ', 'Do ', 'While ', 
                'With ', 'Select Case', 'Try'
            ]):
                indent_level += 1
        
        return '\n'.join(cleaned_lines)
    
    def _generate_fallback_macro(self, requirement: str, analysis: Dict, category: MacroCategory) -> str:
        """Generate basic fallback macro when AI fails"""
        main_op = analysis.get("main_operation", "")
        
        if main_op == "sort":
            return '''Option Explicit

Sub SortData()
    Dim ws As Worksheet
    Dim dataRange As Range
    
    On Error GoTo ErrorHandler
    
    Set ws = ActiveSheet
    Set dataRange = ws.UsedRange
    
    If dataRange.Rows.Count > 1 Then
        dataRange.Sort Key1:=dataRange.Columns(1), Order1:=xlAscending, Header:=xlYes
        MsgBox "Data sorted successfully!"
    Else
        MsgBox "No data to sort!"
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error sorting data: " & Err.Description
End Sub'''
        else:
            return '''Option Explicit

Sub AutoGeneratedMacro()
    ' This macro was auto-generated based on your requirements
    ' Please review and modify as needed
    
    On Error GoTo ErrorHandler
    
    MsgBox "Macro executed successfully!"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description
End Sub'''
    
    def _assess_security_level(self, code: str) -> SecurityLevel:
        """Assess security level of generated macro"""
        security_score = 0
        
        # Check for dangerous functions
        for func in self.security_patterns["dangerous_functions"]:
            if func in code:
                security_score += 3
        
        # Check for risky patterns
        for pattern in self.security_patterns["risky_patterns"]:
            if re.search(pattern, code, re.IGNORECASE):
                security_score += 2
        
        # Check for file operations
        for pattern in self.security_patterns["file_operations"]:
            if re.search(pattern, code, re.IGNORECASE):
                security_score += 1
        
        # Check for external access
        for pattern in self.security_patterns["external_access"]:
            if re.search(pattern, code, re.IGNORECASE):
                security_score += 2
        
        if security_score == 0:
            return SecurityLevel.SAFE
        elif security_score <= 2:
            return SecurityLevel.MODERATE
        elif security_score <= 5:
            return SecurityLevel.HIGH_RISK
        else:
            return SecurityLevel.DANGEROUS
    
    def _generate_macro_name(self, requirement: str, analysis: Dict) -> str:
        """Generate appropriate macro name"""
        main_op = analysis.get("main_operation", "")
        
        # Create name based on main operation
        name_map = {
            "sort": "SortData",
            "filter": "FilterData",
            "format": "FormatCells",
            "copy": "CopyData",
            "delete": "DeleteData",
            "email": "SendEmail",
            "save": "SaveData",
            "chart": "CreateChart"
        }
        
        base_name = name_map.get(main_op, "AutoGeneratedMacro")
        
        # Add suffix if multiple operations
        operations = analysis.get("operations", [])
        if len(operations) > 1:
            base_name += "Advanced"
        
        return base_name
    
    async def _generate_description(self, requirement: str, code: str, analysis: Dict) -> str:
        """Generate macro description"""
        prompt = f"""
        Generate a clear, concise description for this VBA macro:
        
        Original Requirement: "{requirement}"
        Generated Code: {code[:500]}...
        Analysis: {json.dumps(analysis, indent=2)}
        
        Provide a description that explains:
        1. What the macro does
        2. Key functionality
        3. Prerequisites or requirements
        4. Expected outcome
        
        Keep it user-friendly and professional.
        """
        
        try:
            return await self.ai_service.generate_response(prompt, max_tokens=200)
        except Exception as e:
            logger.warning(f"Could not generate description: {str(e)}")
            return f"This macro automates the requested task: {requirement}"
    
    def _extract_dependencies(self, code: str) -> List[str]:
        """Extract dependencies from macro code"""
        dependencies = []
        
        # Check for external objects
        if "CreateObject" in code:
            objects = re.findall(r'CreateObject\("([^"]+)"\)', code)
            dependencies.extend(objects)
        
        # Check for references
        if "References" in code:
            dependencies.append("Additional References Required")
        
        # Check for add-ins
        if "AddIns" in code:
            dependencies.append("Excel Add-ins")
        
        return list(set(dependencies))
    
    def _extract_parameters(self, code: str) -> List[Dict[str, str]]:
        """Extract parameters from macro code"""
        parameters = []
        
        # Find Sub and Function declarations
        patterns = [
            r'Sub\s+\w+\s*\(([^)]+)\)',
            r'Function\s+\w+\s*\(([^)]+)\)'
        ]
        
        for pattern in patterns:
            matches = re.findall(pattern, code)
            for match in matches:
                if match.strip():
                    params = [p.strip() for p in match.split(',')]
                    for param in params:
                        parts = param.split(' As ')
                        param_name = parts[0].strip()
                        param_type = parts[1].strip() if len(parts) > 1 else "Variant"
                        
                        parameters.append({
                            "name": param_name,
                            "type": param_type,
                            "description": f"Parameter {param_name} of type {param_type}"
                        })
        
        return parameters
    
    async def _generate_usage_examples(self, code: str, requirement: str) -> List[str]:
        """Generate usage examples for the macro"""
        prompt = f"""
        Generate 2-3 practical usage examples for this VBA macro:
        
        Macro Code: {code[:300]}...
        Original Requirement: "{requirement}"
        
        For each example, provide:
        - Scenario description
        - How to run the macro
        - Expected result
        
        Keep examples realistic and helpful.
        """
        
        try:
            response = await self.ai_service.generate_response(prompt, max_tokens=300)
            examples = [line.strip() for line in response.split('\n') if line.strip()]
            return examples[:3]
        except Exception as e:
            logger.warning(f"Could not generate examples: {str(e)}")
            return ["Run this macro from the VBA editor or assign it to a button"]
    
    def _generate_warnings(self, code: str, security_level: SecurityLevel) -> List[str]:
        """Generate warnings based on code analysis"""
        warnings = []
        
        if security_level == SecurityLevel.DANGEROUS:
            warnings.append("⚠️ DANGEROUS: This macro contains potentially harmful operations")
        elif security_level == SecurityLevel.HIGH_RISK:
            warnings.append("⚠️ HIGH RISK: Review this macro carefully before running")
        elif security_level == SecurityLevel.MODERATE:
            warnings.append("⚠️ MODERATE RISK: This macro performs system operations")
        
        # Check for specific risks
        if "Shell" in code:
            warnings.append("This macro executes external programs")
        if "CreateObject" in code:
            warnings.append("This macro creates external objects")
        if "FileSystem" in code:
            warnings.append("This macro modifies files or folders")
        if "EnableEvents" in code and "False" in code:
            warnings.append("This macro disables Excel events")
        
        return warnings
    
    def _generate_performance_notes(self, code: str, complexity: MacroComplexity) -> List[str]:
        """Generate performance notes"""
        notes = []
        
        if complexity in [MacroComplexity.COMPLEX, MacroComplexity.ADVANCED]:
            notes.append("Complex macro - may take longer to execute")
        
        if "ScreenUpdating" in code and "False" in code:
            notes.append("Screen updating disabled for better performance")
        
        if "UsedRange" in code:
            notes.append("Performance depends on data size")
        
        if "Loop" in code or "For" in code:
            notes.append("Contains loops - execution time varies with data")
        
        if "AutoFilter" in code:
            notes.append("Uses AutoFilter - ensure data is properly formatted")
        
        return notes

    async def analyze_macro(self, code: str, context: Optional[Dict] = None) -> MacroAnalysis:
        """Analyze existing VBA macro code"""
        try:
            # Security analysis
            security_issues = await self._analyze_security(code)
            
            # Performance analysis
            performance_issues = await self._analyze_performance(code)
            
            # Best practices check
            best_practices = await self._check_best_practices(code)
            
            # Optimization suggestions
            optimizations = await self._suggest_optimizations(code)
            
            # Quality scores
            code_quality = self._calculate_code_quality(code)
            maintainability = self._calculate_maintainability(code)
            
            # Overall rating
            overall_rating = self._calculate_overall_rating(code_quality, maintainability, len(security_issues))
            
            return MacroAnalysis(
                security_issues=security_issues,
                performance_issues=performance_issues,
                best_practices=best_practices,
                optimization_suggestions=optimizations,
                code_quality_score=code_quality,
                maintainability_score=maintainability,
                overall_rating=overall_rating
            )
            
        except Exception as e:
            logger.error(f"Error analyzing macro: {str(e)}")
            raise
    
    async def _analyze_security(self, code: str) -> List[str]:
        """Analyze security issues in macro code"""
        issues = []
        
        # Check dangerous functions
        for func in self.security_patterns["dangerous_functions"]:
            if func in code:
                issues.append(f"Uses potentially dangerous function: {func}")
        
        # Check risky patterns
        for pattern in self.security_patterns["risky_patterns"]:
            if re.search(pattern, code, re.IGNORECASE):
                issues.append(f"Contains risky pattern: {pattern}")
        
        # Use AI for advanced analysis
        try:
            prompt = f"""
            Analyze this VBA code for security vulnerabilities:
            
            Code: {code[:1000]}...
            
            Look for:
            1. File system access
            2. External program execution
            3. Network access
            4. Registry modifications
            5. Privilege escalation
            6. Data exposure risks
            
            Return list of security concerns.
            """
            
            response = await self.ai_service.generate_response(prompt, max_tokens=300)
            ai_issues = [line.strip() for line in response.split('\n') if line.strip()]
            issues.extend(ai_issues)
        except Exception as e:
            logger.warning(f"AI security analysis failed: {str(e)}")
        
        return list(set(issues))
    
    async def _analyze_performance(self, code: str) -> List[str]:
        """Analyze performance issues in macro code"""
        issues = []
        
        # Check common performance problems
        if "Select" in code and "Range" in code:
            issues.append("Unnecessary range selection - can slow down execution")
        
        if "ScreenUpdating" not in code and ("Loop" in code or "For" in code):
            issues.append("Consider disabling screen updating for better performance")
        
        if "Calculation" not in code and "Formula" in code:
            issues.append("Consider disabling automatic calculation during formula operations")
        
        if code.count("Range(") > 10:
            issues.append("Multiple Range() calls - consider using variables")
        
        return issues
    
    async def _check_best_practices(self, code: str) -> List[str]:
        """Check adherence to VBA best practices"""
        practices = []
        
        if "Option Explicit" not in code:
            practices.append("Add 'Option Explicit' for better variable management")
        
        if "On Error" not in code:
            practices.append("Add error handling with 'On Error' statements")
        
        if not re.search(r'Dim\s+\w+\s+As\s+\w+', code):
            practices.append("Use explicit variable declarations with data types")
        
        # Check for comments
        comment_lines = len([line for line in code.split('\n') if line.strip().startswith("'")])
        code_lines = len([line for line in code.split('\n') if line.strip() and not line.strip().startswith("'")])
        
        if code_lines > 20 and comment_lines / code_lines < 0.1:
            practices.append("Add more comments to explain complex logic")
        
        return practices
    
    async def _suggest_optimizations(self, code: str) -> List[str]:
        """Suggest code optimizations"""
        suggestions = []
        
        # Performance optimizations
        if "UsedRange" in code:
            suggestions.append("Consider using specific ranges instead of UsedRange for better performance")
        
        if "Cells(i," in code:
            suggestions.append("Use arrays or ranges instead of cell-by-cell operations")
        
        # Modern alternatives
        if "VLOOKUP" in code:
            suggestions.append("Consider using XLOOKUP or Index/Match for more robust lookups")
        
        # Code structure
        if code.count("Sub ") + code.count("Function ") == 1 and len(code.split('\n')) > 50:
            suggestions.append("Consider breaking large procedures into smaller functions")
        
        return suggestions
    
    def _calculate_code_quality(self, code: str) -> float:
        """Calculate code quality score (0-1)"""
        score = 0.5  # Base score
        
        # Positive factors
        if "Option Explicit" in code:
            score += 0.1
        if "On Error" in code:
            score += 0.1
        if re.search(r'Dim\s+\w+\s+As\s+\w+', code):
            score += 0.1
        
        # Comments ratio
        comment_lines = len([line for line in code.split('\n') if line.strip().startswith("'")])
        total_lines = len([line for line in code.split('\n') if line.strip()])
        if total_lines > 0 and comment_lines / total_lines > 0.1:
            score += 0.1
        
        # Negative factors
        if "GoTo" in code and "On Error GoTo" not in code:
            score -= 0.1
        if code.count("Select") > 3:
            score -= 0.1
        
        return max(0.0, min(1.0, score))
    
    def _calculate_maintainability(self, code: str) -> float:
        """Calculate maintainability score (0-1)"""
        score = 0.6  # Base score
        
        # Function/Sub count
        function_count = code.count("Sub ") + code.count("Function ")
        lines_per_function = len(code.split('\n')) / max(1, function_count)
        
        if lines_per_function < 30:
            score += 0.2
        elif lines_per_function > 100:
            score -= 0.2
        
        # Variable naming
        if re.search(r'Dim\s+[a-z][a-zA-Z]*\s+As', code):  # camelCase variables
            score += 0.1
        
        # Magic numbers
        magic_numbers = len(re.findall(r'\b\d+\b', code))
        if magic_numbers > 5:
            score -= 0.1
        
        return max(0.0, min(1.0, score))
    
    def _calculate_overall_rating(self, code_quality: float, maintainability: float, security_issues: int) -> str:
        """Calculate overall macro rating"""
        base_score = (code_quality + maintainability) / 2
        
        # Penalize security issues
        security_penalty = min(0.3, security_issues * 0.1)
        final_score = base_score - security_penalty
        
        if final_score >= 0.8:
            return "Excellent"
        elif final_score >= 0.6:
            return "Good"
        elif final_score >= 0.4:
            return "Fair"
        else:
            return "Poor"

    async def optimize_macro(self, code: str, optimization_goals: Optional[List[str]] = None) -> str:
        """Optimize existing VBA macro code"""
        try:
            # Analyze current code
            analysis = await self.analyze_macro(code)
            
            # Apply optimizations
            optimized_code = await self._apply_optimizations(code, analysis, optimization_goals)
            
            return optimized_code
            
        except Exception as e:
            logger.error(f"Error optimizing macro: {str(e)}")
            return code  # Return original code if optimization fails
    
    async def _apply_optimizations(self, 
                                 code: str, 
                                 analysis: MacroAnalysis,
                                 goals: Optional[List[str]] = None) -> str:
        """Apply specific optimizations to macro code"""
        
        prompt = f"""
        Optimize this VBA macro code:
        
        Original Code:
        {code}
        
        Current Analysis:
        - Code Quality: {analysis.code_quality_score}
        - Maintainability: {analysis.maintainability_score}
        - Security Issues: {len(analysis.security_issues)}
        
        Issues to Address:
        {json.dumps(analysis.performance_issues + analysis.best_practices, indent=2)}
        
        Optimization Goals: {goals if goals else "General performance and best practices"}
        
        Apply these optimizations:
        1. Improve performance
        2. Add error handling
        3. Follow VBA best practices
        4. Enhance readability
        5. Fix security issues
        6. Add proper variable declarations
        
        Return the optimized VBA code.
        """
        
        try:
            optimized = await self.ai_service.generate_response(prompt, max_tokens=1000)
            return self._clean_generated_code(optimized)
        except Exception as e:
            logger.warning(f"AI optimization failed: {str(e)}")
            return self._apply_basic_optimizations(code)
    
    def _apply_basic_optimizations(self, code: str) -> str:
        """Apply basic optimizations without AI"""
        optimized = code
        
        # Add Option Explicit if missing
        if "Option Explicit" not in optimized:
            optimized = "Option Explicit\n\n" + optimized
        
        # Add basic error handling if missing
        if "On Error" not in optimized and "Sub " in optimized:
            lines = optimized.split('\n')
            new_lines = []
            in_sub = False
            
            for line in lines:
                new_lines.append(line)
                if line.strip().startswith("Sub ") or line.strip().startswith("Function "):
                    in_sub = True
                    # Add error handling after variable declarations
                    new_lines.append("    On Error GoTo ErrorHandler")
                elif line.strip().startswith("End Sub") or line.strip().startswith("End Function"):
                    if in_sub:
                        # Insert error handler before End Sub
                        new_lines.insert(-1, "    Exit Sub")
                        new_lines.insert(-1, "")
                        new_lines.insert(-1, "ErrorHandler:")
                        new_lines.insert(-1, '    MsgBox "Error: " & Err.Description')
                    in_sub = False
            
            optimized = '\n'.join(new_lines)
        
        return optimized