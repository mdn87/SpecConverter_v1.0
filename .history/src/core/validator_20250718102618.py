"""
Validation and correction logic for SpecConverter v1.0

Provides validation of extracted content and automatic correction capabilities.
"""

import re
from typing import List, Dict, Any, Optional
from dataclasses import dataclass
from pathlib import Path
import json

from .models import (
    SpecDocument, 
    ContentBlock, 
    ValidationResults, 
    ExtractionError
)
from utils.logging_utils import get_logger


@dataclass
class ValidationRule:
    """Represents a validation rule"""
    name: str
    description: str
    severity: str  # "error", "warning", "info"
    pattern: Optional[str] = None
    condition: Optional[Any] = None
    correction: Optional[Any] = None


class SpecValidator:
    """Validates and corrects specification document content"""
    
    def __init__(self, config: Optional[Dict[str, Any]] = None):
        self.config = config or {}
        self.logger = get_logger("validator")
        self.validation_rules = self._setup_validation_rules()
    
    def _setup_validation_rules(self) -> List[ValidationRule]:
        """Setup validation rules"""
        rules = [
            ValidationRule(
                name="empty_content_blocks",
                description="Content blocks should not be empty",
                severity="error",
                condition=lambda block: not block.text.strip()
            ),
            ValidationRule(
                name="missing_level_numbers",
                description="Content blocks should have level numbers",
                severity="warning",
                condition=lambda block: block.level_type in ["heading", "subheading"] and not block.level_number
            ),
            ValidationRule(
                name="inconsistent_numbering",
                description="Numbering should be consistent",
                severity="warning",
                pattern=r"^\d+\.\d+",
                condition=lambda block: self._check_numbering_consistency(block)
            ),
            ValidationRule(
                name="missing_section_titles",
                description="Sections should have titles",
                severity="error",
                condition=lambda block: block.level_type == "heading" and len(block.text.strip()) < 5
            ),
            ValidationRule(
                name="duplicate_content",
                description="Duplicate content blocks detected",
                severity="warning",
                condition=lambda block: self._check_duplicate_content(block)
            )
        ]
        return rules
    
    def validate_document(self, document: SpecDocument) -> ValidationResults:
        """Validate a complete specification document"""
        self.logger.info(f"Validating document: {document.file_path}")
        
        errors = []
        corrections = []
        
        # Validate content blocks
        for i, block in enumerate(document.content_blocks):
            block_errors = self._validate_content_block(block, i)
            errors.extend(block_errors)
            
            # Apply corrections if auto-correct is enabled
            if self.config.get("auto_correct", True):
                block_corrections = self._correct_content_block(block, i)
                corrections.extend(block_corrections)
        
        # Validate document structure
        structure_errors = self._validate_document_structure(document)
        errors.extend(structure_errors)
        
        # Create validation summary
        validation_summary = self._create_validation_summary(errors, corrections)
        
        return ValidationResults(
            errors=errors,
            corrections=corrections,
            validation_summary=validation_summary
        )
    
    def _validate_content_block(self, block: ContentBlock, index: int) -> List[ExtractionError]:
        """Validate a single content block"""
        errors = []
        
        for rule in self.validation_rules:
            if rule.condition and rule.condition(block):
                error = ExtractionError(
                    line_number=index + 1,
                    error_type=rule.name,
                    message=rule.description,
                    context=block.text[:100] + "..." if len(block.text) > 100 else block.text,
                    expected=None,
                    found=block.text
                )
                errors.append(error)
        
        return errors
    
    def _correct_content_block(self, block: ContentBlock, index: int) -> List[Dict[str, Any]]:
        """Apply corrections to a content block"""
        corrections = []
        
        # Auto-correct empty content blocks
        if not block.text.strip():
            # Try to infer content from level information
            if block.level_type and block.level_number:
                inferred_text = f"{block.level_type.title()} {block.level_number}"
                block.text = inferred_text
                corrections.append({
                    "block_index": index,
                    "correction_type": "empty_content",
                    "original": "",
                    "corrected": inferred_text,
                    "description": "Inferred content from level information"
                })
        
        # Auto-correct missing level numbers
        if block.level_type in ["heading", "subheading"] and not block.level_number:
            # Try to extract number from text
            number_match = re.search(r'^(\d+)', block.text)
            if number_match:
                block.level_number = int(number_match.group(1))
                corrections.append({
                    "block_index": index,
                    "correction_type": "missing_level_number",
                    "original": None,
                    "corrected": block.level_number,
                    "description": "Extracted level number from text"
                })
        
        return corrections
    
    def _validate_document_structure(self, document: SpecDocument) -> List[ExtractionError]:
        """Validate overall document structure"""
        errors = []
        
        # Check for minimum content
        if len(document.content_blocks) < 3:
            errors.append(ExtractionError(
                line_number=0,
                error_type="insufficient_content",
                message="Document has insufficient content blocks",
                context=f"Found {len(document.content_blocks)} blocks, minimum 3 required",
                expected="3+ content blocks",
                found=str(len(document.content_blocks))
            ))
        
        # Check for proper heading hierarchy
        heading_levels = [block.level_number for block in document.content_blocks 
                         if block.level_type == "heading" and block.level_number]
        if heading_levels and max(heading_levels) - min(heading_levels) > 5:
            errors.append(ExtractionError(
                line_number=0,
                error_type="heading_hierarchy",
                message="Heading hierarchy is too deep",
                context=f"Level range: {min(heading_levels)} to {max(heading_levels)}",
                expected="Level range <= 5",
                found=f"Level range = {max(heading_levels) - min(heading_levels)}"
            ))
        
        return errors
    
    def _check_numbering_consistency(self, block: ContentBlock) -> bool:
        """Check if numbering is consistent"""
        if not block.text:
            return False
        
        # Check for common numbering patterns
        patterns = [
            r'^\d+\.\d+',  # 1.1, 2.3, etc.
            r'^\d+\.\d+\.\d+',  # 1.1.1, 2.3.4, etc.
            r'^[A-Z]\.\d+',  # A.1, B.2, etc.
        ]
        
        for pattern in patterns:
            if re.match(pattern, block.text.strip()):
                return True
        
        return False
    
    def _check_duplicate_content(self, block: ContentBlock) -> bool:
        """Check for duplicate content (simplified implementation)"""
        # This would need access to all blocks for proper duplicate detection
        # For now, return False as this is a placeholder
        return False
    
    def _create_validation_summary(self, errors: List[ExtractionError], 
                                 corrections: List[Dict[str, Any]]) -> Dict[str, Any]:
        """Create a summary of validation results"""
        error_counts = {}
        for error in errors:
            error_counts[error.error_type] = error_counts.get(error.error_type, 0) + 1
        
        correction_counts = {}
        for correction in corrections:
            correction_counts[correction["correction_type"]] = correction_counts.get(
                correction["correction_type"], 0) + 1
        
        return {
            "total_errors": len(errors),
            "total_corrections": len(corrections),
            "error_types": error_counts,
            "correction_types": correction_counts,
            "validation_passed": len([e for e in errors if "error" in e.error_type.lower()]) == 0
        }
    
    def save_validation_report(self, results: ValidationResults, output_path: str):
        """Save validation results to a JSON file"""
        report_data = {
            "validation_summary": results.validation_summary,
            "errors": [
                {
                    "line_number": error.line_number,
                    "error_type": error.error_type,
                    "message": error.message,
                    "context": error.context,
                    "expected": error.expected,
                    "found": error.found
                }
                for error in results.errors
            ],
            "corrections": results.corrections
        }
        
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(report_data, f, indent=2, ensure_ascii=False)
        
        self.logger.info(f"Validation report saved to: {output_path}")
    
    def load_validation_rules(self, rules_file: str):
        """Load custom validation rules from a JSON file"""
        try:
            with open(rules_file, 'r', encoding='utf-8') as f:
                rules_data = json.load(f)
            
            custom_rules = []
            for rule_data in rules_data:
                rule = ValidationRule(
                    name=rule_data["name"],
                    description=rule_data["description"],
                    severity=rule_data["severity"],
                    pattern=rule_data.get("pattern"),
                    condition=None,  # Would need to implement condition parsing
                    correction=None   # Would need to implement correction parsing
                )
                custom_rules.append(rule)
            
            self.validation_rules.extend(custom_rules)
            self.logger.info(f"Loaded {len(custom_rules)} custom validation rules")
            
        except Exception as e:
            self.logger.error(f"Failed to load validation rules: {e}")
    
    def _create_test_error(self, error_type: str, line_number: int):
        """Helper method to create test errors for unit tests"""
        return ExtractionError(
            line_number=line_number,
            error_type=error_type,
            message=f"Test {error_type} error",
            context="Test context"
        ) 