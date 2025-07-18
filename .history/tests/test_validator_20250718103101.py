"""
Unit tests for validator module

Tests the validation and correction logic for specification documents.
"""

import unittest
import tempfile
import json
from pathlib import Path
import sys

# Add src to path for imports
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from core.validator import SpecValidator, ValidationRule
from core.models import ContentBlock, SpecDocument, HeaderFooterData, ValidationResults


class TestValidationRule(unittest.TestCase):
    """Test ValidationRule data model"""
    
    def test_validation_rule_creation(self):
        """Test creating a ValidationRule"""
        rule = ValidationRule(
            name="test_rule",
            description="Test validation rule",
            severity="error",
            pattern=r"^\d+",
            condition=lambda x: True,
            correction=lambda x: x
        )
        
        self.assertEqual(rule.name, "test_rule")
        self.assertEqual(rule.description, "Test validation rule")
        self.assertEqual(rule.severity, "error")
        self.assertEqual(rule.pattern, r"^\d+")
        self.assertTrue(rule.condition(None))
        self.assertEqual(rule.correction("test"), "test")


class TestSpecValidator(unittest.TestCase):
    """Test SpecValidator functionality"""
    
    def setUp(self):
        """Set up test fixtures"""
        self.validator = SpecValidator()
    
    def test_validator_initialization(self):
        """Test validator initialization"""
        self.assertIsNotNone(self.validator)
        self.assertIsInstance(self.validator.validation_rules, list)
        self.assertGreater(len(self.validator.validation_rules), 0)
    
    def test_validate_content_block_empty(self):
        """Test validation of empty content block"""
        block = ContentBlock(text="", level_type="content")
        errors = self.validator._validate_content_block(block, 0)
        
        # Should find empty content error
        self.assertGreater(len(errors), 0)
        self.assertEqual(errors[0].error_type, "empty_content_blocks")
    
    def test_validate_content_block_missing_level_number(self):
        """Test validation of heading without level number"""
        block = ContentBlock(
            text="Test Heading",
            level_type="heading"
        )
        errors = self.validator._validate_content_block(block, 0)
        
        # Should find missing level number warning
        self.assertGreater(len(errors), 0)
        self.assertEqual(errors[0].error_type, "missing_level_numbers")
    
    def test_validate_content_block_valid(self):
        """Test validation of valid content block"""
        block = ContentBlock(
            text="1.1 Test Heading",
            level_type="heading",
            level_number=1
        )
        errors = self.validator._validate_content_block(block, 0)
        
        # Should not find any errors
        self.assertEqual(len(errors), 0)
    
    def test_correct_content_block_empty(self):
        """Test correction of empty content block"""
        block = ContentBlock(
            text="",
            level_type="heading",
            level_number=1
        )
        corrections = self.validator._correct_content_block(block, 0)
        
        # Should apply correction
        self.assertGreater(len(corrections), 0)
        self.assertEqual(corrections[0]["correction_type"], "empty_content")
        self.assertEqual(block.text, "Heading 1")
    
    def test_correct_content_block_missing_level_number(self):
        """Test correction of missing level number"""
        block = ContentBlock(
            text="1.1 Test Heading",
            level_type="heading"
        )
        corrections = self.validator._correct_content_block(block, 0)
        
        # Should extract level number from text
        self.assertGreater(len(corrections), 0)
        self.assertEqual(corrections[0]["correction_type"], "missing_level_number")
        self.assertEqual(block.level_number, 1)
    
    def test_validate_document_structure_insufficient_content(self):
        """Test validation of document with insufficient content"""
        document = SpecDocument(
            file_path="test.docx",
            content_blocks=[ContentBlock(text="Block 1", level_type="content")],
            header_footer=HeaderFooterData(header={}, footer={}, margins={}, document_settings={}),
            template_analysis=None,
            validation_results=ValidationResults(errors=[], corrections=[], validation_summary={})
        )
        
        errors = self.validator._validate_document_structure(document)
        
        # Should find insufficient content error
        self.assertGreater(len(errors), 0)
        self.assertEqual(errors[0].error_type, "insufficient_content")
    
    def test_validate_document_structure_valid(self):
        """Test validation of valid document structure"""
        blocks = [
            ContentBlock(text="Block 1", level_type="heading", level_number=1),
            ContentBlock(text="Block 2", level_type="subheading", level_number=2),
            ContentBlock(text="Block 3", level_type="content"),
            ContentBlock(text="Block 4", level_type="content")
        ]
        
        document = SpecDocument(
            file_path="test.docx",
            content_blocks=blocks,
            header_footer=HeaderFooterData(header={}, footer={}, margins={}, document_settings={}),
            validation_results=ValidationResults(errors=[], corrections=[], validation_summary={})
        )
        
        errors = self.validator._validate_document_structure(document)
        
        # Should not find any errors
        self.assertEqual(len(errors), 0)
    
    def test_check_numbering_consistency(self):
        """Test numbering consistency checking"""
        # Test valid numbering patterns
        valid_patterns = [
            "1.1 Test",
            "2.3.4 Test",
            "A.1 Test"
        ]
        
        for pattern in valid_patterns:
            block = ContentBlock(text=pattern, level_type="content")
            self.assertTrue(self.validator._check_numbering_consistency(block))
        
        # Test invalid numbering patterns
        invalid_patterns = [
            "Test 1.1",
            "No numbering",
            ""
        ]
        
        for pattern in invalid_patterns:
            block = ContentBlock(text=pattern, level_type="content")
            self.assertFalse(self.validator._check_numbering_consistency(block))
    
    def test_create_validation_summary(self):
        """Test validation summary creation"""
        errors = [
            self.validator._create_test_error("error1", 1),
            self.validator._create_test_error("error2", 2),
            self.validator._create_test_error("error1", 3)
        ]
        
        corrections = [
            {"correction_type": "correction1"},
            {"correction_type": "correction2"}
        ]
        
        summary = self.validator._create_validation_summary(errors, corrections)
        
        self.assertEqual(summary["total_errors"], 3)
        self.assertEqual(summary["total_corrections"], 2)
        self.assertEqual(summary["error_types"]["error1"], 2)
        self.assertEqual(summary["error_types"]["error2"], 1)
        self.assertEqual(summary["correction_types"]["correction1"], 1)
        self.assertEqual(summary["correction_types"]["correction2"], 1)
    
    def test_save_validation_report(self):
        """Test saving validation report to JSON"""
        with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False) as f:
            temp_file = f.name
        
        try:
            # Create test validation results
            errors = [self.validator._create_test_error("test_error", 1)]
            corrections = [{"correction_type": "test_correction"}]
            validation_summary = {"total_errors": 1, "total_corrections": 1}
            
            results = ValidationResults(
                errors=errors,
                corrections=corrections,
                validation_summary=validation_summary
            )
            
            # Save report
            self.validator.save_validation_report(results, temp_file)
            
            # Verify file was created and contains expected data
            with open(temp_file, 'r', encoding='utf-8') as f:
                report_data = json.load(f)
            
            self.assertEqual(report_data["validation_summary"]["total_errors"], 1)
            self.assertEqual(len(report_data["errors"]), 1)
            self.assertEqual(len(report_data["corrections"]), 1)
            
        finally:
            # Clean up
            Path(temp_file).unlink(missing_ok=True)
    
    def test_load_validation_rules(self):
        """Test loading custom validation rules"""
        with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False) as f:
            temp_file = f.name
        
        try:
            # Create test rules file
            rules_data = [
                {
                    "name": "custom_rule",
                    "description": "Custom validation rule",
                    "severity": "warning",
                    "pattern": r"custom_pattern"
                }
            ]
            
            with open(temp_file, 'w', encoding='utf-8') as f:
                json.dump(rules_data, f)
            
            # Load rules
            initial_rule_count = len(self.validator.validation_rules)
            self.validator.load_validation_rules(temp_file)
            
            # Verify rule was added
            self.assertEqual(len(self.validator.validation_rules), initial_rule_count + 1)
            
            # Find the custom rule
            custom_rule = None
            for rule in self.validator.validation_rules:
                if rule.name == "custom_rule":
                    custom_rule = rule
                    break
            
            self.assertIsNotNone(custom_rule)
            self.assertEqual(custom_rule.description, "Custom validation rule")
            self.assertEqual(custom_rule.severity, "warning")
            self.assertEqual(custom_rule.pattern, r"custom_pattern")
            
        finally:
            # Clean up
            Path(temp_file).unlink(missing_ok=True)


if __name__ == '__main__':
    unittest.main() 