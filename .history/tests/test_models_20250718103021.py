"""
Unit tests for core data models

Tests the data structures used throughout the application.
"""

import unittest
from datetime import datetime
from pathlib import Path
import sys

# Add src to path for imports
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from core.models import (
    ContentBlock, 
    HeaderFooterData, 
    TemplateAnalysis, 
    ValidationResults, 
    ExtractionError,
    SpecDocument, 
    BatchJob, 
    BatchResults
)


class TestContentBlock(unittest.TestCase):
    """Test ContentBlock data model"""
    
    def test_content_block_creation(self):
        """Test creating a ContentBlock with basic data"""
        block = ContentBlock(
            text="Test heading",
            level_type="heading",
            level_number=1,
            content="Test content"
        )
        
        self.assertEqual(block.text, "Test heading")
        self.assertEqual(block.level_type, "heading")
        self.assertEqual(block.level_number, 1)
        self.assertEqual(block.content, "Test content")
    
    def test_content_block_with_styling(self):
        """Test ContentBlock with styling information"""
        block = ContentBlock(
            text="Styled text",
            level_type="subheading",
            font_name="Arial",
            font_size=12.0,
            font_bold=True,
            font_italic=False
        )
        
        self.assertEqual(block.font_name, "Arial")
        self.assertEqual(block.font_size, 12.0)
        self.assertTrue(block.font_bold)
        self.assertFalse(block.font_italic)
    
    def test_content_block_defaults(self):
        """Test ContentBlock with minimal data"""
        block = ContentBlock(text="Minimal block", level_type="content")
        
        self.assertEqual(block.text, "Minimal block")
        self.assertEqual(block.level_type, "content")
        self.assertIsNone(block.level_number)
        self.assertEqual(block.content, "")


class TestExtractionError(unittest.TestCase):
    """Test ExtractionError data model"""
    
    def test_extraction_error_creation(self):
        """Test creating an ExtractionError"""
        error = ExtractionError(
            line_number=10,
            error_type="validation_error",
            message="Test error message",
            context="Error context",
            expected="Expected value",
            found="Found value"
        )
        
        self.assertEqual(error.line_number, 10)
        self.assertEqual(error.error_type, "validation_error")
        self.assertEqual(error.message, "Test error message")
        self.assertEqual(error.context, "Error context")
        self.assertEqual(error.expected, "Expected value")
        self.assertEqual(error.found, "Found value")


class TestHeaderFooterData(unittest.TestCase):
    """Test HeaderFooterData data model"""
    
    def test_header_footer_data_creation(self):
        """Test creating HeaderFooterData"""
        data = HeaderFooterData(
            header={"section": ["header content"]},
            footer={"section": ["footer content"]},
            margins={"top": 1.0, "bottom": 1.0},
            document_settings={"setting1": "value1"}
        )
        
        self.assertIn("section", data.header)
        self.assertIn("section", data.footer)
        self.assertEqual(data.margins["top"], 1.0)
        self.assertEqual(data.document_settings["setting1"], "value1")


class TestTemplateAnalysis(unittest.TestCase):
    """Test TemplateAnalysis data model"""
    
    def test_template_analysis_creation(self):
        """Test creating TemplateAnalysis"""
        analysis = TemplateAnalysis(
            template_path="/path/to/template.docx",
            analysis_timestamp="2024-01-01T00:00:00",
            numbering_definitions={"def1": "value1"},
            bwa_list_levels={"level1": "value1"},
            level_mappings={"mapping1": "value1"},
            summary={"total_levels": 5}
        )
        
        self.assertEqual(analysis.template_path, "/path/to/template.docx")
        self.assertEqual(analysis.analysis_timestamp, "2024-01-01T00:00:00")
        self.assertIn("def1", analysis.numbering_definitions)
        self.assertIn("level1", analysis.bwa_list_levels)
        self.assertIn("mapping1", analysis.level_mappings)
        self.assertEqual(analysis.summary["total_levels"], 5)


class TestValidationResults(unittest.TestCase):
    """Test ValidationResults data model"""
    
    def test_validation_results_creation(self):
        """Test creating ValidationResults"""
        error = ExtractionError(
            line_number=1,
            error_type="test_error",
            message="Test error",
            context="Test context"
        )
        
        results = ValidationResults(
            errors=[error],
            corrections=[{"type": "test_correction"}],
            validation_summary={"total_errors": 1}
        )
        
        self.assertEqual(len(results.errors), 1)
        self.assertEqual(len(results.corrections), 1)
        self.assertEqual(results.validation_summary["total_errors"], 1)


class TestSpecDocument(unittest.TestCase):
    """Test SpecDocument data model"""
    
    def test_spec_document_creation(self):
        """Test creating SpecDocument"""
        block = ContentBlock(text="Test block", level_type="content")
        header_footer = HeaderFooterData(
            header={}, footer={}, margins={}, document_settings={}
        )
        validation_results = ValidationResults(
            errors=[], corrections=[], validation_summary={}
        )
        
        document = SpecDocument(
            file_path="/path/to/document.docx",
            content_blocks=[block],
            header_footer=header_footer,
            validation_results=validation_results,
            extraction_timestamp="2024-01-01T00:00:00",
            section_number="1.1",
            section_title="Test Section"
        )
        
        self.assertEqual(document.file_path, "/path/to/document.docx")
        self.assertEqual(len(document.content_blocks), 1)
        self.assertEqual(document.section_number, "1.1")
        self.assertEqual(document.section_title, "Test Section")


class TestBatchJob(unittest.TestCase):
    """Test BatchJob data model"""
    
    def test_batch_job_creation(self):
        """Test creating BatchJob"""
        job = BatchJob(
            name="Test Job",
            description="Test batch job",
            input_paths=["/path/to/file1.docx", "/path/to/file2.docx"],
            template_path="/path/to/template.docx",
            output_dir="/path/to/output",
            options={"extract_only": True}
        )
        
        self.assertEqual(job.name, "Test Job")
        self.assertEqual(job.description, "Test batch job")
        self.assertEqual(len(job.input_paths), 2)
        self.assertEqual(job.template_path, "/path/to/template.docx")
        self.assertEqual(job.output_dir, "/path/to/output")
        self.assertTrue(job.options["extract_only"])


class TestBatchResults(unittest.TestCase):
    """Test BatchResults data model"""
    
    def test_batch_results_creation(self):
        """Test creating BatchResults"""
        job = BatchJob(
            name="Test Job",
            input_paths=[],
            template_path="/path/to/template.docx",
            output_dir="/path/to/output",
            options={}
        )
        
        start_time = datetime.now()
        end_time = datetime.now()
        
        results = BatchResults(
            job=job,
            successful=["file1.docx", "file2.docx"],
            failed=["file3.docx"],
            errors=["Error processing file3.docx"],
            start_time=start_time,
            end_time=end_time,
            total_processed=3,
            total_successful=2,
            total_failed=1
        )
        
        self.assertEqual(results.job.name, "Test Job")
        self.assertEqual(len(results.successful), 2)
        self.assertEqual(len(results.failed), 1)
        self.assertEqual(len(results.errors), 1)
        self.assertEqual(results.total_processed, 3)
        self.assertEqual(results.total_successful, 2)
        self.assertEqual(results.total_failed, 1)


if __name__ == '__main__':
    unittest.main() 