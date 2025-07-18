"""
Core data models for SpecConverter v1.0

Defines the main data structures used throughout the application.
"""

from dataclasses import dataclass
from typing import Dict, List, Optional, Any, Tuple
from datetime import datetime


@dataclass
class ExtractionError:
    """Represents an error found during content extraction"""
    line_number: int
    error_type: str
    message: str
    context: str
    expected: Optional[str] = None
    found: Optional[str] = None


@dataclass
class ContentBlock:
    """Represents a content block with level information and styling"""
    text: str
    level_type: str
    number: Optional[str] = None
    content: str = ""
    level_number: Optional[int] = None
    bwa_level_name: Optional[str] = None
    numbering_id: Optional[str] = None
    numbering_level: Optional[int] = None
    style_name: Optional[str] = None
    
    # Styling information
    font_name: Optional[str] = None
    font_size: Optional[float] = None
    font_bold: Optional[bool] = None
    font_italic: Optional[bool] = None
    font_underline: Optional[str] = None
    font_color: Optional[str] = None
    font_strike: Optional[bool] = None
    font_small_caps: Optional[bool] = None
    font_all_caps: Optional[bool] = None
    
    # Paragraph formatting
    paragraph_alignment: Optional[str] = None
    paragraph_indent_left: Optional[float] = None
    paragraph_indent_right: Optional[float] = None
    paragraph_indent_first_line: Optional[float] = None
    paragraph_spacing_before: Optional[float] = None
    paragraph_spacing_after: Optional[float] = None
    paragraph_line_spacing: Optional[float] = None
    paragraph_line_spacing_rule: Optional[str] = None
    paragraph_keep_with_next: Optional[bool] = None
    paragraph_keep_lines_together: Optional[bool] = None
    paragraph_page_break_before: Optional[bool] = None
    paragraph_widow_control: Optional[bool] = None
    paragraph_dont_add_space_between_same_style: Optional[bool] = None
    
    # Level list properties
    number_alignment: Optional[str] = None
    aligned_at: Optional[float] = None
    text_indent_at: Optional[float] = None
    follow_number_with: Optional[str] = None
    add_tab_stop_at: Optional[float] = None
    link_level_to_style: Optional[str] = None
    
    # Fallback tracking
    used_fallback_styling: Optional[bool] = None


@dataclass
class HeaderFooterData:
    """Represents header, footer, margin, and document settings"""
    header: Dict[str, List]
    footer: Dict[str, List]
    margins: Dict[str, float]
    document_settings: Dict[str, Any]


@dataclass
class TemplateAnalysis:
    """Represents complete template analysis"""
    template_path: str
    analysis_timestamp: str
    numbering_definitions: Dict[str, Any]
    bwa_list_levels: Dict[str, Any]
    level_mappings: Dict[str, str]
    summary: Dict[str, Any]


@dataclass
class ValidationResults:
    """Represents validation results and corrections"""
    errors: List[ExtractionError]
    corrections: List[Dict[str, Any]]
    validation_summary: Dict[str, Any]


@dataclass
class SpecDocument:
    """Represents a complete specification document"""
    file_path: str
    content_blocks: List[ContentBlock]
    header_footer: HeaderFooterData
    template_analysis: Optional[TemplateAnalysis]
    validation_results: ValidationResults
    extraction_timestamp: str = ""
    section_number: str = ""
    section_title: str = ""


@dataclass
class BatchJob:
    """Represents a batch processing job"""
    name: str
    input_paths: List[str]
    template_path: str
    output_dir: str
    options: Dict[str, Any]
    description: str = ""


@dataclass
class BatchResults:
    """Represents batch processing results"""
    job: BatchJob
    successful: List[str]
    failed: List[str]
    errors: List[str]
    start_time: datetime
    end_time: datetime
    total_processed: int = 0
    total_successful: int = 0
    total_failed: int = 0 