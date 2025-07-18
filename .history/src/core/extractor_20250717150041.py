#!/usr/bin/env python3
"""
Specification Content Extractor - Version 3

This script extracts multi-level list content from Word documents (.docx) and converts it to JSON format.
It combines the best of both approaches:
- JSON output structure (working well)
- Header/footer/margin extraction from rip scripts
- Comments extraction
- BWA list level detection and mapping
- Template-based validation

Features:
- Extracts section headers, titles, parts, subsections, items, and lists
- Handles both numbered and unnumbered structures
- Validates numbering sequences and reports errors
- Extracts header, footer, margin, and comment information
- Maps content to BWA list levels from template
- Outputs comprehensive JSON with level information
- Generates detailed error reports

Usage:
    python extract_spec_content_v3.py <docx_file> [output_dir] [template_file]

Example:
    python extract_spec_content_v3.py "SECTION 26 05 00.docx"
"""

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import json
import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from typing import Dict, List, Optional, Tuple, Any
from dataclasses import dataclass
from datetime import datetime

# Import the header/footer extractor module
from header_footer_extractor import HeaderFooterExtractor

# Import the template list detector module
from template_list_detector import TemplateListDetector

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

class SpecContentExtractorV3:
    """Extracts specification content with comprehensive metadata"""
    
    def __init__(self, template_path: Optional[str] = None):
        self.errors: List[ExtractionError] = []
        self.line_count: int = 0
        self.section_header_found: bool = False
        self.section_title_found: bool = False
        
        # Document structure
        self.section_number: str = ""
        self.section_title: str = ""
        self.end_of_section: str = ""
        
        # Template analysis
        self.template_path: Optional[str] = template_path
        self.bwa_list_levels: Dict[str, Any] = {}
        self.template_numbering: Dict[str, Any] = {}
        self.template_analysis: Optional[Any] = None
        
        # Content blocks
        self.content_blocks: List[ContentBlock] = []
        
        # List numbering tracking
        self.list_counters: Dict[Tuple[str, int], int] = {}  # {(numId, ilvl): current_number}
        self.list_fixes: List[Dict[str, Any]] = []  # Track numbering fixes for reporting
        
        # Regex patterns
        self.section_pattern = re.compile(r'^SECTION\s+(.+)$', re.IGNORECASE)
        self.end_section_pattern = re.compile(r'^END\s+OF\s+SECTION\s*(.+)?$', re.IGNORECASE)
        self.part_pattern = re.compile(r'^(\d+\.0)\s+(.+)$')
        self.subsection_pattern = re.compile(r'^(\d+\.\d{2})\s+(.+)$')
        self.subsection_alt_pattern = re.compile(r'^(\d+\.\d)\s+(.+)$')
        self.item_pattern = re.compile(r'^([A-Z])\.\s+(.+)$')
        self.list_pattern = re.compile(r'^(\d+)\.\s+(.+)$')
        self.sub_list_pattern = re.compile(r'^([a-z])\.\s+(.+)$')
        
        # Load template if provided
        if template_path:
            self.load_template_analysis(template_path)
    
    def load_template_analysis(self, template_path: str):
        """Load and analyze template structure using the template list detector module"""
        try:
            print(f"Loading template analysis from: {template_path}")
            
            # Use the template list detector module
            detector = TemplateListDetector()
            analysis = detector.analyze_template(template_path)
            
            # Store the analysis results
            self.template_numbering = analysis.numbering_definitions
            self.bwa_list_levels = analysis.bwa_list_levels
            self.template_analysis = analysis
            
            print(f"Template loaded: {len(self.bwa_list_levels)} BWA list levels found")
            
        except Exception as e:
            print(f"Warning: Could not load template analysis: {e}")
            self.template_analysis = None
    
    def get_paragraph_level(self, paragraph: Any) -> Optional[int]:
        """Get the list level of a paragraph"""
        try:
            pPr = paragraph._p.pPr
            if pPr is not None and pPr.numPr is not None:
                if pPr.numPr.ilvl is not None:
                    return pPr.numPr.ilvl.val
        except:
            pass
        return None
    
    def get_paragraph_numbering_id(self, paragraph: Any) -> Optional[str]:
        """Get the numbering ID of a paragraph"""
        try:
            pPr = paragraph._p.pPr
            if pPr is not None and pPr.numPr is not None:
                if pPr.numPr.numId is not None:
                    return str(pPr.numPr.numId.val)
        except:
            pass
        return None
    
    def extract_paragraph_styling(self, paragraph: Any) -> Dict[str, Any]:
        """Extract styling information from a paragraph"""
        styling = {}
        
        try:
            # Paragraph properties
            pPr = paragraph._p.pPr
            if pPr is not None:
                # Alignment
                jc_elem = pPr.find(qn('w:jc'))
                if jc_elem is not None:
                    styling['paragraph_alignment'] = jc_elem.get(qn('w:val'))
                
                # Indentation
                indent_elem = pPr.find(qn('w:ind'))
                if indent_elem is not None:
                    left = indent_elem.get(qn('w:left'))
                    if left is not None:
                        styling['paragraph_indent_left'] = float(left) / 20.0  # Convert twips to points
                    
                    right = indent_elem.get(qn('w:right'))
                    if right is not None:
                        styling['paragraph_indent_right'] = float(right) / 20.0
                    
                    first_line = indent_elem.get(qn('w:firstLine'))
                    if first_line is not None:
                        styling['paragraph_indent_first_line'] = float(first_line) / 20.0
                    
                    hanging = indent_elem.get(qn('w:hanging'))
                    if hanging is not None:
                        styling['paragraph_indent_hanging'] = float(hanging) / 20.0
                
                # Spacing
                spacing_elem = pPr.find(qn('w:spacing'))
                if spacing_elem is not None:
                    before = spacing_elem.get(qn('w:before'))
                    if before is not None:
                        styling['paragraph_spacing_before'] = float(before) / 20.0
                    
                    after = spacing_elem.get(qn('w:after'))
                    if after is not None:
                        styling['paragraph_spacing_after'] = float(after) / 20.0
                    
                    line = spacing_elem.get(qn('w:line'))
                    if line is not None:
                        styling['paragraph_line_spacing'] = float(line) / 240.0  # Convert to line spacing ratio
                    
                    line_rule = spacing_elem.get(qn('w:lineRule'))
                    if line_rule is not None:
                        styling['paragraph_line_spacing_rule'] = line_rule
                
                # Other paragraph properties
                keep_with_next = pPr.find(qn('w:keepNext'))
                if keep_with_next is not None:
                    styling['paragraph_keep_with_next'] = True
                
                keep_lines = pPr.find(qn('w:keepLines'))
                if keep_lines is not None:
                    styling['paragraph_keep_lines_together'] = True
                
                page_break = pPr.find(qn('w:pageBreakBefore'))
                if page_break is not None:
                    styling['paragraph_page_break_before'] = True
                
                widow_control = pPr.find(qn('w:widowControl'))
                if widow_control is not None:
                    styling['paragraph_widow_control'] = True
                
                # Dont add space between same style
                dont_add_space = pPr.find(qn('w:dontAddSpaceBetweenSameStyle'))
                if dont_add_space is not None:
                    styling['paragraph_dont_add_space_between_same_style'] = True
            
            # Run properties (font information)
            # Get the most common font properties from all runs
            font_properties = self.extract_run_styling(paragraph.runs)
            styling.update(font_properties)
            
        except Exception as e:
            print(f"Warning: Could not extract paragraph styling: {e}")
        
        return styling
    
    def extract_run_styling(self, runs: List[Any]) -> Dict[str, Any]:
        """Extract styling information from paragraph runs"""
        styling = {}
        
        if not runs:
            return styling
        
        try:
            # Collect font properties from all runs
            font_names = []
            font_sizes = []
            font_bolds = []
            font_italics = []
            font_underlines = []
            font_colors = []
            font_strikes = []
            font_small_caps_list = []
            font_all_caps_list = []
            
            for run in runs:
                rPr = run._r.rPr
                if rPr is not None:
                    # Font family
                    r_fonts = rPr.find(qn('w:rFonts'))
                    if r_fonts is not None:
                        ascii_font = r_fonts.get(qn('w:ascii'))
                        if ascii_font:
                            font_names.append(ascii_font)
                    
                    # Font size
                    sz = rPr.find(qn('w:sz'))
                    if sz is not None:
                        size_val = sz.get(qn('w:val'))
                        if size_val:
                            font_sizes.append(float(size_val) / 2.0)  # Convert half-points to points
                    
                    # Bold
                    b = rPr.find(qn('w:b'))
                    if b is not None:
                        bold_val = b.get(qn('w:val'))
                        if bold_val is not None:
                            font_bolds.append(bold_val == 'true' or bold_val == '1')
                        else:
                            font_bolds.append(True)  # Default to True if present but no value
                    
                    # Italic
                    i = rPr.find(qn('w:i'))
                    if i is not None:
                        italic_val = i.get(qn('w:val'))
                        if italic_val is not None:
                            font_italics.append(italic_val == 'true' or italic_val == '1')
                        else:
                            font_italics.append(True)
                    
                    # Underline
                    u = rPr.find(qn('w:u'))
                    if u is not None:
                        underline_val = u.get(qn('w:val'))
                        if underline_val:
                            font_underlines.append(underline_val)
                    
                    # Color
                    color = rPr.find(qn('w:color'))
                    if color is not None:
                        color_val = color.get(qn('w:val'))
                        if color_val:
                            font_colors.append(color_val)
                    
                    # Strike
                    strike = rPr.find(qn('w:strike'))
                    if strike is not None:
                        strike_val = strike.get(qn('w:val'))
                        if strike_val is not None:
                            font_strikes.append(strike_val == 'true' or strike_val == '1')
                        else:
                            font_strikes.append(True)
                    
                    # Small caps
                    small_caps = rPr.find(qn('w:smallCaps'))
                    if small_caps is not None:
                        small_caps_val = small_caps.get(qn('w:val'))
                        if small_caps_val is not None:
                            font_small_caps_list.append(small_caps_val == 'true' or small_caps_val == '1')
                        else:
                            font_small_caps_list.append(True)
                    
                    # All caps
                    all_caps = rPr.find(qn('w:caps'))
                    if all_caps is not None:
                        all_caps_val = all_caps.get(qn('w:val'))
                        if all_caps_val is not None:
                            font_all_caps_list.append(all_caps_val == 'true' or all_caps_val == '1')
                        else:
                            font_all_caps_list.append(True)
            
            # Use the most common font properties
            if font_names:
                styling['font_name'] = max(set(font_names), key=font_names.count)
            if font_sizes:
                styling['font_size'] = max(set(font_sizes), key=font_sizes.count)
            if font_bolds:
                styling['font_bold'] = max(set(font_bolds), key=font_bolds.count)
            if font_italics:
                styling['font_italic'] = max(set(font_italics), key=font_italics.count)
            if font_underlines:
                styling['font_underline'] = max(set(font_underlines), key=font_underlines.count)
            if font_colors:
                styling['font_color'] = max(set(font_colors), key=font_colors.count)
            if font_strikes:
                styling['font_strike'] = max(set(font_strikes), key=font_strikes.count)
            if font_small_caps_list:
                styling['font_small_caps'] = max(set(font_small_caps_list), key=font_small_caps_list.count)
            if font_all_caps_list:
                styling['font_all_caps'] = max(set(font_all_caps_list), key=font_all_caps_list.count)
            
        except Exception as e:
            print(f"Warning: Could not extract run styling: {e}")
        
        return styling
    
    def extract_header_footer_margins(self, docx_path: str) -> Dict[str, Any]:
        """Extract header, footer, and margin information using the header/footer extractor module"""
        try:
            extractor = HeaderFooterExtractor()
            return extractor.extract_header_footer_margins(docx_path)
        except Exception as e:
            print(f"Error extracting header/footer/margins: {e}")
            return {"header": {}, "footer": {}, "margins": {}}
    
    def extract_comments(self, docx_path: str) -> List[Dict[str, Any]]:
        """Extract comments from document using the header/footer extractor module"""
        try:
            extractor = HeaderFooterExtractor()
            return extractor.extract_comments(docx_path)
        except Exception as e:
            print(f"Error extracting comments: {e}")
            return []
    
    def classify_paragraph_level(self, text: str) -> Tuple[str, Optional[str], str]:
        """
        Classify a paragraph into its hierarchical level
        Returns: (level_type, number, content)
        """
        text = text.strip()
        if not text:
            return "empty", None, ""
        
        # Check for section header (must be the very first line)
        if text.upper().startswith("SECTION") and not self.section_header_found:
            match = self.section_pattern.match(text)
            if match:
                self.section_header_found = True
                return "section", match.group(1), ""
        elif text.upper().startswith("SECTION") and self.section_header_found:
            self.add_error("Structure Error", "Multiple section headers found", text)
            return "content", None, text
        
        # Check for section title (must be the second line after section header)
        if (self.section_header_found and 
            not self.section_title_found and
            len(text.strip()) > 0):
            self.section_title_found = True
            return "title", None, text
        
        # Check for end of section
        if text.upper().startswith("END OF SECTION"):
            match = self.end_section_pattern.match(text)
            if match:
                self.end_of_section = match.group(1).strip() if match.group(1) else ""
                return "end_of_section", None, self.end_of_section
        
        # Check for part level with numbering (1.0, 2.0, etc.)
        match = self.part_pattern.match(text)
        if match:
            return "part", match.group(1), match.group(2)
        
        # Check for part titles with various formats
        part_names = ["DESCRIPTION", "PRODUCTS", "EXECUTION", "GENERAL"]
        for part_name in part_names:
            match = re.match(rf'(?:PART\s*)?(\d+)\.0?\s*[-]?\s*{part_name}$', text.upper())
            if match:
                part_number = f"{match.group(1)}.0"
                return "part_title", part_number, part_name
            elif text.strip().upper() == part_name:
                part_number = f"{len([b for b in self.content_blocks if b.level_type == 'part']) + 1}.0"
                return "part_title", part_number, part_name
        
        # Check for subsection level with numbering (1.01, 1.02, etc.)
        match = self.subsection_pattern.match(text)
        if match:
            return "subsection", match.group(1), match.group(2)
        
        # Check for subsection level with alternative numbering (1.1, 1.2, etc.)
        match = self.subsection_alt_pattern.match(text)
        if match:
            return "subsection", match.group(1), match.group(2)
        
        # Check for item level (A., B., C., etc.) - MOVED BEFORE subsection titles
        match = self.item_pattern.match(text)
        if match:
            return "item", match.group(1), match.group(2)
        
        # Check for list level (1., 2., etc.)
        match = self.list_pattern.match(text)
        if match:
            return "list", match.group(1), match.group(2)
        
        # Check for sub-list level (a., b., etc.)
        match = self.sub_list_pattern.match(text)
        if match:
            return "sub_list", match.group(1), match.group(2)
        
        # Check for subsection titles without numbering - MOVED AFTER item patterns
        # Subsection titles can be any string, so we need to be more intelligent about detection
        # Look for characteristics of subsection titles:
        # 1. Short text (typically subsection titles are concise)
        # 2. All caps or title case (common for subsection titles)
        # 3. No obvious content indicators
        # 4. Not starting with common patterns
        
        is_short_text = len(text.strip()) < 100
        is_all_caps_or_title_case = (text.strip().isupper() or 
                                    text.strip().istitle() or 
                                    text.strip() == text.strip().title())
        doesnt_start_with_patterns = not any(text.upper().startswith(prefix) 
                                            for prefix in ["SECTION", "PART", "GENERAL", "1.", "2.", "A.", "B.", "a.", "b."])
        no_content_indicators = not any(indicator in text.lower() 
                                       for indicator in ["shall", "will", "must", "should", "note:", "example:", ":", ".", ";"])
        
        # If it looks like a subsection title, classify it as such
        if (is_short_text and is_all_caps_or_title_case and 
            doesnt_start_with_patterns and no_content_indicators):
            return "subsection_title", None, text
        
        # If no pattern matches, it's regular content
        return "content", None, text
    
    def correct_level_type_based_on_numbering(self, level_type: str, numbering_id: Optional[str], 
                                            numbering_level: Optional[int], text: str) -> str:
        """Correct level type based on paragraph numbering information"""
        # Simplified logic: If content has numbering, it's a list item
        if level_type == "content" and numbering_id is not None:
            # Any content with numbering should be classified as a list item
            # Use numbering level to determine list vs sub_list
            if numbering_level == 0:
                return "list"  # Top-level list items
            elif numbering_level == 1:
                return "sub_list"  # Sub-list items
            else:
                return "list"  # Default to list for any numbered content
        
        return level_type
    
    def extract_level_list_properties(self, numbering_id: str, numbering_level: Optional[int]) -> Dict[str, Any]:
        """Extract level list position values from numbering definitions"""
        properties = {}
        
        try:
            # Check if we have template analysis with numbering definitions
            if not hasattr(self, 'template_analysis') or self.template_analysis is None:
                return properties
            
            # Find the numbering definition for this numbering_id
            num_key = f"num_{numbering_id}"
            if num_key in self.template_analysis.numbering_definitions:
                abstract_num_id = self.template_analysis.numbering_definitions[num_key].get("abstract_num_id")
                if abstract_num_id in self.template_analysis.numbering_definitions:
                    abstract_info = self.template_analysis.numbering_definitions[abstract_num_id]
                    level_str = str(numbering_level) if numbering_level is not None else "0"
                    
                    if level_str in abstract_info.get("levels", {}):
                        level_info = abstract_info["levels"][level_str]
                        
                        # Extract number alignment (left, center, right)
                        properties["number_alignment"] = level_info.get("lvlJc")
                        
                        # Extract follow number with (tab, space, nothing)
                        properties["follow_number_with"] = level_info.get("suff")
                        
                        # Extract link level to style
                        properties["link_level_to_style"] = level_info.get("pStyle")
                        
                        # Extract position values from paragraph properties
                        p_pr = level_info.get("pPr", {})
                        if "indent" in p_pr:
                            indent = p_pr["indent"]
                            # Convert twips to points (1 point = 20 twips)
                            if indent.get("left"):
                                properties["aligned_at"] = float(indent["left"]) / 20.0
                            if indent.get("firstLine"):
                                properties["text_indent_at"] = float(indent["firstLine"]) / 20.0
                        
                        # Extract tab stop information
                        if "tabs" in p_pr:
                            tabs = p_pr["tabs"]
                            if "tab" in tabs and tabs["tab"]:
                                # Get the first tab position
                                if isinstance(tabs["tab"], list) and len(tabs["tab"]) > 0:
                                    tab_pos = tabs["tab"][0].get("pos")
                                    if tab_pos:
                                        properties["add_tab_stop_at"] = float(tab_pos) / 20.0
                        
        except Exception as e:
            print(f"Error extracting level list properties: {e}")
        
        return properties
    
    def extract_list_number(self, numbering_id: Optional[str], numbering_level: Optional[int], 
                          detected_number: Optional[str], text: str) -> Tuple[Optional[str], bool]:
        """
        Extract the correct list number from Word's numbering system
        Returns: (correct_number, was_fixed)
        """
        if numbering_id is None or numbering_level is None:
            return detected_number, False
        
        # Create key for tracking this specific list
        key = (numbering_id, numbering_level)
        
        # Initialize counter if this is a new list or level
        if key not in self.list_counters:
            self.list_counters[key] = 1
        else:
            self.list_counters[key] += 1
        
        correct_number = str(self.list_counters[key])
        
        # Check if we need to fix the detected number
        was_fixed = False
        if detected_number is not None and detected_number != correct_number:
            was_fixed = True
            self.list_fixes.append({
                "line_number": self.line_count,
                "text": text[:50] + "..." if len(text) > 50 else text,
                "detected_number": detected_number,
                "correct_number": correct_number,
                "numbering_id": numbering_id,
                "numbering_level": numbering_level
            })
        
        return correct_number, was_fixed

    def map_to_bwa_level(self, paragraph: Any, level_type: str) -> Tuple[Optional[int], Optional[str]]:
        """Map paragraph to BWA list level based on template analysis"""
        try:
            # Standard mapping for level_number
            level_mapping = {
                "part": 0,
                "part_title": 0,
                "subsection": 1,
                "subsection_title": 1,
                "item": 2,
                "list": 3,
                "sub_list": 4
            }
            level_number = level_mapping.get(level_type)

            # Map to BWA style name for label
            level_type_to_bwa_mapping = {
                "section": "BWA-SectionNumber",
                "title": "BWA-SectionTitle",
                "part": "BWA-PART",
                "part_title": "BWA-PART",
                "subsection": "BWA-SUBSECTION",
                "subsection_title": "BWA-SUBSECTION",
                "item": "BWA-Item",
                "list": "BWA-List",
                "sub_list": "BWA-SubList"
            }
            bwa_level_name = None
            if level_type in level_type_to_bwa_mapping:
                bwa_style_name = level_type_to_bwa_mapping[level_type]
                if bwa_style_name in self.bwa_list_levels:
                    bwa_level_name = bwa_style_name
                else:
                    bwa_level_name = bwa_style_name  # Even if not in template, use the label

            return level_number, bwa_level_name
        except Exception as e:
            print(f"Error mapping to BWA level: {e}")
            return None, None
    
    def add_error(self, error_type: str, message: str, context: str = "", 
                  expected: Optional[str] = None, found: Optional[str] = None):
        """Add an error to the error list"""
        error = ExtractionError(
            line_number=self.line_count,
            error_type=error_type,
            message=message,
            context=context,
            expected=expected,
            found=found
        )
        self.errors.append(error)
    
    def validate_and_correct_level_consistency(self) -> Dict[str, Any]:
        """
        Second pass: Validate and correct logical list level patterns
        Returns: Dictionary with validation results and corrections made
        """
        validation_results = {
            "inconsistencies_found": [],
            "corrections_made": [],
            "validation_summary": {}
        }
        
        if not self.content_blocks:
            return validation_results
        
        # ENHANCED: Loop until no more corrections are made
        iteration = 0
        max_iterations = 10  # Prevent infinite loops
        corrections_made_this_iteration = True
        
        while corrections_made_this_iteration and iteration < max_iterations:
            iteration += 1
            corrections_made_this_iteration = False
            
            # Track expected level progression
            expected_level = None
            level_transitions = []
            
            for i, block in enumerate(self.content_blocks):
                current_level = block.level_number
                level_type = block.level_type
                
                # Skip non-hierarchical content
                if level_type in ["section", "title", "end_of_section"]:
                    continue
                
                # Define expected level based on level_type
                expected_level_for_type = {
                    "part": 0,
                    "part_title": 0,
                    "subsection": 1,
                    "subsection_title": 1,
                    "item": 2,
                    "list": 3,
                    "sub_list": 4
                }.get(level_type)
                
                # Check if current level matches expected level for type
                if expected_level_for_type is not None and current_level != expected_level_for_type:
                    inconsistency = {
                        "block_index": i,
                        "text": block.text[:100] + "..." if len(block.text) > 100 else block.text,
                        "level_type": level_type,
                        "current_level": current_level,
                        "expected_level": expected_level_for_type,
                        "correction_applied": False
                    }
                    validation_results["inconsistencies_found"].append(inconsistency)
                    
                    # Apply correction
                    old_level = block.level_number
                    block.level_number = expected_level_for_type
                    block.used_fallback_styling = True  # Mark as corrected
                    
                    validation_results["corrections_made"].append({
                        "block_index": i,
                        "text": block.text[:100] + "..." if len(block.text) > 100 else block.text,
                        "level_type": level_type,
                        "old_level": old_level,
                        "new_level": expected_level_for_type
                    })
                    inconsistency["correction_applied"] = True
                    corrections_made_this_iteration = True
                
                # ENHANCED: Check for logical hierarchy inconsistencies
                # This detects when a block should be at a different level based on context
                skip_remaining_checks = False
                if expected_level is not None and current_level is not None:
                    # Check for jumps that don't make logical sense
                    level_diff = current_level - expected_level
                    
                    # CRITICAL RULE: Part level 0's should never be empty of children and should never appear in sequence
                    # If we have two level 0 blocks in a row, the second one should be level 1
                    if current_level == 0 and i > 0:
                        prev_block = self.content_blocks[i - 1]
                        prev_level = prev_block.level_number
                        
                        if prev_level == 0:
                            # Two level 0 blocks in a row - this violates the rule
                            # The second one should be a subsection title (level 1)
                            suggested_level_type = "subsection_title"
                            suggested_level = 1
                            
                            inconsistency = {
                                "block_index": i,
                                "text": block.text[:100] + "..." if len(block.text) > 100 else block.text,
                                "level_type": level_type,
                                "current_level": current_level,
                                "suggested_level_type": suggested_level_type,
                                "suggested_level": suggested_level,
                                "reason": f"Two level 0 blocks in sequence (prev: {prev_level}, current: {current_level}) - second should be level 1",
                                "correction_applied": False
                            }
                            validation_results["inconsistencies_found"].append(inconsistency)
                            
                            # Apply correction
                            old_level_type = block.level_type
                            old_level = block.level_number
                            block.level_type = suggested_level_type
                            block.level_number = suggested_level
                            block.bwa_level_name = "BWA-SUBSECTION"  # Update BWA level name
                            block.used_fallback_styling = True
                            
                            validation_results["corrections_made"].append({
                                "block_index": i,
                                "text": block.text[:100] + "..." if len(block.text) > 100 else block.text,
                                "old_level_type": old_level_type,
                                "new_level_type": suggested_level_type,
                                "old_level": old_level,
                                "new_level": suggested_level,
                                "reason": "Corrected consecutive level 0 blocks - second should be level 1"
                            })
                            inconsistency["correction_applied"] = True
                            corrections_made_this_iteration = True
                            # Skip other checks for this block since we've already corrected it
                            expected_level = suggested_level
                            skip_remaining_checks = True
                    
                    # ENHANCED: Check for jumps from subsection (level 1) to list (level 3+)
                    # If we jump from level 1 to level 3 or higher, the content should probably be level 2 (item)
                    if not skip_remaining_checks and i > 0:
                        # Look at the previous block to understand context
                        prev_block = self.content_blocks[i - 1]
                        prev_level = prev_block.level_number
                        
                        # If we're jumping from level 1 (subsection) to level 3+ (list/sub_list),
                        # the content should probably be level 2 (item)
                        if prev_level == 1 and current_level >= 3:
                            print(f"DEBUG: Found jump from level {prev_level} to {current_level} at block {i}: {block.text[:50]}...")
                            
                            # This should probably be an item (level 2), not a list (level 3+)
                            suggested_level_type = "item"
                            suggested_level = 2
                            
                            # Adjust the Word numbering level to match the corrected level
                            # If we're changing from sub_list (numbering_level: 1) to item (numbering_level: 0)
                            # or from list (numbering_level: 2) to item (numbering_level: 0)
                            suggested_numbering_level = 0  # Items typically use numbering_level 0
                            
                            inconsistency = {
                                "block_index": i,
                                "text": block.text[:100] + "..." if len(block.text) > 100 else block.text,
                                "level_type": level_type,
                                "current_level": current_level,
                                "suggested_level_type": suggested_level_type,
                                "suggested_level": suggested_level,
                                "reason": f"Jump from level {prev_level} to {current_level} suggests misclassification - should be level 2",
                                "correction_applied": False
                            }
                            validation_results["inconsistencies_found"].append(inconsistency)
                            
                            # Apply correction
                            old_level_type = block.level_type
                            old_level = block.level_number
                            old_numbering_level = block.numbering_level
                            block.level_type = suggested_level_type
                            block.level_number = suggested_level
                            block.numbering_level = suggested_numbering_level
                            block.bwa_level_name = "BWA-Item"  # Update BWA level name
                            block.used_fallback_styling = True
                            
                            validation_results["corrections_made"].append({
                                "block_index": i,
                                "text": block.text[:100] + "..." if len(block.text) > 100 else block.text,
                                "old_level_type": old_level_type,
                                "new_level_type": suggested_level_type,
                                "old_level": old_level,
                                "new_level": suggested_level,
                                "old_numbering_level": old_numbering_level,
                                "new_numbering_level": suggested_numbering_level,
                                "reason": "Corrected jump from subsection to list - should be item level"
                            })
                            inconsistency["correction_applied"] = True
                            corrections_made_this_iteration = True
                
                # Track level transitions for sequence validation
                if current_level is not None:
                    if expected_level is not None:
                        level_transitions.append({
                            "from_block": i - 1,
                            "to_block": i,
                            "from_level": expected_level,
                            "to_level": current_level,
                            "transition_type": self._classify_transition(expected_level, current_level)
                        })
                    expected_level = current_level
                    
                    # Debug: Print level transitions for troubleshooting
                    if i < 10:  # Only print first 10 blocks for debugging
                        print(f"Block {i}: {block.text[:50]}... -> Level {current_level} (expected: {expected_level})")
            
            print(f"Validation iteration {iteration}: {'Corrections made' if corrections_made_this_iteration else 'No corrections needed'}")
        
        # Analyze level transitions for logical consistency
        transition_analysis = self._analyze_level_transitions(level_transitions)
        validation_results["validation_summary"] = {
            "total_blocks": len(self.content_blocks),
            "hierarchical_blocks": len([b for b in self.content_blocks if b.level_type not in ["section", "title", "end_of_section"]]),
            "inconsistencies_found": len(validation_results["inconsistencies_found"]),
            "corrections_applied": len(validation_results["corrections_made"]),
            "level_transitions": len(level_transitions),
            "transition_analysis": transition_analysis,
            "validation_iterations": iteration
        }
        
        return validation_results
    
    def _classify_transition(self, from_level: int, to_level: int) -> str:
        """Classify the type of level transition"""
        if to_level == from_level:
            return "same_level"
        elif to_level == from_level + 1:
            return "increase_by_one"
        elif to_level == from_level - 1:
            return "decrease_by_one"
        elif to_level > from_level + 1:
            return "jump_up"
        elif to_level < from_level - 1:
            return "jump_down"
        else:
            return "irregular"
    
    def _is_numbering_logically_correct(self, detected_number: str, level_type: str, block_index: int) -> bool:
        """
        Check if the detected number appears to be logically correct based on context
        
        Args:
            detected_number: The number detected in the source document
            level_type: The current level type
            block_index: Index of the current block
            
        Returns:
            True if the numbering appears to be logically correct, False otherwise
        """
        try:
            # For item level (A, B, C, etc.), check if it's a reasonable letter
            if level_type == "item" and detected_number.isalpha():
                # Check if it's a reasonable letter (A-Z)
                if detected_number.upper() in "ABCDEFGHIJKLMNOPQRSTUVWXYZ":
                    # Look at previous items to see if this makes sense
                    if block_index > 0:
                        prev_block = self.content_blocks[block_index - 1]
                        if prev_block.level_type == "item" and prev_block.number:
                            # If previous was A, this could be B, C, D, etc.
                            # If previous was J, this could be K, L, M, etc.
                            # Allow reasonable progression
                            return True
                    # First item in a section, A is always reasonable
                    return True
            
            # For list level (1, 2, 3, etc.), check if it's a reasonable number
            elif level_type == "list" and detected_number.isdigit():
                num = int(detected_number)
                # Check if it's a reasonable number (1-99)
                if 1 <= num <= 99:
                    # Look at previous lists to see if this makes sense
                    if block_index > 0:
                        prev_block = self.content_blocks[block_index - 1]
                        if prev_block.level_type == "list" and prev_block.number:
                            try:
                                prev_num = int(prev_block.number)
                                # Allow reasonable progression (prev + 1, or reasonable jump)
                                if num == prev_num + 1 or (num > prev_num and num <= prev_num + 10):
                                    return True
                            except ValueError:
                                pass
                    # First list item, 1 is always reasonable
                    return True
            
            # For other level types, be more permissive
            else:
                return True
                
        except Exception as e:
            print(f"Warning: Error checking numbering logic: {e}")
            return True  # Default to trusting the source if we can't determine
        
        return False
    
    def _analyze_level_transitions(self, transitions: List[Dict[str, Any]]) -> Dict[str, Any]:
        """Analyze level transitions for patterns and issues"""
        if not transitions:
            return {"total_transitions": 0, "patterns": {}}
        
        transition_counts = {}
        irregular_transitions = []
        
        for transition in transitions:
            transition_type = transition["transition_type"]
            transition_counts[transition_type] = transition_counts.get(transition_type, 0) + 1
            
            if transition_type in ["jump_up", "jump_down", "irregular"]:
                irregular_transitions.append(transition)
        
        return {
            "total_transitions": len(transitions),
            "patterns": transition_counts,
            "irregular_transitions": len(irregular_transitions),
            "irregular_details": irregular_transitions[:10]  # Limit to first 10 for reporting
        }
    
    def extract_content(self, docx_path: str) -> Dict[str, Any]:
        """Extract content from a Word document"""
        try:
            doc = Document(docx_path)
            
            # Extract header, footer, margin, and comment information
            header_footer_data = self.extract_header_footer_margins(docx_path)
            comments = self.extract_comments(docx_path)
            
            # Extract all paragraphs
            paragraphs = []
            for paragraph in doc.paragraphs:
                text = paragraph.text.strip()
                if text:  # Only include non-empty paragraphs
                    paragraphs.append((text, paragraph))
                    self.line_count += 1
            
            # Extract section header and title
            section_number, section_title = self.extract_section_header_and_title([p[0] for p in paragraphs])
            
            # Process each paragraph
            for text, paragraph in paragraphs:
                level_type, number, content = self.classify_paragraph_level(text)
                
                # Skip empty content
                if level_type == "empty":
                    continue
                
                # Get numbering information
                numbering_id = self.get_paragraph_numbering_id(paragraph)
                numbering_level = self.get_paragraph_level(paragraph)
                style_name = paragraph.style.name if paragraph.style else None
                
                # Extract styling information
                styling = self.extract_paragraph_styling(paragraph)
                
                # Extract level list properties if this is a numbered paragraph
                level_list_properties = {}
                if numbering_id is not None:
                    level_list_properties = self.extract_level_list_properties(numbering_id, numbering_level)
                
                # Post-process classification based on numbering information
                corrected_level_type = self.correct_level_type_based_on_numbering(
                    level_type, numbering_id, numbering_level, text
                )
                
                # Extract correct list number and check for fixes
                correct_number, was_fixed = self.extract_list_number(
                    numbering_id, numbering_level, number, text
                )
                
                # Map to BWA level using corrected level type
                level_number, bwa_level_name = self.map_to_bwa_level(paragraph, corrected_level_type)
                
                # Normalize level_number: Use logical mapping for known level types, otherwise use Word's numbering_level
                if corrected_level_type in ["part", "part_title", "subsection", "subsection_title", "item", "list", "sub_list"]:
                    # Use our logical mapping for known level types
                    # This ensures consistent level numbering regardless of Word's internal numbering
                    pass  # level_number is already set by map_to_bwa_level
                elif numbering_level is not None:
                    # For unknown level types, use Word's numbering level as fallback
                    level_number = numbering_level
                # If neither, level_number remains None
                
                # Create content block with styling information
                block = ContentBlock(
                    text=text,
                    level_type=corrected_level_type,
                    number=correct_number,  # Use the correct number from Word's numbering system
                    content=content,
                    level_number=level_number,
                    bwa_level_name=bwa_level_name,
                    numbering_id=numbering_id,
                    numbering_level=numbering_level,
                    style_name=style_name,
                    # Styling information
                    font_name=styling.get('font_name'),
                    font_size=styling.get('font_size'),
                    font_bold=styling.get('font_bold'),
                    font_italic=styling.get('font_italic'),
                    font_underline=styling.get('font_underline'),
                    font_color=styling.get('font_color'),
                    font_strike=styling.get('font_strike'),
                    font_small_caps=styling.get('font_small_caps'),
                    font_all_caps=styling.get('font_all_caps'),
                    # Paragraph formatting
                    paragraph_alignment=styling.get('paragraph_alignment'),
                    paragraph_indent_left=styling.get('paragraph_indent_left'),
                    paragraph_indent_right=styling.get('paragraph_indent_right'),
                    paragraph_indent_first_line=styling.get('paragraph_indent_first_line'),
                    paragraph_spacing_before=styling.get('paragraph_spacing_before'),
                    paragraph_spacing_after=styling.get('paragraph_spacing_after'),
                    paragraph_line_spacing=styling.get('paragraph_line_spacing'),
                    paragraph_line_spacing_rule=styling.get('paragraph_line_spacing_rule'),
                    paragraph_keep_with_next=styling.get('paragraph_keep_with_next'),
                    paragraph_keep_lines_together=styling.get('paragraph_keep_lines_together'),
                    paragraph_page_break_before=styling.get('paragraph_page_break_before'),
                    paragraph_widow_control=styling.get('paragraph_widow_control'),
                    paragraph_dont_add_space_between_same_style=styling.get('paragraph_dont_add_space_between_same_style'),
                    # Level list properties
                    number_alignment=level_list_properties.get('number_alignment'),
                    aligned_at=level_list_properties.get('aligned_at'),
                    text_indent_at=level_list_properties.get('text_indent_at'),
                    follow_number_with=level_list_properties.get('follow_number_with'),
                    add_tab_stop_at=level_list_properties.get('add_tab_stop_at'),
                    link_level_to_style=level_list_properties.get('link_level_to_style'),
                    # Fallback tracking
                    used_fallback_styling=False  # Will be set to True if fallback is used
                )
                
                self.content_blocks.append(block)
            
            # Second pass: Validate and correct level consistency
            print("Performing second pass validation and correction...")
            self.validation_results = self.validate_and_correct_level_consistency()
            
            # Build the final data structure
            extracted_data = {
                "header": header_footer_data["header"],
                "footer": header_footer_data["footer"],
                "margins": header_footer_data["margins"],
                "document_settings": header_footer_data.get("document_settings", {}),
                "comments": comments,
                "section_number": section_number,
                "section_title": section_title,
                "end_of_section": self.end_of_section,
                "validation_results": self.validation_results,
                "content_blocks": [
                    {
                        "text": block.text,
                        "level_type": block.level_type,
                        "number": block.number,
                        "content": block.content,
                        "level_number": block.level_number,
                        "bwa_level_name": block.bwa_level_name,
                        "numbering_id": block.numbering_id,
                        "numbering_level": block.numbering_level,
                        "style_name": block.style_name,
                        # Styling information
                        "font_name": block.font_name,
                        "font_size": block.font_size,
                        "font_bold": block.font_bold,
                        "font_italic": block.font_italic,
                        "font_underline": block.font_underline,
                        "font_color": block.font_color,
                        "font_strike": block.font_strike,
                        "font_small_caps": block.font_small_caps,
                        "font_all_caps": block.font_all_caps,
                        "paragraph_alignment": block.paragraph_alignment,
                        "paragraph_indent_left": block.paragraph_indent_left,
                        "paragraph_indent_right": block.paragraph_indent_right,
                        "paragraph_indent_first_line": block.paragraph_indent_first_line,
                        "paragraph_spacing_before": block.paragraph_spacing_before,
                        "paragraph_spacing_after": block.paragraph_spacing_after,
                        "paragraph_line_spacing": block.paragraph_line_spacing,
                        "paragraph_line_spacing_rule": block.paragraph_line_spacing_rule,
                        "paragraph_keep_with_next": block.paragraph_keep_with_next,
                        "paragraph_keep_lines_together": block.paragraph_keep_lines_together,
                        "paragraph_page_break_before": block.paragraph_page_break_before,
                        "paragraph_widow_control": block.paragraph_widow_control,
                        "paragraph_dont_add_space_between_same_style": block.paragraph_dont_add_space_between_same_style,
                        "number_alignment": block.number_alignment,
                        "aligned_at": block.aligned_at,
                        "text_indent_at": block.text_indent_at,
                        "follow_number_with": block.follow_number_with,
                        "add_tab_stop_at": block.add_tab_stop_at,
                        "link_level_to_style": block.link_level_to_style,
                        "used_fallback_styling": block.used_fallback_styling
                    }
                    for block in self.content_blocks
                ],
                "template_analysis": self.get_template_analysis_section()
            }
            
            return extracted_data
            
        except Exception as e:
            self.add_error("Extraction Error", f"Failed to extract content: {str(e)}", "")
            return {}
    
    def get_template_analysis_section(self) -> Dict[str, Any]:
        """Get the template analysis section for JSON output"""
        if not hasattr(self, 'template_analysis') or self.template_analysis is None:
            return {
                "template_path": self.template_path,
                "bwa_list_levels": {},
                "template_numbering": {},
                "level_mappings": {},
                "summary": {
                    "total_abstract_numbering": 0,
                    "total_num_mappings": 0,
                    "total_bwa_levels": 0,
                    "level_mappings_count": 0,
                    "level_types": {},
                    "analysis_timestamp": datetime.now().isoformat(),
                    "error": "No template analysis available"
                }
            }
        
        # Convert ListLevelInfo objects to dictionaries for JSON serialization
        bwa_list_levels_dict = {}
        for key, level_info in self.template_analysis.bwa_list_levels.items():
            bwa_list_levels_dict[key] = {
                "level_number": level_info.level_number,
                "numbering_id": level_info.numbering_id,
                "abstract_num_id": level_info.abstract_num_id,
                "level_text": level_info.level_text,
                "number_format": level_info.number_format,
                "start_value": level_info.start_value,
                "suffix": level_info.suffix,
                "justification": level_info.justification,
                "style_name": level_info.style_name,
                "bwa_label": level_info.bwa_label,
                "is_bwa_level": level_info.is_bwa_level
            }
        
        return {
            "template_path": self.template_analysis.template_path,
            "bwa_list_levels": bwa_list_levels_dict,
            "template_numbering": self.template_analysis.numbering_definitions,
            "level_mappings": self.template_analysis.level_mappings,
            "summary": self.template_analysis.summary
        }
    
    def extract_section_header_and_title(self, paragraphs: List[str]) -> Tuple[str, str]:
        """Extract section header and title from the first few paragraphs"""
        section_number = ""
        section_title = ""
        
        if len(paragraphs) >= 2:
            # First paragraph should be section header
            section_text = paragraphs[0].strip()
            if section_text.upper().startswith("SECTION"):
                section_match = re.search(r'^SECTION\s+(.+)$', section_text, re.IGNORECASE)
                if section_match:
                    section_content = section_match.group(1).strip()
                    
                    # Try to extract section number from various formats
                    number_match = re.search(r'(\d+)\s+(\d+)\s+(\d+)', section_content)
                    if number_match:
                        section_number = f"{number_match.group(1)}{number_match.group(2)}{number_match.group(3)}"
                    else:
                        number_match = re.search(r'(\d+)-(\d+)-(\d+)', section_content)
                        if number_match:
                            section_number = f"{number_match.group(1)}{number_match.group(2)}{number_match.group(3)}"
                        else:
                            number_match = re.search(r'(\d{6})', section_content)
                            if number_match:
                                section_number = number_match.group(1)
                            else:
                                section_number = section_content.replace(" ", "").replace("-", "")
            
            # Second paragraph should be section title
            title_text = paragraphs[1].strip()
            if title_text and not title_text.upper().startswith("SECTION"):
                section_title = title_text
        
        return section_number, section_title
    
    def generate_error_report(self) -> str:
        """Generate a comprehensive error report"""
        report = f"ERROR REPORT - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
        report += "=" * 60 + "\n\n"
        
        # Report list numbering fixes
        if self.list_fixes:
            report += f"LIST NUMBERING FIXES ({len(self.list_fixes)} found):\n"
            report += "-" * 40 + "\n"
            
            for i, fix in enumerate(self.list_fixes, 1):
                report += f"{i}. Line {fix['line_number']}: {fix['text']}\n"
                report += f"   Detected: {fix['detected_number']}, Corrected: {fix['correct_number']}\n"
                report += f"   Numbering ID: {fix['numbering_id']}, Level: {fix['numbering_level']}\n"
                report += "\n"
        
        # Report level consistency validation results
        if hasattr(self, 'validation_results') and self.validation_results:
            validation = self.validation_results
            report += f"LEVEL CONSISTENCY VALIDATION:\n"
            report += "-" * 40 + "\n"
            
            summary = validation.get("validation_summary", {})
            report += f"Total blocks: {summary.get('total_blocks', 0)}\n"
            report += f"Hierarchical blocks: {summary.get('hierarchical_blocks', 0)}\n"
            report += f"Inconsistencies found: {summary.get('inconsistencies_found', 0)}\n"
            report += f"Corrections applied: {summary.get('corrections_applied', 0)}\n"
            report += f"Level transitions: {summary.get('level_transitions', 0)}\n"
            
            # Report transition analysis
            transition_analysis = summary.get("transition_analysis", {})
            if transition_analysis:
                report += f"\nTransition Analysis:\n"
                patterns = transition_analysis.get("patterns", {})
                for pattern, count in patterns.items():
                    report += f"  {pattern}: {count}\n"
                
                irregular_count = transition_analysis.get("irregular_transitions", 0)
                if irregular_count > 0:
                    report += f"  Irregular transitions: {irregular_count}\n"
            
            # Report specific inconsistencies
            inconsistencies = validation.get("inconsistencies_found", [])
            if inconsistencies:
                report += f"\nLEVEL INCONSISTENCIES FOUND:\n"
                report += "-" * 40 + "\n"
                
                for i, inconsistency in enumerate(inconsistencies, 1):
                    report += f"{i}. Block {inconsistency['block_index']}: {inconsistency['text']}\n"
                    report += f"   Type: {inconsistency['level_type']}\n"
                    
                    # Handle both old and new inconsistency formats
                    if 'expected_level' in inconsistency:
                        # Old format: level number mismatch
                        report += f"   Current level: {inconsistency['current_level']}, Expected: {inconsistency['expected_level']}\n"
                    elif 'suggested_level_type' in inconsistency:
                        # New format: hierarchy inconsistency
                        report += f"   Current: {inconsistency['level_type']} (level {inconsistency['current_level']})\n"
                        report += f"   Suggested: {inconsistency['suggested_level_type']} (level {inconsistency['suggested_level']})\n"
                        report += f"   Reason: {inconsistency['reason']}\n"
                    
                    report += f"   Correction applied: {'Yes' if inconsistency['correction_applied'] else 'No'}\n"
                    report += "\n"
            
            # Report corrections made
            corrections = validation.get("corrections_made", [])
            if corrections:
                report += f"LEVEL CORRECTIONS APPLIED:\n"
                report += "-" * 40 + "\n"
                
                for i, correction in enumerate(corrections, 1):
                    report += f"{i}. Block {correction['block_index']}: {correction['text']}\n"
                    
                    # Handle both old and new correction formats
                    if 'old_level_type' in correction:
                        # New format: hierarchy correction
                        report += f"   Type: {correction['old_level_type']} → {correction['new_level_type']}\n"
                        report += f"   Level: {correction['old_level']} → {correction['new_level']}\n"
                        report += f"   Reason: {correction['reason']}\n"
                    else:
                        # Old format: level number correction
                        report += f"   Type: {correction['level_type']}\n"
                        report += f"   Level: {correction['old_level']} → {correction['new_level']}\n"
                    
                    report += "\n"
        
        # Report other errors
        if self.errors:
            # Group errors by type
            error_types = {}
            for error in self.errors:
                if error.error_type not in error_types:
                    error_types[error.error_type] = []
                error_types[error.error_type].append(error)
            
            for error_type, errors in error_types.items():
                report += f"{error_type} ERRORS ({len(errors)} found):\n"
                report += "-" * 40 + "\n"
                
                for i, error in enumerate(errors, 1):
                    report += f"{i}. Line {error.line_number}: {error.message}\n"
                    if error.context:
                        report += f"   Context: {error.context}\n"
                    if error.expected and error.found:
                        report += f"   Expected: {error.expected}, Found: {error.found}\n"
                    report += "\n"
        else:
            report += "No errors found during extraction.\n\n"
        
        # Add summary statistics
        report += "SUMMARY:\n"
        report += "-" * 20 + "\n"
        total_errors = len(self.errors)
        total_fixes = len(self.list_fixes)
        report += f"Total errors: {total_errors}\n"
        report += f"List numbering fixes: {total_fixes}\n"
        
        if hasattr(self, 'validation_results') and self.validation_results:
            validation = self.validation_results
            summary = validation.get("validation_summary", {})
            report += f"Level inconsistencies: {summary.get('inconsistencies_found', 0)}\n"
            report += f"Level corrections: {summary.get('corrections_applied', 0)}\n"
        
        if self.errors:
            error_types = {}
            for error in self.errors:
                if error.error_type not in error_types:
                    error_types[error.error_type] = []
                error_types[error.error_type].append(error)
            
            for error_type, errors in error_types.items():
                report += f"{error_type}: {len(errors)} errors\n"
        
        return report
    
    def save_to_json(self, data: Dict[str, Any], output_path: str):
        """Save extracted data to JSON file"""
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
    
    def save_error_report(self, report: str, output_path: str):
        """Save error report to text file"""
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(report)
    
    def save_modular_json_files(self, data: Dict[str, Any], base_name: str, output_dir: str):
        """Save separate JSON files for each modular component"""
        try:
            # 1. Header/Footer JSON
            header_footer_data = {
                "header": data.get("header", {}),
                "footer": data.get("footer", {}),
                "margins": data.get("margins", {}),
                "extraction_timestamp": datetime.now().isoformat(),
                "source_file": base_name
            }
            header_footer_path = os.path.join(output_dir, f"{base_name}_header_footer.json")
            with open(header_footer_path, 'w', encoding='utf-8') as f:
                json.dump(header_footer_data, f, indent=2, ensure_ascii=False)
            print(f"Header/footer data saved to: {header_footer_path}")
            
            # 2. Comments JSON
            comments_data = {
                "comments": data.get("comments", []),
                "extraction_timestamp": datetime.now().isoformat(),
                "source_file": base_name
            }
            comments_path = os.path.join(output_dir, f"{base_name}_comments.json")
            with open(comments_path, 'w', encoding='utf-8') as f:
                json.dump(comments_data, f, indent=2, ensure_ascii=False)
            print(f"Comments data saved to: {comments_path}")
            
            # 3. Template Analysis JSON
            template_data = data.get("template_analysis", {})
            if template_data:
                template_path = os.path.join(output_dir, f"{base_name}_template_analysis.json")
                with open(template_path, 'w', encoding='utf-8') as f:
                    json.dump(template_data, f, indent=2, ensure_ascii=False)
                print(f"Template analysis saved to: {template_path}")
            
            # 4. Content Blocks JSON (with list levels and numbering)
            content_data = {
                "section_number": data.get("section_number", ""),
                "section_title": data.get("section_title", ""),
                "end_of_section": data.get("end_of_section", ""),
                "content_blocks": data.get("content_blocks", []),
                "extraction_timestamp": datetime.now().isoformat(),
                "source_file": base_name
            }
            content_path = os.path.join(output_dir, f"{base_name}_content_blocks.json")
            with open(content_path, 'w', encoding='utf-8') as f:
                json.dump(content_data, f, indent=2, ensure_ascii=False)
            print(f"Content blocks saved to: {content_path}")
            
        except Exception as e:
            print(f"Warning: Could not save modular JSON files: {e}")

def main():
    """Main function to run the extraction"""
    if len(sys.argv) < 2:
        print("Usage: python extract_spec_content_v3.py <docx_file> [output_dir] [template_file]")
        print("Example: python extract_spec_content_v3.py 'SECTION 26 05 00.docx'")
        print("Example: python extract_spec_content_v3.py 'SECTION 26 05 00.docx' . 'templates/test_template_cleaned.docx'")
        print("Note: All output files will be saved to <output_dir>/output/")
        print("Note: Template file must be explicitly specified - no auto-detection")
        sys.exit(1)
    
    docx_path = sys.argv[1]
    output_dir = sys.argv[2] if len(sys.argv) > 2 else "."
    template_path = sys.argv[3] if len(sys.argv) > 3 else None
    
    if not os.path.exists(docx_path):
        print(f"Error: File '{docx_path}' not found.")
        sys.exit(1)
    
    # Require explicit template specification
    if not template_path:
        print("Error: Template file must be specified as the third argument.")
        print("Example: python extract_spec_content_v3.py 'document.docx' . 'templates/test_template_cleaned.docx'")
        print("Available templates:")
        print("  - templates/test_template_cleaned.docx (recommended)")
        print("  - templates/test_template.docx")
        print("  - templates/test_template_orig.docx")
        sys.exit(1)
    
    if not os.path.exists(template_path):
        print(f"Error: Template file '{template_path}' not found.")
        print("Please ensure the template file exists and the path is correct.")
        sys.exit(1)
    
    # Create output directory if it doesn't exist
    # If we're in src directory, go up one level to project root
    if os.path.basename(os.getcwd()) == "src":
        output_dir = os.path.join(os.path.dirname(os.getcwd()), "output")
    else:
        output_dir = os.path.join(output_dir, "output")
    os.makedirs(output_dir, exist_ok=True)
    
    # Initialize extractor with template if provided
    extractor = SpecContentExtractorV3(template_path)
    
    # Extract content
    print(f"Extracting content from '{docx_path}'...")
    data = extractor.extract_content(docx_path)
    
    if not data:
        print("Error: Failed to extract content.")
        sys.exit(1)
    
    # Generate output filenames
    base_name = os.path.splitext(os.path.basename(docx_path))[0]
    main_json_path = os.path.join(output_dir, f"{base_name}_v3.json")
    error_path = os.path.join(output_dir, f"{base_name}_v3_errors.txt")
    
    # Save main comprehensive JSON (contains everything)
    extractor.save_to_json(data, main_json_path)
    print(f"Main content saved to: {main_json_path}")
    
    # Save separate modular JSON files
    extractor.save_modular_json_files(data, base_name, output_dir)
    
    # Generate and save error report
    error_report = extractor.generate_error_report()
    extractor.save_error_report(error_report, error_path)
    print(f"Error report saved to: {error_path}")
    
    # Print summary
    print(f"\nExtraction Summary:")
    print(f"- Content blocks found: {len(data.get('content_blocks', []))}")
    print(f"- Header paragraphs: {len(data.get('header', {}).get('paragraphs', []))}")
    print(f"- Footer paragraphs: {len(data.get('footer', {}).get('paragraphs', []))}")
    print(f"- Comments found: {len(data.get('comments', []))}")
    print(f"- BWA list levels: {len(data.get('template_info', {}).get('bwa_list_levels', {}))}")
    print(f"- Errors found: {len(extractor.errors)}")
    
    if extractor.errors:
        print(f"\nWARNING: {len(extractor.errors)} errors were found during extraction.")
        print(f"Please review the error report: {error_path}")
    else:
        print("\nExtraction completed successfully with no errors.")

if __name__ == "__main__":
    main() 