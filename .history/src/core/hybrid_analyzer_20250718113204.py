"""
Hybrid Analyzer for SpecConverter v1.0

Combines PDF extraction, template analysis, and source document parsing
to validate and correct numbering using cross-reference analysis.
"""

import re
import logging
from typing import Dict, List, Optional, Tuple, Any
from pathlib import Path

from .models import ContentBlock, SpecDocument, HeaderFooterData, ValidationResults
from .extractor import SpecContentExtractorV3 as SpecExtractor
from .template_analyzer import TemplateListDetector as TemplateAnalyzer
from .pdf_extractor import PDFExtractor


class HybridAnalyzer:
    """
    Hybrid analyzer that combines multiple extraction methods for validation.
    
    Uses PDF extraction as ground truth, template analysis for expected numbering,
    and source document parsing for detailed structure, then cross-references
    all three to validate and correct numbering.
    """
    
    def __init__(self, template_path: Optional[str] = None):
        """
        Initialize the hybrid analyzer.
        
        Args:
            template_path: Optional template path for analysis
        """
        self.template_path = template_path
        self.logger = logging.getLogger(__name__)
        
        # Initialize components
        self.pdf_extractor = PDFExtractor()
        self.spec_extractor = SpecExtractor(template_path=template_path)
        self.template_analyzer = TemplateAnalyzer()
        
        # Analysis results
        self.pdf_content = ""
        self.template_analysis = None
        self.source_blocks = []
        self.numbering_patterns = []
        self.validation_results = []
    
    def analyze_document(self, docx_path: str, template_path: Optional[str] = None) -> SpecDocument:
        """
        Perform comprehensive hybrid analysis of a document.
        
        Args:
            docx_path: Path to the Word document
            template_path: Optional template path (overrides instance template)
            
        Returns:
            Validated and corrected SpecDocument
        """
        self.logger.info(f"Starting hybrid analysis of: {docx_path}")
        
        # Step 1: Extract PDF content as ground truth
        self.logger.info("Step 1: Extracting PDF content as ground truth...")
        pdf_document = self.pdf_extractor.extract_document(docx_path, template_path)
        self.pdf_content = self._get_pdf_text_content(pdf_document)
        
        # Step 2: Analyze template for expected numbering patterns
        self.logger.info("Step 2: Analyzing template for numbering patterns...")
        template_path = template_path or self.template_path
        if template_path:
            self.template_analysis = self.template_analyzer.analyze_template(template_path)
            self.numbering_patterns = self._extract_numbering_patterns(self.template_analysis)
        
        # Step 3: Parse source document for detailed structure
        self.logger.info("Step 3: Parsing source document for detailed structure...")
        source_document = self.spec_extractor.extract_content(docx_path)
        self.source_blocks = self._get_source_blocks(source_document)
        
        # Step 4: Cross-reference and validate numbering
        self.logger.info("Step 4: Cross-referencing and validating numbering...")
        validated_blocks = self._cross_reference_numbering()
        
        # Step 5: Create final validated document
        self.logger.info("Step 5: Creating final validated document...")
        final_document = self._create_validated_document(docx_path, validated_blocks)
        
        self.logger.info(f"Hybrid analysis complete. Validated {len(validated_blocks)} blocks.")
        return final_document
    
    def _get_pdf_text_content(self, pdf_document: SpecDocument) -> str:
        """Extract all text content from PDF document."""
        content_parts = []
        for block in pdf_document.content_blocks:
            content_parts.append(block.text)
            if block.content:
                content_parts.append(block.content)
        return '\n'.join(content_parts)
    
    def _extract_numbering_patterns(self, template_analysis: Any) -> List[Dict[str, Any]]:
        """Extract numbering patterns from template analysis."""
        patterns = []
        
        if not template_analysis or not template_analysis.numbering_definitions:
            return patterns
        
        # Handle different template analysis structures
        numbering_defs = template_analysis.numbering_definitions
        if isinstance(numbering_defs, dict):
            for num_id, num_def in numbering_defs.items():
                if isinstance(num_def, dict) and 'levels' in num_def:
                    for level in num_def['levels']:
                        if isinstance(level, dict):
                            pattern = {
                                'num_id': num_id,
                                'level': level.get('level', 0),
                                'format': level.get('format', ''),
                                'text': level.get('text', ''),
                                'start_at': level.get('start_at', 1),
                                'bwa_level_name': level.get('bwa_level_name', '')
                            }
                            patterns.append(pattern)
        
        # Also check BWA list levels if available
        if template_analysis.bwa_list_levels:
            for level_name, level_info in template_analysis.bwa_list_levels.items():
                if hasattr(level_info, 'numbering_id'):  # Handle ListLevelInfo objects
                    pattern = {
                        'num_id': level_info.numbering_id or '',
                        'level': getattr(level_info, 'level_number', 0),
                        'format': level_info.number_format or '',
                        'text': level_info.level_text or '',
                        'start_at': level_info.start_value or 1,
                        'bwa_level_name': level_name
                    }
                    patterns.append(pattern)
                elif isinstance(level_info, dict):
                    pattern = {
                        'num_id': level_info.get('numbering_id', ''),
                        'level': level_info.get('level', 0),
                        'format': level_info.get('format', ''),
                        'text': level_info.get('text', ''),
                        'start_at': level_info.get('start_at', 1),
                        'bwa_level_name': level_name
                    }
                    patterns.append(pattern)
        
        return patterns
    
    def _get_source_blocks(self, source_document: Dict[str, Any]) -> List[ContentBlock]:
        """Extract content blocks from source document."""
        blocks = []
        
        # Handle different source document formats
        if 'content_blocks' in source_document:
            for block_data in source_document['content_blocks']:
                block = ContentBlock(
                    text=block_data.get('text', ''),
                    level_type=block_data.get('level_type', ''),
                    number=block_data.get('number'),
                    content=block_data.get('content', ''),
                    level_number=block_data.get('level_number'),
                    bwa_level_name=block_data.get('bwa_level_name'),
                    numbering_id=block_data.get('numbering_id'),
                    numbering_level=block_data.get('numbering_level'),
                    style_name=block_data.get('style_name')
                )
                blocks.append(block)
        
        return blocks
    
    def _cross_reference_numbering(self) -> List[ContentBlock]:
        """Cross-reference numbering between PDF content and source blocks."""
        validated_blocks = []
        
        for block in self.source_blocks:
            self.logger.debug(f"Validating block: {block.text[:50]}...")
            
            # Find this block's text in PDF content
            pdf_match = self._find_text_in_pdf(block.text)
            
            if pdf_match:
                # Extract numbering from PDF context
                pdf_numbering = self._extract_numbering_from_pdf_context(pdf_match)
                
                # Validate against template patterns
                validated_numbering = self._validate_numbering_against_template(
                    block, pdf_numbering
                )
                
                # Create validated block
                validated_block = self._create_validated_block(block, validated_numbering)
                validated_blocks.append(validated_block)
                
                self.logger.debug(f"✓ Validated: {block.text[:30]}... -> {validated_numbering}")
            else:
                # Keep original block if not found in PDF
                self.logger.warning(f"⚠ Block not found in PDF: {block.text[:50]}...")
                validated_blocks.append(block)
        
        return validated_blocks
    
    def _find_text_in_pdf(self, text: str) -> Optional[Tuple[int, int]]:
        """Find text in PDF content and return position."""
        if not text or not self.pdf_content:
            return None
        
        # Clean text for comparison
        clean_text = self._clean_text_for_comparison(text)
        clean_pdf = self._clean_text_for_comparison(self.pdf_content)
        
        # Find match
        start_pos = clean_pdf.find(clean_text)
        if start_pos != -1:
            end_pos = start_pos + len(clean_text)
            return (start_pos, end_pos)
        
        return None
    
    def _clean_text_for_comparison(self, text: str) -> str:
        """Clean text for comparison by removing extra whitespace and normalizing."""
        # Remove extra whitespace
        text = re.sub(r'\s+', ' ', text.strip())
        # Normalize quotes and dashes
        text = text.replace('"', '"').replace('"', '"').replace('–', '-').replace('—', '-')
        return text
    
    def _extract_numbering_from_pdf_context(self, pdf_match: Tuple[int, int]) -> Optional[str]:
        """Extract numbering from PDF context around the matched text."""
        start_pos, end_pos = pdf_match
        
        # Look for numbering patterns before the matched text with larger context
        context_before = self.pdf_content[max(0, start_pos-500):start_pos]
        
        # More specific and hierarchical numbering patterns
        patterns = [
            # Section-level patterns (highest priority)
            (r'SECTION\s+(\d+\s+\d+\s+\d+)', 10),  # SECTION 26 05 00
            (r'(\d+\.\d+\.\d+)', 9),              # 26.05.00
            (r'(\d+\s+\d+\s+\d+)', 8),            # 26 05 00
            
            # Division/Part patterns
            (r'DIVISION\s+(\d+)', 7),             # DIVISION 26
            (r'PART\s+(\d+)', 6),                 # PART 1
            
            # Subsection patterns
            (r'(\d+\.\d+)', 5),                   # 2.01
            (r'(\d+\.)', 4),                      # 2.
            
            # Item patterns (lower priority)
            (r'([A-Z]\.)', 3),                    # A.
            (r'(\d+\.)', 2),                      # 1.
        ]
        
        best_match = None
        best_priority = 0
        best_position = -1
        
        for pattern, priority in patterns:
            matches = re.finditer(pattern, context_before, re.IGNORECASE)
            for match in matches:
                # Check if this is a complete word/numbering (not a fragment)
                if self._is_complete_numbering(match.group(1), context_before, match.start()):
                    # Prefer matches closer to our text and with higher priority
                    if priority > best_priority or (priority == best_priority and match.start() > best_position):
                        best_match = match.group(1)
                        best_priority = priority
                        best_position = match.start()
        
        return best_match
    
    def _is_complete_numbering(self, numbering: str, context: str, position: int) -> bool:
        """Check if the numbering is complete and not a word fragment."""
        if not numbering:
            return False
        
        # Get characters before and after the numbering
        before_char = context[position - 1] if position > 0 else ' '
        after_char = context[position + len(numbering)] if position + len(numbering) < len(context) else ' '
        
        # Check for word boundaries
        word_boundary_chars = ' \t\n\r.,;:!?()[]{}'
        
        # Numbering should be preceded and followed by word boundaries
        if before_char not in word_boundary_chars or after_char not in word_boundary_chars:
            return False
        
        # Additional validation for specific patterns
        if numbering in ['s', 't', 'd', 'e', 'n', 'c'] and len(numbering) == 1:
            # Single letters are likely fragments unless they're followed by a period
            return after_char == '.'
        
        # Validate numbering format
        if re.match(r'^\d+\.\d+\.\d+$', numbering):  # 26.05.00
            return True
        if re.match(r'^\d+\s+\d+\s+\d+$', numbering):  # 26 05 00
            return True
        if re.match(r'^\d+\.\d+$', numbering):  # 2.01
            return True
        if re.match(r'^\d+\.$', numbering):  # 2.
            return True
        if re.match(r'^[A-Z]\.$', numbering):  # A.
            return True
        if re.match(r'^\d+\.$', numbering):  # 1.
            return True
        
        return True
    
    def _validate_numbering_against_template(self, block: ContentBlock, pdf_numbering: Optional[str]) -> Optional[str]:
        """Validate numbering against template patterns with confidence scoring."""
        if not pdf_numbering:
            return block.number
        
        # Calculate confidence score for the PDF numbering
        confidence = self._calculate_numbering_confidence(pdf_numbering, block)
        
        # Only apply correction if confidence is high enough
        if confidence < 0.7:  # 70% confidence threshold
            self.logger.debug(f"⚠ Low confidence ({confidence:.2f}) for numbering '{pdf_numbering}', keeping original '{block.number}'")
            return block.number
        
        # Check if the PDF numbering matches expected patterns
        template_match = False
        for pattern in self.numbering_patterns:
            if self._numbering_matches_pattern(pdf_numbering, pattern):
                self.logger.debug(f"✓ Numbering matches template pattern: {pdf_numbering} (confidence: {confidence:.2f})")
                template_match = True
                break
        
        # Apply correction if confident enough
        if pdf_numbering != block.number:
            if template_match:
                self.logger.info(f"📝 High confidence template match: '{block.number}' -> '{pdf_numbering}' (confidence: {confidence:.2f})")
            else:
                self.logger.info(f"📝 High confidence PDF numbering: '{block.number}' -> '{pdf_numbering}' (confidence: {confidence:.2f})")
            return pdf_numbering
        
        return block.number
    
    def _calculate_numbering_confidence(self, numbering: str, block: ContentBlock) -> float:
        """Calculate confidence score for a numbering match."""
        confidence = 0.0
        
        # Base confidence on numbering format
        if re.match(r'^\d+\.\d+\.\d+$', numbering):  # 26.05.00
            confidence += 0.4
        elif re.match(r'^\d+\s+\d+\s+\d+$', numbering):  # 26 05 00
            confidence += 0.4
        elif re.match(r'^\d+\.\d+$', numbering):  # 2.01
            confidence += 0.3
        elif re.match(r'^\d+\.$', numbering):  # 2.
            confidence += 0.2
        elif re.match(r'^[A-Z]\.$', numbering):  # A.
            confidence += 0.2
        elif re.match(r'^\d+\.$', numbering):  # 1.
            confidence += 0.2
        
        # Bonus for section-level patterns
        if 'SECTION' in numbering.upper() or re.match(r'^\d+\s+\d+\s+\d+$', numbering):
            confidence += 0.3
        
        # Bonus for matching block level type
        if block.level_type == 'section' and re.match(r'^\d+\s+\d+\s+\d+$', numbering):
            confidence += 0.2
        elif block.level_type == 'part_title' and re.match(r'^\d+\.$', numbering):
            confidence += 0.2
        elif block.level_type == 'subsection_title' and re.match(r'^\d+\.\d+$', numbering):
            confidence += 0.2
        elif block.level_type == 'item' and re.match(r'^[A-Z]\.$', numbering):
            confidence += 0.2
        
        # Penalty for single character fragments
        if len(numbering) == 1 and numbering in ['s', 't', 'd', 'e', 'n', 'c']:
            confidence -= 0.5
        
        return min(confidence, 1.0)  # Cap at 1.0
    
    def _numbering_matches_pattern(self, numbering: str, pattern: Dict[str, Any]) -> bool:
        """Check if numbering matches a template pattern."""
        if not numbering or not pattern:
            return False
        
        # Simple pattern matching - could be enhanced
        format_pattern = pattern.get('format', '')
        bwa_level = pattern.get('bwa_level_name', '')
        
        # Check if numbering format matches
        if format_pattern:
            try:
                if re.match(format_pattern, numbering, re.IGNORECASE):
                    return True
            except re.error:
                pass
        
        # Check BWA level patterns
        if bwa_level:
            if 'SECTION' in bwa_level and 'SECTION' in numbering.upper():
                return True
            if 'PART' in bwa_level and 'PART' in numbering.upper():
                return True
            if 'DIVISION' in bwa_level and 'DIVISION' in numbering.upper():
                return True
        
        return False
    
    def _create_validated_block(self, original_block: ContentBlock, validated_numbering: Optional[str]) -> ContentBlock:
        """Create a validated block with corrected numbering."""
        # Create new block with validated numbering
        validated_block = ContentBlock(
            text=original_block.text,
            level_type=original_block.level_type,
            number=validated_numbering,
            content=original_block.content,
            level_number=original_block.level_number,
            bwa_level_name=original_block.bwa_level_name,
            numbering_id=original_block.numbering_id,
            numbering_level=original_block.numbering_level,
            style_name=original_block.style_name,
            # Copy all styling attributes
            font_name=original_block.font_name,
            font_size=original_block.font_size,
            font_bold=original_block.font_bold,
            font_italic=original_block.font_italic,
            font_underline=original_block.font_underline,
            font_color=original_block.font_color,
            font_strike=original_block.font_strike,
            font_small_caps=original_block.font_small_caps,
            font_all_caps=original_block.font_all_caps,
            paragraph_alignment=original_block.paragraph_alignment,
            paragraph_indent_left=original_block.paragraph_indent_left,
            paragraph_indent_right=original_block.paragraph_indent_right,
            paragraph_indent_first_line=original_block.paragraph_indent_first_line,
            paragraph_spacing_before=original_block.paragraph_spacing_before,
            paragraph_spacing_after=original_block.paragraph_spacing_after,
            paragraph_line_spacing=original_block.paragraph_line_spacing,
            paragraph_line_spacing_rule=original_block.paragraph_line_spacing_rule,
            paragraph_keep_with_next=original_block.paragraph_keep_with_next,
            paragraph_keep_lines_together=original_block.paragraph_keep_lines_together,
            paragraph_page_break_before=original_block.paragraph_page_break_before,
            paragraph_widow_control=original_block.paragraph_widow_control,
            paragraph_dont_add_space_between_same_style=original_block.paragraph_dont_add_space_between_same_style,
            number_alignment=original_block.number_alignment,
            aligned_at=original_block.aligned_at,
            text_indent_at=original_block.text_indent_at,
            follow_number_with=original_block.follow_number_with,
            add_tab_stop_at=original_block.add_tab_stop_at,
            link_level_to_style=original_block.link_level_to_style,
            used_fallback_styling=original_block.used_fallback_styling
        )
        
        # Record validation result
        if validated_numbering != original_block.number:
            confidence = self._calculate_numbering_confidence(validated_numbering or '', original_block)
            self.validation_results.append({
                'block_text': original_block.text[:50],
                'original_number': original_block.number,
                'validated_number': validated_numbering,
                'confidence': confidence,
                'validation_type': 'numbering_correction'
            })
        
        return validated_block
    
    def _create_validated_document(self, docx_path: str, validated_blocks: List[ContentBlock]) -> SpecDocument:
        """Create the final validated document."""
        # Create validation results
        validation_results = ValidationResults(
            errors=[],
            corrections=self.validation_results,
            validation_summary={
                'total_blocks': len(validated_blocks),
                'blocks_validated': len([r for r in self.validation_results if r['validation_type'] == 'numbering_correction']),
                'validation_method': 'hybrid_cross_reference'
            }
        )
        
        # Create document
        document = SpecDocument(
            file_path=docx_path,
            content_blocks=validated_blocks,
            header_footer=HeaderFooterData(
                header={},
                footer={},
                margins={},
                document_settings={}
            ),
            template_analysis=None,  # We'll handle template analysis separately
            validation_results=validation_results
        )
        
        return document
    
    def get_validation_report(self) -> Dict[str, Any]:
        """Get a detailed validation report."""
        return {
            'pdf_content_length': len(self.pdf_content),
            'template_patterns': len(self.numbering_patterns),
            'source_blocks': len(self.source_blocks),
            'validation_results': self.validation_results,
            'summary': {
                'total_blocks_processed': len(self.source_blocks),
                'blocks_with_numbering_corrections': len([r for r in self.validation_results if r['validation_type'] == 'numbering_correction']),
                'blocks_not_found_in_pdf': len([r for r in self.validation_results if r['validation_type'] == 'not_found_in_pdf'])
            }
        }


def analyze_with_hybrid_validation(docx_path: str, template_path: Optional[str] = None) -> SpecDocument:
    """
    Convenience function for hybrid analysis with validation.
    
    Args:
        docx_path: Path to Word document
        template_path: Optional template path
        
    Returns:
        Validated SpecDocument
    """
    analyzer = HybridAnalyzer(template_path)
    return analyzer.analyze_document(docx_path, template_path) 