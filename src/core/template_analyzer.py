#!/usr/bin/env python3
"""
Template List Level Detector Module

This module detects and analyzes list level items present in Word document templates.
It extracts numbering definitions, BWA list levels, and provides detailed analysis
of the template's hierarchical structure.

Features:
- Extract numbering definitions from template's numbering.xml
- Detect BWA-labeled list levels
- Analyze list level patterns and properties
- Map numbering IDs to abstract numbering definitions
- Provide comprehensive template analysis

Usage:
    from template_list_detector import TemplateListDetector
    
    detector = TemplateListDetector()
    analysis = detector.analyze_template(template_path)
"""

from docx import Document
import json
import os
import zipfile
import xml.etree.ElementTree as ET
from typing import Dict, List, Optional, Any, Tuple
from dataclasses import dataclass
from datetime import datetime

@dataclass
class ListLevelInfo:
    """Represents information about a list level"""
    level_number: int
    numbering_id: Optional[str] = None
    abstract_num_id: Optional[str] = None
    level_text: Optional[str] = None
    number_format: Optional[str] = None
    start_value: Optional[str] = None
    suffix: Optional[str] = None
    justification: Optional[str] = None
    style_name: Optional[str] = None
    bwa_label: Optional[str] = None
    is_bwa_level: bool = False

@dataclass
class TemplateAnalysis:
    """Represents complete template analysis"""
    template_path: str
    analysis_timestamp: str
    numbering_definitions: Dict[str, Any]
    bwa_list_levels: Dict[str, ListLevelInfo]
    level_mappings: Dict[str, str]
    summary: Dict[str, Any]

class TemplateListDetector:
    """Detects and analyzes list level items in Word document templates"""
    
    def __init__(self):
        self.namespace = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    
    def analyze_template(self, template_path: str) -> TemplateAnalysis:
        """
        Perform comprehensive analysis of template list levels
        
        Args:
            template_path: Path to the template Word document
            
        Returns:
            TemplateAnalysis object with complete analysis
        """
        try:
            print(f"Analyzing template list levels: {template_path}")
            
            # Extract numbering definitions
            numbering_definitions = self.extract_numbering_definitions(template_path)
            
            # Find BWA list levels
            bwa_list_levels = self.find_bwa_list_levels(template_path, numbering_definitions)
            
            # Create level mappings
            level_mappings = self.create_level_mappings(numbering_definitions)
            
            # Generate summary
            summary = self.generate_summary(numbering_definitions, bwa_list_levels, level_mappings)
            
            analysis = TemplateAnalysis(
                template_path=template_path,
                analysis_timestamp=datetime.now().isoformat(),
                numbering_definitions=numbering_definitions,
                bwa_list_levels=bwa_list_levels,
                level_mappings=level_mappings,
                summary=summary
            )
            
            print(f"Template analysis complete: {len(bwa_list_levels)} BWA list levels found")
            return analysis
            
        except Exception as e:
            print(f"Error analyzing template: {e}")
            return self.create_empty_analysis(template_path)
    
    def extract_numbering_definitions(self, template_path: str) -> Dict[str, Any]:
        """
        Extract numbering definitions from template's numbering.xml
        
        Args:
            template_path: Path to the template Word document
            
        Returns:
            Dictionary containing numbering definitions
        """
        numbering_definitions = {}
        
        try:
            with zipfile.ZipFile(template_path) as zf:
                if "word/numbering.xml" in zf.namelist():
                    num_xml = zf.read("word/numbering.xml")
                    root = ET.fromstring(num_xml)
                    
                    # Extract abstract numbering definitions
                    for abstract_num in root.findall(".//w:abstractNum", self.namespace):
                        abstract_num_id = abstract_num.get(f"{{{self.namespace['w']}}}abstractNumId")
                        numbering_definitions[abstract_num_id] = {
                            "levels": {},
                            "bwa_label": None,
                            "nsid": abstract_num.get(f"{{{self.namespace['w']}}}nsid"),
                            "multiLevelType": abstract_num.get(f"{{{self.namespace['w']}}}multiLevelType"),
                            "tmpl": abstract_num.get(f"{{{self.namespace['w']}}}tmpl")
                        }
                        
                        # Check for BWA label in first level
                        first_level = abstract_num.find("w:lvl", self.namespace)
                        if first_level is not None:
                            lvl_text_elem = first_level.find("w:lvlText", self.namespace)
                            if lvl_text_elem is not None:
                                lvl_text = lvl_text_elem.get(f"{{{self.namespace['w']}}}val", "")
                                if "BWA" in lvl_text.upper():
                                    numbering_definitions[abstract_num_id]["bwa_label"] = lvl_text
                        
                        # Extract all levels
                        for lvl in abstract_num.findall("w:lvl", self.namespace):
                            ilvl = lvl.get(f"{{{self.namespace['w']}}}ilvl")
                            level_data = self.extract_level_data(lvl)
                            numbering_definitions[abstract_num_id]["levels"][ilvl] = level_data
                    
                    # Extract num mappings
                    for num_elem in root.findall(".//w:num", self.namespace):
                        num_id = num_elem.get(f"{{{self.namespace['w']}}}numId")
                        abstract_num_ref = num_elem.find("w:abstractNumId", self.namespace)
                        if abstract_num_ref is not None:
                            abstract_num_id = abstract_num_ref.get(f"{{{self.namespace['w']}}}val")
                            numbering_definitions[f"num_{num_id}"] = {
                                "abstract_num_id": abstract_num_id,
                                "bwa_label": numbering_definitions.get(abstract_num_id, {}).get("bwa_label")
                            }
                            
        except Exception as e:
            print(f"Error extracting numbering definitions: {e}")
        
        return numbering_definitions
    
    def extract_level_data(self, level_element: Any) -> Dict[str, Any]:
        """
        Extract detailed data from a level element
        
        Args:
            level_element: XML element representing a level
            
        Returns:
            Dictionary containing level data
        """
        level_data = {
            "ilvl": level_element.get(f"{{{self.namespace['w']}}}ilvl"),
            "lvlText": None,
            "numFmt": None,
            "start": None,
            "suff": None,
            "lvlJc": None,
            "pStyle": None,
            "pPr": {},
            "rPr": {}
        }
        
        # Extract lvlText (the pattern like "%1.0", "%1.%2")
        lvl_text_elem = level_element.find("w:lvlText", self.namespace)
        if lvl_text_elem is not None:
            level_data["lvlText"] = lvl_text_elem.get(f"{{{self.namespace['w']}}}val")
        
        # Extract numFmt (decimal, lowerLetter, upperLetter, etc.)
        num_fmt_elem = level_element.find("w:numFmt", self.namespace)
        if num_fmt_elem is not None:
            level_data["numFmt"] = num_fmt_elem.get(f"{{{self.namespace['w']}}}val")
        
        # Extract start value
        start_elem = level_element.find("w:start", self.namespace)
        if start_elem is not None:
            level_data["start"] = start_elem.get(f"{{{self.namespace['w']}}}val")
        
        # Extract suffix (tab, space, nothing)
        suff_elem = level_element.find("w:suff", self.namespace)
        if suff_elem is not None:
            level_data["suff"] = suff_elem.get(f"{{{self.namespace['w']}}}val")
        
        # Extract lvlJc (justification: left, center, right)
        lvl_jc_elem = level_element.find("w:lvlJc", self.namespace)
        if lvl_jc_elem is not None:
            level_data["lvlJc"] = lvl_jc_elem.get(f"{{{self.namespace['w']}}}val")
        
        # Extract pStyle (paragraph style)
        p_style_elem = level_element.find("w:pStyle", self.namespace)
        if p_style_elem is not None:
            level_data["pStyle"] = p_style_elem.get(f"{{{self.namespace['w']}}}val")
        
        # Extract paragraph properties (pPr)
        p_pr_elem = level_element.find("w:pPr", self.namespace)
        if p_pr_elem is not None:
            level_data["pPr"] = self.extract_paragraph_properties(p_pr_elem)
        
        # Extract run properties (rPr) - font info
        r_pr_elem = level_element.find("w:rPr", self.namespace)
        if r_pr_elem is not None:
            level_data["rPr"] = self.extract_run_properties(r_pr_elem)
        
        return level_data
    
    def extract_paragraph_properties(self, p_pr_elem: Any) -> Dict[str, Any]:
        """Extract paragraph properties from element"""
        p_pr = {}
        
        # Indentation
        indent_elem = p_pr_elem.find("w:ind", self.namespace)
        if indent_elem is not None:
            p_pr["indent"] = {
                "left": indent_elem.get(f"{{{self.namespace['w']}}}left"),
                "right": indent_elem.get(f"{{{self.namespace['w']}}}right"),
                "hanging": indent_elem.get(f"{{{self.namespace['w']}}}hanging"),
                "firstLine": indent_elem.get(f"{{{self.namespace['w']}}}firstLine")
            }
        
        # Spacing
        spacing_elem = p_pr_elem.find("w:spacing", self.namespace)
        if spacing_elem is not None:
            p_pr["spacing"] = {
                "before": spacing_elem.get(f"{{{self.namespace['w']}}}before"),
                "after": spacing_elem.get(f"{{{self.namespace['w']}}}after"),
                "line": spacing_elem.get(f"{{{self.namespace['w']}}}line"),
                "lineRule": spacing_elem.get(f"{{{self.namespace['w']}}}lineRule")
            }
        
        # Tab stops
        tabs_elem = p_pr_elem.find("w:tabs", self.namespace)
        if tabs_elem is not None:
            tabs = {"tab": []}
            for tab_elem in tabs_elem.findall("w:tab", self.namespace):
                tab_info = {
                    "pos": tab_elem.get(f"{{{self.namespace['w']}}}pos"),
                    "val": tab_elem.get(f"{{{self.namespace['w']}}}val"),
                    "leader": tab_elem.get(f"{{{self.namespace['w']}}}leader")
                }
                tabs["tab"].append(tab_info)
            p_pr["tabs"] = tabs
        
        return p_pr
    
    def extract_run_properties(self, r_pr_elem: Any) -> Dict[str, Any]:
        """Extract run properties from element"""
        r_pr = {}
        
        # Font family
        r_fonts_elem = r_pr_elem.find("w:rFonts", self.namespace)
        if r_fonts_elem is not None:
            r_pr["rFonts"] = {
                "ascii": r_fonts_elem.get(f"{{{self.namespace['w']}}}ascii"),
                "hAnsi": r_fonts_elem.get(f"{{{self.namespace['w']}}}hAnsi"),
                "eastAsia": r_fonts_elem.get(f"{{{self.namespace['w']}}}eastAsia"),
                "cs": r_fonts_elem.get(f"{{{self.namespace['w']}}}cs")
            }
        
        # Font size
        sz_elem = r_pr_elem.find("w:sz", self.namespace)
        if sz_elem is not None:
            r_pr["sz"] = sz_elem.get(f"{{{self.namespace['w']}}}val")
        
        # Bold
        b_elem = r_pr_elem.find("w:b", self.namespace)
        if b_elem is not None:
            r_pr["bold"] = b_elem.get(f"{{{self.namespace['w']}}}val")
        
        # Italic
        i_elem = r_pr_elem.find("w:i", self.namespace)
        if i_elem is not None:
            r_pr["italic"] = i_elem.get(f"{{{self.namespace['w']}}}val")
        
        return r_pr
    
    def find_bwa_list_levels(self, template_path: str, numbering_definitions: Dict[str, Any]) -> Dict[str, ListLevelInfo]:
        """
        Find BWA-labeled list levels in template
        
        Args:
            template_path: Path to the template Word document
            numbering_definitions: Dictionary of numbering definitions
            
        Returns:
            Dictionary mapping level identifiers to ListLevelInfo objects
        """
        bwa_list_levels = {}
        
        try:
            doc = Document(template_path)
            
            # Find BWA levels by style name
            for paragraph in doc.paragraphs:
                style_name = getattr(paragraph.style, "name", None)
                if style_name and "bwa" in style_name.lower():
                    level = self.get_paragraph_level(paragraph)
                    numbering_id = self.get_paragraph_numbering_id(paragraph)
                    
                    list_level_info = ListLevelInfo(
                        level_number=level if level is not None else -1,
                        numbering_id=numbering_id,
                        style_name=getattr(paragraph.style, "name", ""),
                        is_bwa_level=True
                    )
                    # Only add to bwa_list_levels if style_name is not None
                    if style_name:
                        bwa_list_levels[style_name] = list_level_info

                    # Map by numbering ID if available
                    if numbering_id:
                        bwa_list_levels[f"num_{numbering_id}"] = list_level_info
                        
                        # Get abstract numbering info
                        if f"num_{numbering_id}" in numbering_definitions:
                            abstract_num_id = numbering_definitions[f"num_{numbering_id}"]["abstract_num_id"]
                            if abstract_num_id in numbering_definitions:
                                abstract_info = numbering_definitions[abstract_num_id]
                                list_level_info.abstract_num_id = abstract_num_id
                                list_level_info.bwa_label = abstract_info.get("bwa_label")
                                
                                # Add level-specific info
                                if str(level) in abstract_info.get("levels", {}):
                                    level_info = abstract_info["levels"][str(level)]
                                    list_level_info.level_text = level_info.get("lvlText")
                                    list_level_info.number_format = level_info.get("numFmt")
                                    list_level_info.start_value = level_info.get("start")
                                    list_level_info.suffix = level_info.get("suff")
                                    list_level_info.justification = level_info.get("lvlJc")
            
            # ENHANCED: Find BWA levels by numbering definitions and link them to style names
            for abstract_num_id, abstract_info in numbering_definitions.items():
                # Skip num_ entries (these are mappings, not abstract definitions)
                if abstract_num_id.startswith("num_"):
                    continue
                    
                for level_num, level_info in abstract_info.get("levels", {}).items():
                    p_style = level_info.get("pStyle")
                    if p_style and "bwa" in p_style.lower():
                        # This is a BWA style linked to a numbering level
                        level_number = int(level_num)
                        
                        # Create or update the BWA level info
                        if p_style in bwa_list_levels:
                            # Update existing entry
                            existing_info = bwa_list_levels[p_style]
                            existing_info.level_number = level_number
                            existing_info.abstract_num_id = abstract_num_id
                            existing_info.level_text = level_info.get("lvlText")
                            existing_info.number_format = level_info.get("numFmt")
                            existing_info.start_value = level_info.get("start")
                            existing_info.suffix = level_info.get("suff")
                            existing_info.justification = level_info.get("lvlJc")
                        else:
                            # Create new entry
                            list_level_info = ListLevelInfo(
                                level_number=level_number,
                                abstract_num_id=abstract_num_id,
                                level_text=level_info.get("lvlText"),
                                number_format=level_info.get("numFmt"),
                                start_value=level_info.get("start"),
                                suffix=level_info.get("suff"),
                                justification=level_info.get("lvlJc"),
                                style_name=p_style,
                                is_bwa_level=True
                            )
                            bwa_list_levels[p_style] = list_level_info
                        
                        print(f"DEBUG: Linked BWA style '{p_style}' to level {level_number} with format '{level_info.get('numFmt')}'")
                        
        except Exception as e:
            print(f"Error finding BWA list levels: {e}")
        
        return bwa_list_levels
    
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
    
    def create_level_mappings(self, numbering_definitions: Dict[str, Any]) -> Dict[str, str]:
        """
        Create mappings between numbering IDs and abstract numbering definitions
        
        Args:
            numbering_definitions: Dictionary of numbering definitions
            
        Returns:
            Dictionary mapping numbering IDs to abstract numbering IDs
        """
        level_mappings = {}
        
        for key, value in numbering_definitions.items():
            if key.startswith("num_"):
                num_id = key.replace("num_", "")
                if "abstract_num_id" in value:
                    level_mappings[num_id] = value["abstract_num_id"]
        
        return level_mappings
    
    def generate_summary(self, numbering_definitions: Dict[str, Any], 
                        bwa_list_levels: Dict[str, ListLevelInfo],
                        level_mappings: Dict[str, str]) -> Dict[str, Any]:
        """
        Generate summary statistics for the template analysis
        
        Args:
            numbering_definitions: Dictionary of numbering definitions
            bwa_list_levels: Dictionary of BWA list levels
            level_mappings: Dictionary of level mappings
            
        Returns:
            Dictionary containing summary statistics
        """
        # Count different types of definitions
        abstract_nums = sum(1 for k in numbering_definitions.keys() if not k.startswith("num_"))
        num_mappings = sum(1 for k in numbering_definitions.keys() if k.startswith("num_"))
        bwa_levels = sum(1 for level in bwa_list_levels.values() if level.is_bwa_level)
        
        # Count levels by type
        level_types = {}
        for level in bwa_list_levels.values():
            if level.number_format:
                level_types[level.number_format] = level_types.get(level.number_format, 0) + 1
        
        return {
            "total_abstract_numbering": abstract_nums,
            "total_num_mappings": num_mappings,
            "total_bwa_levels": bwa_levels,
            "level_mappings_count": len(level_mappings),
            "level_types": level_types,
            "analysis_timestamp": datetime.now().isoformat()
        }
    
    def create_empty_analysis(self, template_path: str) -> TemplateAnalysis:
        """Create empty analysis when template analysis fails"""
        return TemplateAnalysis(
            template_path=template_path,
            analysis_timestamp=datetime.now().isoformat(),
            numbering_definitions={},
            bwa_list_levels={},
            level_mappings={},
            summary={
                "total_abstract_numbering": 0,
                "total_num_mappings": 0,
                "total_bwa_levels": 0,
                "level_mappings_count": 0,
                "level_types": {},
                "analysis_timestamp": datetime.now().isoformat(),
                "error": "Template analysis failed"
            }
        )
    
    def save_analysis_to_json(self, analysis: TemplateAnalysis, output_path: str):
        """
        Save template analysis to JSON file
        
        Args:
            analysis: TemplateAnalysis object
            output_path: Path to save the JSON file
        """
        # Convert dataclass to dictionary for JSON serialization
        analysis_dict = {
            "template_path": analysis.template_path,
            "analysis_timestamp": analysis.analysis_timestamp,
            "numbering_definitions": analysis.numbering_definitions,
            "bwa_list_levels": {
                key: {
                    "level_number": level.level_number,
                    "numbering_id": level.numbering_id,
                    "abstract_num_id": level.abstract_num_id,
                    "level_text": level.level_text,
                    "number_format": level.number_format,
                    "start_value": level.start_value,
                    "suffix": level.suffix,
                    "justification": level.justification,
                    "style_name": level.style_name,
                    "bwa_label": level.bwa_label,
                    "is_bwa_level": level.is_bwa_level
                }
                for key, level in analysis.bwa_list_levels.items()
            },
            "level_mappings": analysis.level_mappings,
            "summary": analysis.summary
        }
        
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(analysis_dict, f, indent=2, ensure_ascii=False)
        
        print(f"Template analysis saved to: {output_path}")

def main():
    """Main function for standalone usage"""
    import sys
    
    if len(sys.argv) < 2:
        print("Usage: python template_list_detector.py <template_file.docx>")
        print("Example: python template_list_detector.py 'test_template_cleaned.docx'")
        sys.exit(1)
    
    template_file = sys.argv[1]
    
    if not os.path.exists(template_file):
        print(f"Error: File '{template_file}' not found.")
        sys.exit(1)
    
    try:
        detector = TemplateListDetector()
        analysis = detector.analyze_template(template_file)
        
        if analysis:
            # Generate output filename
            base_name = os.path.splitext(os.path.basename(template_file))[0]
            json_file = f"{base_name}_list_analysis.json"
            
            # Save analysis
            detector.save_analysis_to_json(analysis, json_file)
            
            print(f"\nTemplate list analysis completed for {template_file}")
            print(f"Analysis saved to: {json_file}")
            
            # Show summary
            summary = analysis.summary
            print(f"\nSummary:")
            print(f"  Abstract numbering definitions: {summary['total_abstract_numbering']}")
            print(f"  Number mappings: {summary['total_num_mappings']}")
            print(f"  BWA list levels: {summary['total_bwa_levels']}")
            print(f"  Level mappings: {summary['level_mappings_count']}")
            
            if summary.get('level_types'):
                print(f"  Level types: {summary['level_types']}")
            
            # Show BWA levels
            if analysis.bwa_list_levels:
                print(f"\nBWA List Levels:")
                for key, level in analysis.bwa_list_levels.items():
                    if level.is_bwa_level:
                        print(f"  {key}: Level {level.level_number}, Format: {level.number_format}")
                        if level.bwa_label:
                            print(f"    BWA Label: {level.bwa_label}")
        else:
            print("No analysis generated")
            
    except Exception as e:
        print(f"Error processing template: {e}")

if __name__ == "__main__":
    main() 