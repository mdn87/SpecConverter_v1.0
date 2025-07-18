"""
Main CLI entry point for SpecConverter v1.0

Provides the command-line interface for the application.
"""

import argparse
import sys
from pathlib import Path
from typing import Optional

# Add the src directory to the path for imports
sys.path.insert(0, str(Path(__file__).parent.parent))

from utils.logging_utils import setup_logging, get_logger
from utils.file_utils import ensure_directory
from core.extractor import SpecExtractor
from core.generator import SpecGenerator
from core.template_analyzer import TemplateAnalyzer


def create_parser() -> argparse.ArgumentParser:
    """Create the command-line argument parser"""
    parser = argparse.ArgumentParser(
        prog="specconverter",
        description="SpecConverter v1.0 - Specification document conversion and processing toolkit",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Extract content from a single document
  specconverter extract document.docx --template template.docx --output output/

  # Generate a document from JSON
  specconverter generate document.json --template template.docx --output result.docx

  # Analyze a template
  specconverter template analyze template.docx

  # Process a batch job
  specconverter batch process job.yaml
        """
    )
    
    # Global options
    parser.add_argument(
        "--log-level",
        choices=["DEBUG", "INFO", "WARNING", "ERROR"],
        default="INFO",
        help="Set the logging level (default: INFO)"
    )
    
    parser.add_argument(
        "--config",
        type=str,
        help="Path to configuration file"
    )
    
    # Subcommands
    subparsers = parser.add_subparsers(dest="command", help="Available commands")
    
    # Extract command
    extract_parser = subparsers.add_parser("extract", help="Extract content from a document")
    extract_parser.add_argument("document", help="Path to the document to extract")
    extract_parser.add_argument("--template", help="Path to template file")
    extract_parser.add_argument("--output", default="output", help="Output directory")
    extract_parser.add_argument("--save-modular", action="store_true", help="Save modular JSON files")
    
    # Generate command
    generate_parser = subparsers.add_parser("generate", help="Generate document from JSON")
    generate_parser.add_argument("json_file", help="Path to JSON file")
    generate_parser.add_argument("--template", required=True, help="Path to template file")
    generate_parser.add_argument("--output", required=True, help="Output document path")
    
    # Template command
    template_parser = subparsers.add_parser("template", help="Template management")
    template_subparsers = template_parser.add_subparsers(dest="template_command")
    
    analyze_parser = template_subparsers.add_parser("analyze", help="Analyze template structure")
    analyze_parser.add_argument("template", help="Path to template file")
    analyze_parser.add_argument("--output", help="Output file for analysis")
    
    # Batch command
    batch_parser = subparsers.add_parser("batch", help="Batch processing")
    batch_subparsers = batch_parser.add_subparsers(dest="batch_command")
    
    process_parser = batch_subparsers.add_parser("process", help="Process a batch job")
    process_parser.add_argument("job_file", help="Path to batch job configuration file")
    
    validate_parser = batch_subparsers.add_parser("validate", help="Validate a batch job")
    validate_parser.add_argument("job_file", help="Path to batch job configuration file")
    
    return parser


def extract_command(args: argparse.Namespace) -> int:
    """Handle the extract command"""
    logger = get_logger("extract")
    
    try:
        # Ensure output directory exists
        ensure_directory(args.output)
        
        # Initialize extractor
        extractor = SpecExtractor(template_path=args.template)
        
        # Extract content
        logger.info(f"Extracting content from: {args.document}")
        result = extractor.extract_document(args.document)
        
        # Save results
        base_name = Path(args.document).stem
        output_path = Path(args.output) / f"{base_name}_v3.json"
        
        # Save main JSON file
        extractor.save_to_json(result, str(output_path))
        logger.info(f"Saved main JSON to: {output_path}")
        
        # Save modular files if requested
        if args.save_modular:
            extractor.save_modular_json_files(result, base_name, args.output)
            logger.info("Saved modular JSON files")
        
        return 0
        
    except Exception as e:
        logger.error(f"Extraction failed: {e}")
        return 1


def generate_command(args: argparse.Namespace) -> int:
    """Handle the generate command"""
    logger = get_logger("generate")
    
    try:
        # Initialize generator
        generator = SpecGenerator()
        
        # Generate document
        logger.info(f"Generating document from: {args.json_file}")
        result_path = generator.generate_document(args.json_file, args.template, args.output)
        
        logger.info(f"Generated document: {result_path}")
        return 0
        
    except Exception as e:
        logger.error(f"Generation failed: {e}")
        return 1


def template_analyze_command(args: argparse.Namespace) -> int:
    """Handle the template analyze command"""
    logger = get_logger("template")
    
    try:
        # Initialize analyzer
        analyzer = TemplateAnalyzer()
        
        # Analyze template
        logger.info(f"Analyzing template: {args.template}")
        analysis = analyzer.analyze_template(args.template)
        
        # Save or display results
        if args.output:
            analyzer.save_analysis_to_json(analysis, args.output)
            logger.info(f"Analysis saved to: {args.output}")
        else:
            print(f"Template Analysis for: {analysis.template_path}")
            print(f"BWA List Levels: {len(analysis.bwa_list_levels)}")
            print(f"Numbering Definitions: {len(analysis.numbering_definitions)}")
        
        return 0
        
    except Exception as e:
        logger.error(f"Template analysis failed: {e}")
        return 1


def batch_process_command(args: argparse.Namespace) -> int:
    """Handle the batch process command"""
    logger = get_logger("batch")
    
    try:
        # Import here to avoid circular imports
        from batch.processor import BatchProcessor
        
        # Initialize processor
        processor = BatchProcessor()
        
        # Process batch job
        logger.info(f"Processing batch job: {args.job_file}")
        results = processor.process_job(args.job_file)
        
        # Display results
        print(f"Batch processing complete:")
        print(f"  Total processed: {results.total_processed}")
        print(f"  Successful: {results.total_successful}")
        print(f"  Failed: {results.total_failed}")
        
        return 0 if results.total_failed == 0 else 1
        
    except Exception as e:
        logger.error(f"Batch processing failed: {e}")
        return 1


def batch_validate_command(args: argparse.Namespace) -> int:
    """Handle the batch validate command"""
    logger = get_logger("batch")
    
    try:
        # Import here to avoid circular imports
        from batch.processor import BatchProcessor
        
        # Initialize processor
        processor = BatchProcessor()
        
        # Validate batch job
        logger.info(f"Validating batch job: {args.job_file}")
        is_valid = processor.validate_job(args.job_file)
        
        if is_valid:
            print("✓ Batch job configuration is valid")
            return 0
        else:
            print("✗ Batch job configuration has errors")
            return 1
        
    except Exception as e:
        logger.error(f"Batch validation failed: {e}")
        return 1


def main() -> int:
    """Main entry point"""
    parser = create_parser()
    args = parser.parse_args()
    
    # Set up logging
    setup_logging(level=args.log_level)
    logger = get_logger("main")
    
    # Handle commands
    if args.command == "extract":
        return extract_command(args)
    elif args.command == "generate":
        return generate_command(args)
    elif args.command == "template" and args.template_command == "analyze":
        return template_analyze_command(args)
    elif args.command == "batch" and args.batch_command == "process":
        return batch_process_command(args)
    elif args.command == "batch" and args.batch_command == "validate":
        return batch_validate_command(args)
    else:
        parser.print_help()
        return 1


if __name__ == "__main__":
    sys.exit(main()) 