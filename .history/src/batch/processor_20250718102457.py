"""
Batch processing for SpecConverter v1.0

Provides batch processing capabilities for multiple specification documents.
"""

import yaml
import os
import sys
from datetime import datetime
from pathlib import Path
from typing import List, Dict, Any, Optional
from concurrent.futures import ThreadPoolExecutor, as_completed
import time

# Add the src directory to the path for imports
sys.path.insert(0, str(Path(__file__).parent.parent))

from core.models import BatchJob, BatchResults
from core.extractor import SpecExtractor
from core.generator import SpecGenerator
from core.validator import SpecValidator
from utils.logging_utils import get_logger
from utils.file_utils import ensure_directory
from .reporter import BatchReporter


class BatchProcessor:
    """Handles batch processing of specification documents"""
    
    def __init__(self, config: Optional[Dict[str, Any]] = None):
        self.config = config or {}
        self.logger = get_logger("batch_processor")
        self.reporter = BatchReporter()
    
    def load_job_config(self, job_file: str) -> BatchJob:
        """Load batch job configuration from YAML file"""
        self.logger.info(f"Loading batch job configuration: {job_file}")
        
        try:
            with open(job_file, 'r', encoding='utf-8') as f:
                job_data = yaml.safe_load(f)
            
            # Validate required fields
            required_fields = ['name', 'input', 'template', 'output_directory']
            for field in required_fields:
                if field not in job_data:
                    raise ValueError(f"Missing required field: {field}")
            
            # Build input paths
            input_paths = self._build_input_paths(job_data['input'])
            
            # Create BatchJob object
            job = BatchJob(
                name=job_data['name'],
                description=job_data.get('description', ''),
                input_paths=input_paths,
                template_path=job_data['template'],
                output_dir=job_data['output_directory'],
                options=job_data.get('options', {})
            )
            
            self.logger.info(f"Loaded job: {job.name} with {len(job.input_paths)} input files")
            return job
            
        except Exception as e:
            self.logger.error(f"Failed to load job configuration: {e}")
            raise
    
    def _build_input_paths(self, input_config: Dict[str, Any]) -> List[str]:
        """Build list of input file paths from configuration"""
        input_paths = []
        
        # Handle directory-based input
        if 'directory' in input_config:
            directory = Path(input_config['directory'])
            pattern = input_config.get('pattern', '*.docx')
            exclude_patterns = input_config.get('exclude', [])
            
            if not directory.exists():
                raise ValueError(f"Input directory does not exist: {directory}")
            
            # Find all matching files
            for file_path in directory.glob(pattern):
                # Check exclusion patterns
                should_exclude = False
                for exclude_pattern in exclude_patterns:
                    if file_path.match(exclude_pattern):
                        should_exclude = True
                        break
                
                if not should_exclude:
                    input_paths.append(str(file_path))
        
        # Handle explicit file list
        elif 'files' in input_config:
            for file_path in input_config['files']:
                if Path(file_path).exists():
                    input_paths.append(file_path)
                else:
                    self.logger.warning(f"Input file not found: {file_path}")
        
        return input_paths
    
    def process_job(self, job_file: str) -> BatchResults:
        """Process a batch job from configuration file"""
        # Load job configuration
        job = self.load_job_config(job_file)
        
        # Validate job configuration
        self._validate_job(job)
        
        # Start processing
        start_time = datetime.now()
        self.logger.info(f"Starting batch processing: {job.name}")
        
        successful_files = []
        failed_files = []
        error_details = []
        
        # Determine processing mode
        extract_only = job.options.get('extract_only', False)
        validate_only = job.options.get('validate_only', False)
        parallel_processing = job.options.get('parallel_processing', True)
        max_workers = job.options.get('max_workers', 4)
        
        if parallel_processing and len(job.input_paths) > 1:
            # Parallel processing
            results = self._process_parallel(job, extract_only, validate_only, max_workers)
            successful_files = results['successful']
            failed_files = results['failed']
            error_details = results['errors']
        else:
            # Sequential processing
            for file_path in job.input_paths:
                try:
                    self._process_single_file(job, file_path, extract_only, validate_only)
                    successful_files.append(file_path)
                except Exception as e:
                    failed_files.append(file_path)
                    error_details.append(str(e))
                    self.logger.error(f"Failed to process {file_path}: {e}")
        
        end_time = datetime.now()
        
        # Create batch results
        results = BatchResults(
            job=job,
            successful=successful_files,
            failed=failed_files,
            errors=error_details,
            start_time=start_time,
            end_time=end_time,
            total_processed=len(job.input_paths),
            total_successful=len(successful_files),
            total_failed=len(failed_files)
        )
        
        # Generate and save reports
        self._generate_reports(results)
        
        self.logger.info(f"Batch processing complete: {results.total_successful}/{results.total_processed} successful")
        return results
    
    def _process_parallel(self, job: BatchJob, extract_only: bool, validate_only: bool, max_workers: int) -> Dict[str, List[str]]:
        """Process files in parallel"""
        successful_files = []
        failed_files = []
        error_details = []
        
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            # Submit all tasks
            future_to_file = {
                executor.submit(self._process_single_file, job, file_path, extract_only, validate_only): file_path
                for file_path in job.input_paths
            }
            
            # Collect results
            for future in as_completed(future_to_file):
                file_path = future_to_file[future]
                try:
                    future.result()  # This will raise any exception that occurred
                    successful_files.append(file_path)
                except Exception as e:
                    failed_files.append(file_path)
                    error_details.append(str(e))
                    self.logger.error(f"Failed to process {file_path}: {e}")
        
        return {
            'successful': successful_files,
            'failed': failed_files,
            'errors': error_details
        }
    
    def _process_single_file(self, job: BatchJob, file_path: str, extract_only: bool, validate_only: bool):
        """Process a single file"""
        self.logger.info(f"Processing file: {file_path}")
        
        # Ensure output directory exists
        ensure_directory(job.output_dir)
        
        # Initialize components
        extractor = SpecExtractor(template_path=job.template_path)
        generator = SpecGenerator()
        validator = SpecValidator(config=job.options.get('validation', {}))
        
        # Extract content
        document = extractor.extract_document(file_path)
        
        # Validate if requested
        if not extract_only:
            validation_results = validator.validate_document(document)
            document.validation_results = validation_results
            
            # Save validation report if requested
            if job.options.get('save_error_reports', True):
                base_name = Path(file_path).stem
                validation_report_path = Path(job.output_dir) / f"{base_name}_validation.json"
                validator.save_validation_report(validation_results, str(validation_report_path))
        
        # Save extracted content
        base_name = Path(file_path).stem
        json_output_path = Path(job.output_dir) / f"{base_name}_v3.json"
        extractor.save_to_json(document, str(json_output_path))
        
        # Generate document if not extract-only
        if not extract_only and not validate_only:
            output_doc_path = Path(job.output_dir) / f"{base_name}_regenerated.docx"
            generator.generate_document(str(json_output_path), job.template_path, str(output_doc_path))
        
        self.logger.info(f"Completed processing: {file_path}")
    
    def _validate_job(self, job: BatchJob):
        """Validate batch job configuration"""
        # Check template file exists
        if not Path(job.template_path).exists():
            raise ValueError(f"Template file not found: {job.template_path}")
        
        # Check input files exist
        for file_path in job.input_paths:
            if not Path(file_path).exists():
                raise ValueError(f"Input file not found: {file_path}")
        
        # Check output directory is writable
        try:
            ensure_directory(job.output_dir)
        except Exception as e:
            raise ValueError(f"Output directory not writable: {job.output_dir} - {e}")
    
    def _generate_reports(self, results: BatchResults):
        """Generate and save batch processing reports"""
        # Generate report data
        report_data = self.reporter.generate_report(results, results.job)
        
        # Save reports
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        job_name_safe = results.job.name.replace(" ", "_").replace("/", "_")
        
        # JSON report
        json_report_path = self.reporter.save_json_report(
            report_data, 
            f"batch_report_{job_name_safe}_{timestamp}.json"
        )
        
        # CSV report
        csv_report_path = self.reporter.save_csv_report(
            report_data,
            f"batch_report_{job_name_safe}_{timestamp}.csv"
        )
        
        # Summary report
        summary_report_path = self.reporter.save_summary_report(
            report_data,
            f"batch_summary_{job_name_safe}_{timestamp}.txt"
        )
        
        # Print summary to console
        self.reporter.print_summary(report_data)
        
        self.logger.info(f"Reports saved:")
        self.logger.info(f"  JSON: {json_report_path}")
        self.logger.info(f"  CSV: {csv_report_path}")
        self.logger.info(f"  Summary: {summary_report_path}")
    
    def validate_job_config(self, job_file: str) -> bool:
        """Validate a batch job configuration file"""
        try:
            job = self.load_job_config(job_file)
            self._validate_job(job)
            
            print(f"✓ Job configuration is valid: {job.name}")
            print(f"  Input files: {len(job.input_paths)}")
            print(f"  Template: {job.template_path}")
            print(f"  Output directory: {job.output_dir}")
            
            return True
            
        except Exception as e:
            print(f"✗ Job configuration is invalid: {e}")
            return False 