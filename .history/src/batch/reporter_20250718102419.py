"""
Batch reporting and analysis for SpecConverter v1.0

Provides comprehensive reporting capabilities for batch processing operations.
"""

import json
import csv
from datetime import datetime, timedelta
from pathlib import Path
from typing import Dict, List, Any, Optional
import statistics

from .models import BatchResults, BatchJob
from utils.logging_utils import get_logger


class BatchReporter:
    """Handles reporting and analysis of batch processing results"""
    
    def __init__(self, output_dir: str = "output/reports"):
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.logger = get_logger("batch_reporter")
    
    def generate_report(self, results: BatchResults, job: BatchJob) -> Dict[str, Any]:
        """Generate a comprehensive batch processing report"""
        self.logger.info(f"Generating batch report for job: {job.name}")
        
        # Calculate processing statistics
        processing_time = results.end_time - results.start_time
        success_rate = (results.total_successful / results.total_processed * 100) if results.total_processed > 0 else 0
        
        # Create report data
        report_data = {
            "job_info": {
                "name": job.name,
                "description": job.description,
                "template_path": job.template_path,
                "output_directory": job.output_dir,
                "options": job.options
            },
            "processing_summary": {
                "start_time": results.start_time.isoformat(),
                "end_time": results.end_time.isoformat(),
                "processing_duration": str(processing_time),
                "processing_duration_seconds": processing_time.total_seconds(),
                "total_processed": results.total_processed,
                "total_successful": results.total_successful,
                "total_failed": results.total_failed,
                "success_rate_percent": round(success_rate, 2)
            },
            "file_results": {
                "successful_files": results.successful,
                "failed_files": results.failed,
                "error_details": results.errors
            },
            "performance_metrics": {
                "average_time_per_file": processing_time.total_seconds() / results.total_processed if results.total_processed > 0 else 0,
                "files_per_minute": (results.total_processed / processing_time.total_seconds() * 60) if processing_time.total_seconds() > 0 else 0
            }
        }
        
        return report_data
    
    def save_json_report(self, report_data: Dict[str, Any], filename: Optional[str] = None) -> str:
        """Save report as JSON file"""
        if not filename:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"batch_report_{timestamp}.json"
        
        output_path = self.output_dir / filename
        
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(report_data, f, indent=2, ensure_ascii=False)
        
        self.logger.info(f"JSON report saved to: {output_path}")
        return str(output_path)
    
    def save_csv_report(self, report_data: Dict[str, Any], filename: Optional[str] = None) -> str:
        """Save report as CSV file"""
        if not filename:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"batch_report_{timestamp}.csv"
        
        output_path = self.output_dir / filename
        
        with open(output_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            
            # Write header
            writer.writerow([
                "File Path", "Status", "Processing Time", "Error Message"
            ])
            
            # Write successful files
            for file_path in report_data["file_results"]["successful_files"]:
                writer.writerow([file_path, "SUCCESS", "", ""])
            
            # Write failed files
            for i, file_path in enumerate(report_data["file_results"]["failed_files"]):
                error_msg = report_data["file_results"]["error_details"][i] if i < len(report_data["file_results"]["error_details"]) else "Unknown error"
                writer.writerow([file_path, "FAILED", "", error_msg])
        
        self.logger.info(f"CSV report saved to: {output_path}")
        return str(output_path)
    
    def generate_summary_report(self, report_data: Dict[str, Any]) -> str:
        """Generate a human-readable summary report"""
        summary = []
        summary.append("=" * 80)
        summary.append("BATCH PROCESSING REPORT")
        summary.append("=" * 80)
        summary.append("")
        
        # Job information
        summary.append("JOB INFORMATION:")
        summary.append(f"  Name: {report_data['job_info']['name']}")
        summary.append(f"  Description: {report_data['job_info']['description']}")
        summary.append(f"  Template: {report_data['job_info']['template_path']}")
        summary.append(f"  Output Directory: {report_data['job_info']['output_directory']}")
        summary.append("")
        
        # Processing summary
        summary.append("PROCESSING SUMMARY:")
        summary.append(f"  Start Time: {report_data['processing_summary']['start_time']}")
        summary.append(f"  End Time: {report_data['processing_summary']['end_time']}")
        summary.append(f"  Duration: {report_data['processing_summary']['processing_duration']}")
        summary.append(f"  Total Files: {report_data['processing_summary']['total_processed']}")
        summary.append(f"  Successful: {report_data['processing_summary']['total_successful']}")
        summary.append(f"  Failed: {report_data['processing_summary']['total_failed']}")
        summary.append(f"  Success Rate: {report_data['processing_summary']['success_rate_percent']}%")
        summary.append("")
        
        # Performance metrics
        summary.append("PERFORMANCE METRICS:")
        summary.append(f"  Average Time per File: {report_data['performance_metrics']['average_time_per_file']:.2f} seconds")
        summary.append(f"  Files per Minute: {report_data['performance_metrics']['files_per_minute']:.2f}")
        summary.append("")
        
        # Error summary
        if report_data['processing_summary']['total_failed'] > 0:
            summary.append("ERROR SUMMARY:")
            error_counts = {}
            for error in report_data['file_results']['error_details']:
                error_type = error.split(':')[0] if ':' in error else error
                error_counts[error_type] = error_counts.get(error_type, 0) + 1
            
            for error_type, count in error_counts.items():
                summary.append(f"  {error_type}: {count} occurrences")
            summary.append("")
        
        # File list
        if report_data['file_results']['successful_files']:
            summary.append("SUCCESSFUL FILES:")
            for file_path in report_data['file_results']['successful_files']:
                summary.append(f"  ✓ {file_path}")
            summary.append("")
        
        if report_data['file_results']['failed_files']:
            summary.append("FAILED FILES:")
            for i, file_path in enumerate(report_data['file_results']['failed_files']):
                error_msg = report_data['file_results']['error_details'][i] if i < len(report_data['file_results']['error_details']) else "Unknown error"
                summary.append(f"  ✗ {file_path}")
                summary.append(f"    Error: {error_msg}")
            summary.append("")
        
        summary.append("=" * 80)
        
        return "\n".join(summary)
    
    def save_summary_report(self, report_data: Dict[str, Any], filename: Optional[str] = None) -> str:
        """Save human-readable summary report"""
        if not filename:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"batch_summary_{timestamp}.txt"
        
        output_path = self.output_dir / filename
        summary_text = self.generate_summary_report(report_data)
        
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(summary_text)
        
        self.logger.info(f"Summary report saved to: {output_path}")
        return str(output_path)
    
    def generate_comparison_report(self, reports: List[Dict[str, Any]]) -> Dict[str, Any]:
        """Generate a comparison report for multiple batch jobs"""
        if len(reports) < 2:
            raise ValueError("At least 2 reports required for comparison")
        
        comparison_data = {
            "comparison_summary": {
                "total_jobs": len(reports),
                "comparison_date": datetime.now().isoformat()
            },
            "job_comparisons": []
        }
        
        # Compare each job
        for i, report in enumerate(reports):
            job_comparison = {
                "job_name": report["job_info"]["name"],
                "total_processed": report["processing_summary"]["total_processed"],
                "total_successful": report["processing_summary"]["total_successful"],
                "total_failed": report["processing_summary"]["total_failed"],
                "success_rate": report["processing_summary"]["success_rate_percent"],
                "processing_duration": report["processing_summary"]["processing_duration_seconds"],
                "files_per_minute": report["performance_metrics"]["files_per_minute"]
            }
            comparison_data["job_comparisons"].append(job_comparison)
        
        # Calculate statistics
        success_rates = [job["success_rate"] for job in comparison_data["job_comparisons"]]
        processing_times = [job["processing_duration"] for job in comparison_data["job_comparisons"]]
        files_per_minute = [job["files_per_minute"] for job in comparison_data["job_comparisons"]]
        
        comparison_data["statistics"] = {
            "success_rate": {
                "average": statistics.mean(success_rates),
                "min": min(success_rates),
                "max": max(success_rates),
                "std_dev": statistics.stdev(success_rates) if len(success_rates) > 1 else 0
            },
            "processing_time": {
                "average": statistics.mean(processing_times),
                "min": min(processing_times),
                "max": max(processing_times),
                "std_dev": statistics.stdev(processing_times) if len(processing_times) > 1 else 0
            },
            "files_per_minute": {
                "average": statistics.mean(files_per_minute),
                "min": min(files_per_minute),
                "max": max(files_per_minute),
                "std_dev": statistics.stdev(files_per_minute) if len(files_per_minute) > 1 else 0
            }
        }
        
        return comparison_data
    
    def print_summary(self, report_data: Dict[str, Any]):
        """Print a summary to console"""
        summary_text = self.generate_summary_report(report_data)
        print(summary_text)
    
    def create_dashboard_data(self, report_data: Dict[str, Any]) -> Dict[str, Any]:
        """Create data suitable for dashboard visualization"""
        dashboard_data = {
            "job_name": report_data["job_info"]["name"],
            "processing_stats": {
                "total": report_data["processing_summary"]["total_processed"],
                "successful": report_data["processing_summary"]["total_successful"],
                "failed": report_data["processing_summary"]["total_failed"],
                "success_rate": report_data["processing_summary"]["success_rate_percent"]
            },
            "performance": {
                "duration_seconds": report_data["processing_summary"]["processing_duration_seconds"],
                "files_per_minute": report_data["performance_metrics"]["files_per_minute"],
                "average_time_per_file": report_data["performance_metrics"]["average_time_per_file"]
            },
            "timeline": {
                "start": report_data["processing_summary"]["start_time"],
                "end": report_data["processing_summary"]["end_time"]
            },
            "error_breakdown": {}
        }
        
        # Count error types
        if report_data["file_results"]["error_details"]:
            error_counts = {}
            for error in report_data["file_results"]["error_details"]:
                error_type = error.split(':')[0] if ':' in error else "Unknown"
                error_counts[error_type] = error_counts.get(error_type, 0) + 1
            dashboard_data["error_breakdown"] = error_counts
        
        return dashboard_data 