#!/usr/bin/env python3
"""
Batch Processing Script for Specification Documents

This script processes all specification documents in the examples/Specs folder,
extracting content and regenerating documents with proper formatting.

Usage:
    python batch_process_specs.py
"""

import os
import sys
import subprocess
import time
import shutil
from pathlib import Path

def run_command(command, description):
    """Run a command and handle errors"""
    print(f"\n{'='*60}")
    print(f"Running: {description}")
    print(f"Command: {command}")
    print(f"{'='*60}")
    
    start_time = time.time()
    try:
        result = subprocess.run(command, shell=True, capture_output=True, text=True)
        end_time = time.time()
        
        if result.returncode == 0:
            print(f"✓ SUCCESS: {description}")
            print(f"  Time: {end_time - start_time:.2f} seconds")
            if result.stdout.strip():
                print(f"  Output: {result.stdout.strip()}")
            return True
        else:
            print(f"✗ FAILED: {description}")
            print(f"  Return code: {result.returncode}")
            print(f"  Error: {result.stderr.strip()}")
            return False
            
    except Exception as e:
        print(f"✗ ERROR: {description}")
        print(f"  Exception: {e}")
        return False

def modify_generator_for_document(json_path, output_path):
    """Modify the test_generator.py file for a specific document"""
    try:
        # Read the current test_generator.py
        with open('test_generator.py', 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Replace the paths
        modified_content = content.replace(
            'CONTENT_PATH = \'../output/SECTION 26 05 29_v3.json\'',
            f'CONTENT_PATH = \'{json_path}\''
        )
        modified_content = modified_content.replace(
            'OUTPUT_PATH   = \'../output/generated_spec_v3_fixed_new2.docx\'',
            f'OUTPUT_PATH   = \'{output_path}\''
        )
        
        # Write the modified content back
        with open('test_generator.py', 'w', encoding='utf-8') as f:
            f.write(modified_content)
        
        return True
    except Exception as e:
        print(f"Error modifying generator: {e}")
        return False

def restore_generator():
    """Restore the original test_generator.py configuration"""
    try:
        # Read the current test_generator.py
        with open('test_generator.py', 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Restore original paths
        modified_content = content.replace(
            'CONTENT_PATH = \'../output/SECTION 26 05 29_v3.json\'',
            'CONTENT_PATH = \'../output/SECTION 26 05 29_v3.json\''
        )
        modified_content = modified_content.replace(
            'OUTPUT_PATH   = \'../output/generated_spec_v3_fixed_new2.docx\'',
            'OUTPUT_PATH   = \'../output/generated_spec_v3_fixed_new2.docx\''
        )
        
        # Write the modified content back
        with open('test_generator.py', 'w', encoding='utf-8') as f:
            f.write(modified_content)
        
        return True
    except Exception as e:
        print(f"Error restoring generator: {e}")
        return False

def main():
    """Main batch processing function"""
    print("SpecConverter v0.4 - Batch Processing")
    print("=" * 50)
    
    # Define paths
    examples_dir = Path("../examples/Specs")
    output_dir = Path("../output/Specs")
    template_path = Path("../templates/test_template_cleaned.docx")
    
    # Check if directories exist
    if not examples_dir.exists():
        print(f"Error: Examples directory not found: {examples_dir}")
        return
    
    if not template_path.exists():
        print(f"Error: Template file not found: {template_path}")
        return
    
    # Create output directory
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # Get all DOCX files in the examples directory
    docx_files = list(examples_dir.glob("*.docx"))
    
    if not docx_files:
        print(f"No DOCX files found in {examples_dir}")
        return
    
    print(f"Found {len(docx_files)} documents to process:")
    for docx_file in docx_files:
        print(f"  - {docx_file.name}")
    
    # Process each document
    successful_extractions = 0
    successful_regenerations = 0
    
    try:
        for i, docx_file in enumerate(docx_files, 1):
            print(f"\n{'='*60}")
            print(f"Processing document {i}/{len(docx_files)}: {docx_file.name}")
            print(f"{'='*60}")
            
            # Step 1: Extract content
            extraction_cmd = f'python extract_spec_content_v3.py "{docx_file}" . "{template_path}"'
            if run_command(extraction_cmd, f"Extracting content from {docx_file.name}"):
                successful_extractions += 1
                
                # Step 2: Generate regenerated document
                base_name = docx_file.stem
                json_path = f"../output/{base_name}_v3.json"
                output_path = output_dir / f"{base_name}_regenerated.docx"
                
                # Modify the generator for this document
                if modify_generator_for_document(str(json_path), str(output_path)):
                    # Run the generator
                    if run_command("python test_generator.py", f"Regenerating {docx_file.name}"):
                        successful_regenerations += 1
                else:
                    print(f"Failed to modify generator for {docx_file.name}")
            else:
                print(f"Skipping regeneration for {docx_file.name} due to extraction failure")
    
    finally:
        # Always restore the generator to its original state
        print("\nRestoring generator to original configuration...")
        restore_generator()
    
    # Summary
    print(f"\n{'='*60}")
    print("BATCH PROCESSING COMPLETE")
    print(f"{'='*60}")
    print(f"Total documents: {len(docx_files)}")
    print(f"Successful extractions: {successful_extractions}")
    print(f"Successful regenerations: {successful_regenerations}")
    print(f"Output location: {output_dir}")
    
    if successful_regenerations == len(docx_files):
        print("\n✓ All documents processed successfully!")
    else:
        print(f"\n⚠ {len(docx_files) - successful_regenerations} documents had issues")

if __name__ == "__main__":
    main() 