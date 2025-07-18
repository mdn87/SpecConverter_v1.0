"""
File utilities for SpecConverter v1.0

Common file operations and utilities used throughout the application.
"""

import os
import json
import shutil
from pathlib import Path
from typing import Dict, List, Any, Optional


def ensure_directory(path: str) -> None:
    """Ensure a directory exists, creating it if necessary"""
    Path(path).mkdir(parents=True, exist_ok=True)


def save_json(data: Dict[str, Any], file_path: str, indent: int = 2) -> None:
    """Save data to a JSON file"""
    ensure_directory(os.path.dirname(file_path))
    with open(file_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=indent, ensure_ascii=False)


def load_json(file_path: str) -> Optional[Dict[str, Any]]:
    """Load data from a JSON file"""
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return None


def get_file_extension(file_path: str) -> str:
    """Get the file extension from a file path"""
    return Path(file_path).suffix.lower()


def is_docx_file(file_path: str) -> bool:
    """Check if a file is a DOCX file"""
    return get_file_extension(file_path) == '.docx'


def get_base_name(file_path: str) -> str:
    """Get the base name (without extension) from a file path"""
    return Path(file_path).stem


def find_files_by_pattern(directory: str, pattern: str) -> List[str]:
    """Find files matching a pattern in a directory"""
    directory_path = Path(directory)
    if not directory_path.exists():
        return []
    
    return [str(f) for f in directory_path.glob(pattern)]


def copy_file_with_backup(source: str, destination: str) -> bool:
    """Copy a file with backup if destination exists"""
    try:
        dest_path = Path(destination)
        if dest_path.exists():
            backup_path = dest_path.with_suffix(dest_path.suffix + '.backup')
            shutil.copy2(dest_path, backup_path)
        
        shutil.copy2(source, destination)
        return True
    except Exception:
        return False


def get_file_size(file_path: str) -> int:
    """Get the size of a file in bytes"""
    try:
        return os.path.getsize(file_path)
    except OSError:
        return 0


def format_file_size(size_bytes: int) -> str:
    """Format file size in human-readable format"""
    size = float(size_bytes)
    for unit in ['B', 'KB', 'MB', 'GB']:
        if size < 1024.0:
            return f"{size:.1f} {unit}"
        size /= 1024.0
    return f"{size:.1f} TB" 