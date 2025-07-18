"""
Logging utilities for SpecConverter v1.0

Provides consistent logging setup and utilities across the application.
"""

import logging
import sys
from pathlib import Path
from typing import Optional


def setup_logging(
    level: str = "INFO",
    log_file: Optional[str] = None,
    format_string: Optional[str] = None
) -> logging.Logger:
    """
    Set up logging configuration
    
    Args:
        level: Logging level (DEBUG, INFO, WARNING, ERROR, CRITICAL)
        log_file: Optional log file path
        format_string: Optional custom format string
        
    Returns:
        Configured logger instance
    """
    # Default format
    if format_string is None:
        format_string = "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
    
    # Create logger
    logger = logging.getLogger("SpecConverter")
    logger.setLevel(getattr(logging, level.upper()))
    
    # Clear existing handlers
    logger.handlers.clear()
    
    # Create formatter
    formatter = logging.Formatter(format_string)
    
    # Console handler
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)
    
    # File handler (if specified)
    if log_file:
        # Ensure log directory exists
        log_path = Path(log_file)
        log_path.parent.mkdir(parents=True, exist_ok=True)
        
        file_handler = logging.FileHandler(log_file, encoding='utf-8')
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)
    
    return logger


def get_logger(name: str = "SpecConverter") -> logging.Logger:
    """Get a logger instance with the specified name"""
    return logging.getLogger(name)


class ProgressLogger:
    """Utility class for logging progress of long-running operations"""
    
    def __init__(self, logger: logging.Logger, total: int, description: str = "Processing"):
        self.logger = logger
        self.total = total
        self.current = 0
        self.description = description
        self.last_percentage = 0
    
    def update(self, increment: int = 1) -> None:
        """Update progress by the specified increment"""
        self.current += increment
        percentage = int((self.current / self.total) * 100)
        
        # Log every 10% or when complete
        if percentage >= self.last_percentage + 10 or self.current >= self.total:
            self.logger.info(f"{self.description}: {self.current}/{self.total} ({percentage}%)")
            self.last_percentage = percentage
    
    def complete(self) -> None:
        """Mark the operation as complete"""
        self.current = self.total
        self.logger.info(f"{self.description}: Complete ({self.total}/{self.total})")


def log_error_with_context(logger: logging.Logger, error: Exception, context: str = "") -> None:
    """Log an error with additional context"""
    error_msg = f"Error in {context}: {str(error)}" if context else f"Error: {str(error)}"
    logger.error(error_msg, exc_info=True)


def log_warning_with_context(logger: logging.Logger, message: str, context: str = "") -> None:
    """Log a warning with additional context"""
    warning_msg = f"Warning in {context}: {message}" if context else f"Warning: {message}"
    logger.warning(warning_msg) 