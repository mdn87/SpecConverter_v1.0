# SpecConverter v1.0

A comprehensive specification document conversion and processing toolkit with modular architecture, CLI interface, and batch processing capabilities.

## 🎯 Project Status

**✅ COMPLETED:**
- ✅ Core data models (`src/core/models.py`)
- ✅ CLI interface (`src/cli/main.py`)
- ✅ Configuration system (`config/settings.yaml`, `config/batch_jobs/`)
- ✅ Core extraction logic (`src/core/extractor.py`)
- ✅ Document generation (`src/core/generator.py`)
- ✅ Template analysis (`src/core/template_analyzer.py`)
- ✅ Header/footer extraction (`src/utils/header_footer.py`)
- ✅ File utilities (`src/utils/file_utils.py`)
- ✅ Logging utilities (`src/utils/logging_utils.py`)
- ✅ **NEW: Validation module** (`src/core/validator.py`)
- ✅ **NEW: Batch reporter** (`src/batch/reporter.py`)
- ✅ **NEW: Refactored batch processor** (`src/batch/processor.py`)
- ✅ **NEW: Unit tests** (`tests/test_models.py`, `tests/test_validator.py`)
- ✅ **NEW: PDF extraction module** (`src/core/pdf_extractor.py`)
- ✅ **NEW: Hybrid analysis module** (`src/core/hybrid_analyzer.py`)
- ✅ Package setup (`setup.py`, `requirements.txt`)

**🚧 IN PROGRESS:**
- 🔄 Integration testing with real documents
- 🔄 Performance optimization
- 🔄 Documentation updates

**📋 NEXT STEPS:**
- [ ] Test CLI commands with real documents
- [ ] Validate batch processing with example files
- [ ] Add integration tests
- [ ] Performance benchmarking
- [ ] User documentation
- [ ] Migration guide from v0.4

## 🚀 Quick Start

### Installation

```bash
# Install in development mode
pip install -e .

# Or install dependencies only
pip install -r requirements.txt
```

### Basic Usage

```bash
# Extract content from a document
specconverter extract document.docx --template template.docx --output output/

# Generate a document from JSON
specconverter generate document.json --template template.docx --output result.docx

# Analyze a template
specconverter template analyze template.docx

# Process a batch job
specconverter batch process config/batch_jobs/example.yaml

# Validate a batch job configuration
specconverter batch validate config/batch_jobs/example.yaml

# NEW: Extract content using PDF conversion and OCR
specconverter pdf-extract document.docx --output output/

# NEW: Hybrid analysis with PDF validation
specconverter hybrid document.docx --template template.docx --output output/ --validation-report
```

## 📁 Project Structure

```
SpecConverter_v1.0/
├── src/
│   ├── core/                    # Core processing modules
│   │   ├── models.py           # Data models and structures
│   │   ├── extractor.py        # Content extraction logic
│   │   ├── generator.py        # Document generation
│   │   ├── template_analyzer.py # Template analysis
│   │   ├── validator.py        # Validation and correction
│   │   ├── pdf_extractor.py    # NEW: PDF-based extraction
│   │   └── hybrid_analyzer.py  # NEW: Hybrid analysis
│   ├── utils/                   # Utility modules
│   │   ├── file_utils.py       # File operations
│   │   ├── logging_utils.py    # Logging setup
│   │   └── header_footer.py    # Header/footer extraction
│   ├── cli/                     # Command-line interface
│   │   └── main.py             # CLI entry point
│   └── batch/                   # Batch processing
│       ├── processor.py        # Batch orchestration
│       └── reporter.py         # Batch reporting
├── config/                      # Configuration files
│   ├── settings.yaml           # Global settings
│   ├── batch_jobs/             # Batch job configs
│   └── templates/              # Template files
├── tests/                       # Unit tests
│   ├── test_models.py          # Data model tests
│   └── test_validator.py       # Validation tests
├── examples/                    # Example documents
├── output/                      # Generated output
├── docs/                        # Documentation
├── requirements.txt
├── setup.py
└── README.md
```

## 🔧 Key Features

### Core Functionality
- **Content Extraction**: Extract structured content from Word documents
- **Document Generation**: Regenerate documents with proper formatting
- **Template Analysis**: Analyze document templates and numbering schemes
- **Validation**: Validate extracted content and auto-correct issues
- **Header/Footer Handling**: Extract and preserve document formatting

### NEW: PDF-Based Extraction
- **PDF Conversion**: Convert Word documents to PDF for better text extraction
- **OCR Integration**: Use Tesseract OCR for complex document layouts
- **Multiple Methods**: Fallback to direct text extraction if PDF conversion fails
- **Enhanced Pattern Matching**: Support for various numbering formats (26-05-29, 26.05.00, etc.)

### NEW: Hybrid Analysis
- **Cross-Reference Validation**: Compare source document with PDF extraction
- **Template Integration**: Use template patterns to validate numbering
- **Confidence Scoring**: Only apply corrections with high confidence (>70%)
- **Line-Based Context**: Analyze line breaks and context for better numbering detection
- **Conservative Approach**: Avoid false corrections by being selective

### Batch Processing
- **Parallel Processing**: Process multiple documents simultaneously
- **Job Configuration**: YAML-based batch job configuration
- **Progress Reporting**: Comprehensive reporting and error tracking
- **Flexible Input**: Support for directory patterns and file lists

### CLI Interface
- **Extract Command**: Extract content from single documents
- **Generate Command**: Generate documents from JSON data
- **Template Command**: Analyze and manage templates
- **Batch Command**: Process multiple documents
- **Validation**: Validate configurations and content
- **NEW: PDF Extract Command**: Extract using PDF conversion and OCR
- **NEW: Hybrid Command**: Cross-reference analysis with PDF validation

### Configuration System
- **Global Settings**: Centralized configuration management
- **Batch Jobs**: Reusable job configurations
- **Template Management**: Organized template storage
- **Logging**: Configurable logging levels and outputs

## 🧪 Testing

Run the unit tests:

```bash
# Run all tests
python -m pytest tests/

# Run specific test file
python -m pytest tests/test_models.py

# Run with verbose output
python -m pytest tests/ -v
```

## 📊 Validation Features

The new validation module provides:

- **Content Validation**: Check for empty blocks, missing numbers, etc.
- **Structure Validation**: Validate document hierarchy and organization
- **Auto-Correction**: Automatically fix common issues
- **Custom Rules**: Load custom validation rules from JSON
- **Error Reporting**: Detailed error reports with context

## 🔍 NEW: Hybrid Analysis Features

The hybrid analyzer combines multiple extraction methods for maximum accuracy:

### PDF Extraction Capabilities
- **Word to PDF Conversion**: Using pandoc, LibreOffice, or direct text extraction
- **OCR Text Extraction**: Tesseract-based OCR for complex layouts
- **Pattern Recognition**: Enhanced numbering pattern matching
- **Hierarchical Analysis**: Proper level detection for different formats

### Cross-Reference Validation
- **Source vs PDF Comparison**: Compare extraction results from different methods
- **Template Pattern Matching**: Validate against known template patterns
- **Confidence Scoring**: Calculate confidence for each potential correction
- **Conservative Application**: Only apply corrections with high confidence

### Line-Based Context Analysis
- **Line Break Preservation**: Maintain document structure information
- **Context Window**: Analyze current and previous lines for numbering
- **Word Boundary Validation**: Ensure numbering is complete and not fragmented
- **Priority-Based Matching**: Prioritize higher-level numbering patterns

## 📈 Batch Reporting

The batch reporter generates:

- **JSON Reports**: Machine-readable processing results
- **CSV Reports**: Spreadsheet-friendly data export
- **Summary Reports**: Human-readable processing summaries
- **Performance Metrics**: Processing time and throughput analysis
- **Error Analysis**: Detailed error categorization and statistics

## 🔄 Migration from v0.4

The v1.0 architecture preserves all the valuable logic from v0.4 while providing:

- **Modular Design**: Clean separation of concerns
- **Type Safety**: Comprehensive type hints throughout
- **Error Handling**: Robust error handling and recovery
- **Configuration**: Flexible configuration management
- **Testing**: Comprehensive unit test coverage
- **Documentation**: Clear documentation and examples
- **NEW: Hybrid Analysis**: Advanced validation using multiple extraction methods

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests for new functionality
5. Run the test suite
6. Submit a pull request

## 📝 License

This project is licensed under the MIT License - see the LICENSE file for details.

---

**SpecConverter v1.0** - Ready for robust, maintainable, and user-friendly specification document processing! 🎉

### NEW: Hybrid Analysis Approach

The hybrid analysis represents a significant advancement in document processing accuracy:

1. **Multi-Method Extraction**: Combines traditional Word parsing with PDF conversion and OCR
2. **Cross-Reference Validation**: Compares results from different extraction methods
3. **Template Integration**: Uses template patterns to validate numbering schemes
4. **Confidence-Based Corrections**: Only applies changes when highly confident
5. **Line-Aware Processing**: Preserves document structure and line breaks
6. **Conservative Approach**: Prioritizes accuracy over quantity of corrections

This approach ensures maximum accuracy while minimizing false corrections, making it ideal for production use in specification document processing.
