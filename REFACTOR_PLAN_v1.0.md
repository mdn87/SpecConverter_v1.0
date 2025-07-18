# SpecConverter v1.0 - Refactor Plan

## Overview

This document outlines the comprehensive refactor plan for SpecConverter v1.0, transforming the current script-based approach into a unified, batch-capable tool with proper architecture, CLI interface, and configuration management.

## 1. Project Structure & Git Strategy

### New Repository Structure

```
SpecConverter_v1.0/
├── src/
│   ├── core/                    # Core modules
│   │   ├── __init__.py
│   │   ├── extractor.py         # Main extraction logic
│   │   ├── validator.py         # Validation & correction logic
│   │   ├── generator.py         # Document regeneration
│   │   ├── template_analyzer.py # Template analysis
│   │   └── models.py            # Shared data models
│   ├── utils/                   # Utility modules
│   │   ├── __init__.py
│   │   ├── header_footer.py     # Header/footer extraction
│   │   ├── file_utils.py        # File operations
│   │   └── logging_utils.py     # Logging setup
│   ├── cli/                     # Command-line interface
│   │   ├── __init__.py
│   │   ├── main.py              # Main CLI entry point
│   │   └── commands.py          # CLI commands
│   └── batch/                   # Batch processing
│       ├── __init__.py
│       ├── processor.py         # Batch orchestration
│       └── reporter.py          # Batch reporting
├── config/                      # Configuration files
│   ├── templates/               # Template files
│   ├── batch_jobs/              # Batch job configs
│   └── settings.yaml            # Global settings
├── tests/                       # Unit tests
├── docs/                        # Documentation
├── examples/                    # Example documents
├── output/                      # Generated output
├── requirements.txt
├── setup.py                     # Package setup
├── README.md
└── .gitignore
```

### Git Strategy

- **Fresh repository** for v1.0 (clean slate)
- **Preserve history** by copying key files from v0.4
- **Modular commits** as we build each component
- **Feature branches** for major components
- **Tags** for releases (v1.0.0, v1.0.1, etc.)

## 2. Core Architecture

### Main Classes

```python
# models.py
@dataclass
class SpecDocument:
    file_path: str
    content_blocks: List[ContentBlock]
    header_footer: HeaderFooterData
    template_analysis: TemplateAnalysis
    validation_results: ValidationResults

@dataclass
class BatchJob:
    name: str
    input_paths: List[str]
    template_path: str
    output_dir: str
    options: Dict[str, Any]

# core/extractor.py
class SpecExtractor:
    def extract_document(self, docx_path: str, template_path: str) -> SpecDocument
  
# core/validator.py  
class SpecValidator:
    def validate_and_correct(self, document: SpecDocument) -> ValidationResults

# core/generator.py
class SpecGenerator:
    def generate_document(self, document: SpecDocument, template_path: str) -> str

# batch/processor.py
class BatchProcessor:
    def process_job(self, job: BatchJob) -> BatchResults
```

## 3. CLI Design

### Main Commands

```bash
# Single document processing
specconverter extract document.docx --template template.docx --output output/
specconverter generate document.json --template template.docx --output result.docx

# Batch processing
specconverter batch process job.yaml
specconverter batch validate job.yaml

# Template management
specconverter template analyze template.docx
specconverter template list

# Utility commands
specconverter validate document.json
specconverter report job.yaml
```

## 4. Configuration System

### Global Settings (config/settings.yaml)

```yaml
# Default paths
default_template: "config/templates/default.docx"
output_directory: "output"
log_level: "INFO"

# Processing options
validation:
  max_iterations: 10
  auto_correct: true
  
# Batch processing
batch:
  max_workers: 4
  timeout_minutes: 30
```

### Batch Job Config (config/batch_jobs/example.yaml)

```yaml
name: "Fire Suppression Specs"
description: "Process all fire suppression specification documents"

input:
  directory: "examples/fire_suppression/"
  pattern: "*.docx"
  exclude: ["draft_*", "*_old.*"]

template: "config/templates/RPA_template.docx"
output_directory: "output/fire_suppression/"

options:
  extract_only: false
  validate_only: false
  skip_existing: true
  parallel_processing: true
```

## 5. Migration Strategy

### Phase 1: Core Refactor

1. Create new v1.0 repository
2. Extract core logic from v0.4 scripts into modules
3. Create data models and interfaces
4. Basic CLI structure

### Phase 2: Batch Processing

1. Build batch processor
2. Add configuration system
3. Implement job management
4. Add progress reporting

### Phase 3: Testing & Polish

1. Unit tests for core modules
2. Integration tests with real documents
3. Documentation
4. Performance optimization

### Phase 4: Migration

1. Test with existing v0.4 documents
2. Compare outputs for consistency
3. Gradual migration of workflows

## 6. Key Improvements for v1.0

### Robustness

- Better error handling and recovery
- Progress tracking for long operations
- Detailed logging and debugging
- Validation at each step

### Flexibility

- Configurable processing pipelines
- Plugin architecture for custom validators
- Template-agnostic processing
- Multiple output formats

### Usability

- Clear CLI with help and examples
- Configuration files for common scenarios
- Batch job management
- Progress reporting and notifications

### Maintainability

- Clean separation of concerns
- Comprehensive testing
- Type hints and documentation
- Modular design for easy extension

## 7. Implementation Priorities

### High Priority

1. **Core Data Models** - Define all shared data structures
2. **Extraction Module** - Refactor current extraction logic
3. **Basic CLI** - Simple command-line interface
4. **Configuration System** - YAML-based configuration

### Medium Priority

1. **Batch Processing** - Multi-file processing capability
2. **Validation Module** - Refactor validation logic
3. **Generator Module** - Refactor document generation
4. **Error Handling** - Comprehensive error management

### Low Priority

1. **Advanced CLI Features** - Interactive mode, progress bars
2. **Plugin System** - Extensible architecture
3. **GUI Interface** - Optional graphical interface
4. **Performance Optimization** - Parallel processing, caching

## 8. Success Criteria

### Functional Requirements

- [ ] Process single documents with same quality as v0.4
- [ ] Handle batch processing of multiple documents
- [ ] Maintain backward compatibility with v0.4 outputs
- [ ] Provide clear error messages and debugging info
- [ ] Support configuration-driven processing

### Technical Requirements

- [ ] Clean, modular codebase
- [ ] Comprehensive test coverage
- [ ] Type hints throughout
- [ ] Clear documentation
- [ ] Easy installation and setup

### User Experience Requirements

- [ ] Intuitive CLI interface
- [ ] Clear progress reporting
- [ ] Helpful error messages
- [ ] Configuration examples
- [ ] Migration guide from v0.4

## 9. Risk Mitigation

### Technical Risks

- **Complexity**: Break down into smaller, manageable modules
- **Performance**: Profile and optimize critical paths
- **Compatibility**: Maintain test suite with v0.4 outputs

### Process Risks

- **Scope Creep**: Stick to defined phases and priorities
- **Quality**: Implement comprehensive testing from start
- **Timeline**: Regular checkpoints and progress reviews

## 10. Next Steps

1. **Create new repository** for v1.0
2. **Set up project structure** as defined above
3. **Start with core data models** (models.py)
4. **Refactor extraction logic** into modules
5. **Build basic CLI framework**
6. **Implement configuration system**

This plan provides a roadmap for transforming SpecConverter into a robust, maintainable, and user-friendly tool while preserving the valuable logic developed in v0.4.
