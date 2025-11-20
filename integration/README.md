# Integration Tests for Rate Import Service

This directory contains integration tests that verify the interaction between different components of the Rate Import Service.

## Test Categories

### 🗄️ Database Integration (`test_database_integration.py`)
- Database connectivity across modules
- Cross-module database operations  
- Transaction consistency
- Schema compatibility

### 🌐 API Integration (`test_api_integration.py`)
- JeraSoft API workflow integration
- Microsoft Graph API authentication
- API error handling and resilience
- Timeout and retry mechanisms

### 🔄 Workflow Integration (`test_workflow_integration.py`)  
- End-to-end email processing workflow
- File processing to database workflow
- Database to JeraSoft upload workflow
- Data flow integration
- Error recovery workflows

### 📁 File System Integration (`test_file_integration.py`)
- File system operations across modules
- Attachment processing workflows
- Concurrent file operations
- Data persistence and recovery
- File type validation

## Running Integration Tests

### Run All Integration Tests
```bash
python integration/run_integration_tests.py
```

### Run Specific Category
```bash
python integration/run_integration_tests.py --category database
python integration/run_integration_tests.py --category api
python integration/run_integration_tests.py --category workflow
python integration/run_integration_tests.py --category file
```

### Run with Unittest
```bash
python -m unittest discover integration -v
```

### Run Individual Test Files
```bash
python -m unittest integration.test_database_integration -v
python -m unittest integration.test_api_integration -v
python -m unittest integration.test_workflow_integration -v
python -m unittest integration.test_file_integration -v
```

## Integration Test Focus

### What Integration Tests Check:
- ✅ **Component Interaction** - How modules work together
- ✅ **Data Flow** - Data passing between components
- ✅ **End-to-End Workflows** - Complete process flows
- ✅ **Cross-Module Dependencies** - Module interdependencies
- ✅ **System Integration** - External system connections
- ✅ **Error Propagation** - How errors flow through the system
- ✅ **Resource Sharing** - Shared database, files, etc.

### What Integration Tests Don't Check:
- ❌ Individual module functionality (covered by unit tests)
- ❌ Detailed business logic (covered by unit tests)
- ❌ Edge cases within modules (covered by unit tests)

## Test Environment

Integration tests use:
- **Mock external services** (JeraSoft API, Microsoft Graph)
- **Temporary file systems** for file operations
- **Mocked database connections** for database tests
- **Isolated test environments** to prevent interference

## Expected Results

Integration tests should demonstrate:
- 🔗 **Seamless component integration**
- 📊 **Proper data flow between modules**
- 🛡️ **Error handling across boundaries**
- ⚡ **Performance under integrated load**
- 🔄 **Recovery from integration failures**

## Troubleshooting

### Common Issues:
1. **Import Errors** - Ensure all modules are available
2. **Path Issues** - Check that parent directory is in Python path
3. **Mock Failures** - Verify mock setup matches actual interfaces
4. **Resource Cleanup** - Ensure temporary resources are cleaned up

### Debug Tips:
- Use `--verbose` flag for detailed output
- Run individual test files to isolate issues
- Check that all required modules can be imported
- Verify mock configurations match actual module interfaces

## Integration Test Philosophy

Integration tests complement unit tests by:
- **Testing the "glue"** between components
- **Validating assumptions** about inter-module communication
- **Ensuring compatibility** between different parts of the system
- **Catching issues** that only appear when components work together

These tests provide confidence that the Rate Import Service works as a cohesive system, not just as individual components.