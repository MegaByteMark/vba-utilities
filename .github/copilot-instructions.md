# Copilot Instructions for VBA Utilities

## Repository Overview

This is a collection of reusable VBA (Visual Basic for Applications) utility modules for Microsoft Access and other Office applications. The project provides replacements for common functionality missing in VBA, including data abstraction (DataTable/DataSet), serialization (JSON/CSV), and SQL Server database access.

## Project Structure

- **Data/** - Core data abstraction layer
  - **Common/** - Foundation classes: DataTable, DataRow, DataColumn, DataSet (ADO-like data structures)
  - **SqlServer/** - SQL Server provider implementations
- **Serialization/** - Data serialization modules (JSON, CSV)
- **Text/** - String utilities (StringCollection)
- **Windows/** - Windows interop modules (EventLog, Impersonation)

## Installation & Setup

Users install modules by importing `.cls` files into their MS Access projects via the VBA editor:
1. `Alt + F11` to open VBA editor
2. `File -> Import File...`
3. Select the `.cls` file

Key dependencies that users must reference in their Access projects:
- Microsoft ActiveX Data Objects (ADO) 6.1 Library
- Microsoft Scripting Runtime (for Dictionary objects)

## VBA Code Conventions

### Documentation Style
- Each `.cls` file starts with a header block containing:
  - MIT License notice
  - Class description (between `'========` markers)
  - Dependencies list
  - Custom error codes with descriptions
  - Usage example
- Functions have inline comments with `@param` and `@return` annotations

### Naming & Structure
- Class files: PascalCase (e.g., `SqlServerDataProvider.cls`)
- Module files: PascalCase with `.bas` extension
- Private variables with "Camel" case prefix (e.g., `connString`, `cmdTimeout`)
- Properties use `Property Let`/`Property Get` patterns
- `Option Explicit` mandatory at top of each file

### Error Handling
- Custom error codes documented in header comments (typically starting at 511+)
- Errors raised with: `Err.Raise code, "MethodName", "Description"`
- No global exception handling; errors propagate to caller

### Parameter Passing
- Dictionary objects (from MS Scripting Runtime) used for optional/named parameters
- Method naming: `Verb` + `Noun` pattern (e.g., `ExecuteNonQuery`, `ExecuteScalar`, `GetRecordset`)

## Build & Quality Controls

### CRLF Enforcement
- Workflow: `.github/workflows/enforce-crlf.yml`
- **Line endings must be CRLF** for `.cls`, `.bas`, and `.frm` files
- Automatically enforced on push to `main` branch
- When modifying VBA files, ensure CRLF line endings (not LF)

### No Automated Tests
- This is a library project with manual/integration testing
- Users test by importing into Access and verifying behavior

## Key Architectural Patterns

### Data Abstraction Layer
The Data/Common modules mimic .NET DataTable/DataSet architecture:
- `DataTable`: Container for columns and rows
- `DataRow`: Single row within a table
- `DataColumn`: Column definition (name, type)
- `DataSet`: Collection of DataTable objects

Classes are interconnected—DataTable contains DataColumn and DataRow objects, which reference each other.

### SQL Server Integration
- `SqlServerDataProvider`: Main class for database operations (Connection, ExecuteNonQuery, ExecuteScalar, GetRecordset)
- `SqlConnectionStringBuilder`: Helper for building connection strings
- `SqlBulkCopy`: Bulk insert operations
- Supports parameterized queries with Dictionary objects

### Serialization
- `JsonSerializer`: Converts DataTable ↔ JSON
- `CsvSerializer`: Converts DataTable ↔ CSV
- Both handle headers and nullable values

## Contributing

When adding new modules:
- Follow the documentation header pattern (see existing `.cls` files)
- Use Option Explicit
- Include usage examples in the header
- List dependencies and custom error codes
- Ensure CRLF line endings before committing
- Update README.md if adding a major new module category

## Common Tasks

- **Adding a new utility module**: Create a `.cls` file in appropriate subdirectory, include header docs, ensure CRLF
- **Modifying existing class**: Check for dependent classes (listed in header "Dependencies")
- **Testing changes**: Users import the module into an Access database and test manually (no automated test framework)
