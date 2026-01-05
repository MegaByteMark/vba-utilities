' filepath: /Users/markhart/Documents/github-personal/vba-utilities/Data/SqlServer/SqlBulkCopy.cls
' MIT License
'
' Copyright (c) 2025 MegaByteMark (https://github.com/MegaByteMark)
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.

Option Explicit

Private Const DEFAULT_BATCH_SIZE As Long = 100
Private Const DEFAULT_BULK_COPY_TIMEOUT As Long = 30 ' seconds

'========================================
' Properties
'========================================

Private _batchSize As Long
Private _bulkCopyTimeout As Long
Private _destinationTableName As String
Private _destinationConnectionString As String
Private _columnMappings As Scripting.Dictionary
Private _identityColumns() As String

''' <summary>
''' Sets the number of rows to insert in each batch
''' </summary>
''' <param name="value">The batch size as a Long value. Must be greater than 0</param>
''' <remarks>Default value is 100</remarks>
Public Property Let BatchSize(value As Long)
    _batchSize = value
End Property

''' <summary>
''' Gets the number of rows to insert in each batch
''' </summary>
''' <returns>The batch size</returns>
Public Property Get BatchSize() As Long
    BatchSize = _batchSize
End Property

''' <summary>
''' Sets the timeout in seconds for bulk copy operations
''' </summary>
''' <param name="value">The timeout value in seconds</param>
''' <remarks>Default value is 30 seconds</remarks>
Public Property Let BulkCopyTimeout(value As Long)
    _bulkCopyTimeout = value
End Property

''' <summary>
''' Gets the timeout in seconds for bulk copy operations
''' </summary>
''' <returns>The timeout value in seconds</returns>
Public Property Get BulkCopyTimeout() As Long
    BulkCopyTimeout = _bulkCopyTimeout
End Property

''' <summary>
''' Sets the name of the destination table in SQL Server
''' </summary>
''' <param name="value">The table name. Must be a valid SQL identifier</param>
Public Property Let DestinationTableName(value As String)
    _destinationTableName = value
End Property

''' <summary>
''' Gets the name of the destination table in SQL Server
''' </summary>
''' <returns>The destination table name</returns>
Public Property Get DestinationTableName() As String
    DestinationTableName = _destinationTableName
End Property

''' <summary>
''' Sets the connection string for the destination SQL Server database
''' </summary>
''' <param name="value">The ADO connection string</param>
Public Property Let DestinationConnectionString(value As String)
    _destinationConnectionString = value
End Property

''' <summary>
''' Gets the connection string for the destination SQL Server database
''' </summary>
''' <returns>The connection string</returns>
Public Property Get DestinationConnectionString() As String
    DestinationConnectionString = _destinationConnectionString
End Property

''' <summary>
''' Gets the Dictionary of column mappings from source to destination columns
''' </summary>
''' <returns>A Scripting.Dictionary where keys are source column names and values are destination column names</returns>
Public Property Get ColumnMappings() As Scripting.Dictionary
    Set ColumnMappings = _columnMappings
End Property

''' <summary>
''' Gets the array of identity column names that require IDENTITY_INSERT to be enabled
''' </summary>
''' <returns>An array of String values containing identity column names</returns>
Public Property Get IdentityColumns() As Variant
    IdentityColumns = _identityColumns
End Property

''' <summary>
''' Sets the array of identity column names that require IDENTITY_INSERT to be enabled
''' </summary>
''' <param name="value">A Variant array of String values containing identity column names</param>
''' <remarks>Raises error 513 if value is not an array</remarks>
Public Property Set IdentityColumns(value As Variant)
    Dim i As Long

    If Not IsArray(value) Then
        Err.Raise vbObjectError + 513, "SqlBulkCopy", "IdentityColumns must be an array."
    End If

    ' Clear existing identity columns
    ReDim _identityColumns(0 To -1) 

    ' Set new identity columns
    ReDim _identityColumns(LBound(value) To UBound(value))

    For i = LBound(value) To UBound(value)
        _identityColumns(i) = value(i)
    Next i
End Property

''' <summary>
''' Initializes a new SqlBulkCopy instance
''' </summary>
''' <remarks>Sets default batch size to 100 and timeout to 30 seconds</remarks>
Private Sub Class_Initialize()
    Set _columnMappings = New Scripting.Dictionary
    _batchSize = DEFAULT_BATCH_SIZE
    _bulkCopyTimeout = DEFAULT_BULK_COPY_TIMEOUT
    _identityColumns = Array()
End Sub

''' <summary>
''' Cleans up resources when the SqlBulkCopy is destroyed
''' </summary>
Private Sub Class_Terminate()
    Set _columnMappings = Nothing
End Sub

''' <summary>
''' Checks if an array is empty (has no elements)
''' </summary>
''' <param name="arr">The Variant array to check</param>
''' <returns>True if the array is empty, False otherwise</returns>
''' <remarks>Private helper method</remarks>
Private Function IsArrayEmpty(arr As Variant) As Boolean
    On Error Resume Next
    IsArrayEmpty = (UBound(arr) < LBound(arr))
    On Error GoTo 0
End Function

''' <summary>
''' Adds a column mapping from source column to destination column
''' </summary>
''' <param name="sourceColumn">The name of the column in the source DataTable</param>
''' <param name="destinationColumn">The name of the column in the destination SQL Server table</param>
Public Sub AddColumnMapping(sourceColumn As String, destinationColumn As String)
    If Not _columnMappings.Exists(sourceColumn) Then
        _columnMappings.Add sourceColumn, destinationColumn
    End If
End Sub

''' <summary>
''' Removes all column mappings
''' </summary>
Public Sub ClearColumnMappings()
    _columnMappings.RemoveAll
End Sub

''' <summary>
''' Fluent API method to set the destination table name
''' </summary>
''' <param name="DestinationTableName">The name of the destination table</param>
''' <returns>The current SqlBulkCopy instance for method chaining</returns>
Public Function HasDestinationTableName(ByVal DestinationTableName As String) As SqlBulkCopy
    _destinationTableName = DestinationTableName
    Set HasDestinationTableName = Me
End Function

''' <summary>
''' Fluent API method to set the destination connection string
''' </summary>
''' <param name="DestinationConnectionString">The ADO connection string for the destination database</param>
''' <returns>The current SqlBulkCopy instance for method chaining</returns>
Public Function HasDestinationConnectionString(ByVal DestinationConnectionString As String) As SqlBulkCopy
    _destinationConnectionString = DestinationConnectionString
    Set HasDestinationConnectionString = Me
End Function

''' <summary>
''' Fluent API method to set the batch size
''' </summary>
''' <param name="BatchSize">The number of rows to process in each batch</param>
''' <returns>The current SqlBulkCopy instance for method chaining</returns>
Public Function HasBatchSize(ByVal BatchSize As Long) As SqlBulkCopy
    _batchSize = BatchSize
    Set HasBatchSize = Me
End Function

''' <summary>
''' Fluent API method to set the bulk copy timeout
''' </summary>
''' <param name="BulkCopyTimeout">The timeout in seconds for the operation</param>
''' <returns>The current SqlBulkCopy instance for method chaining</returns>
Public Function HasBulkCopyTimeout(ByVal BulkCopyTimeout As Long) As SqlBulkCopy
    _bulkCopyTimeout = BulkCopyTimeout
    Set HasBulkCopyTimeout = Me
End Function

''' <summary>
''' Fluent API method to add a column mapping
''' </summary>
''' <param name="sourceColumn">The name of the source column</param>
''' <param name="destinationColumn">The name of the destination column</param>
''' <returns>The current SqlBulkCopy instance for method chaining</returns>
Public Function HasColumnMapping(sourceColumn As String, destinationColumn As String) As SqlBulkCopy
    Call AddColumnMapping(sourceColumn, destinationColumn)
    Set HasColumnMapping = Me
End Function

''' <summary>
''' Fluent API method to set identity columns
''' </summary>
''' <param name="identityColumns">An array of column names that are identity columns</param>
''' <returns>The current SqlBulkCopy instance for method chaining</returns>
Public Function HasIdentityColumns(identityColumns() As String) As SqlBulkCopy
    Dim i As Long

    ReDim _identityColumns(LBound(identityColumns) To UBound(identityColumns))
    
    For i = LBound(identityColumns) To UBound(identityColumns)
        _identityColumns(i) = identityColumns(i)
    Next i

    Set HasIdentityColumns = Me
End Function

''' <summary>
''' Writes data from a DataTable to the SQL Server destination table using batch inserts with transaction support
''' </summary>
''' <param name="dataTable">The DataTable containing the data to insert</param>
''' <remarks>
''' Uses transaction-based batch processing for optimal performance.
''' Batch size can be configured via the BatchSize property.
''' If identity columns are specified, IDENTITY_INSERT will be enabled automatically.
''' Raises error 514 if DestinationTableName is not set.
''' Raises error 515 if DestinationConnectionString is not set.
''' Raises error 516 if no column mappings defined.
''' Raises error 517 if DataTable is not provided.
''' Raises error 518 if DataTable contains no rows.
''' Raises error 519 if user lacks INSERT permission on destination table.
''' Raises error 520 if identity column is not mapped in column mappings.
''' Raises error 521 if destination table does not exist.
''' Raises error 522 if destination table does not have required column.
''' Raises error 523 if destination table name is not a valid SQL identifier.
''' Raises error 524 if destination column name is not a valid SQL identifier.
''' </remarks>
Public Sub WriteToServer(dataTable As DataTable)
    Dim success As Boolean
    Dim conn As ADODB.Connection
    Dim i As Long
    Dim sourceColumn As Variant
    Dim destinationColumn As Variant
    Dim permissionsRecordset As ADODB.Recordset
    Dim endRow As Long
    Dim j As Long
    Dim insertColumns As String
    Dim insertValues As String
    Dim insertQuery As String
    Dim row As DataRow
    Dim valueList As String
    Dim identityColumn As Variant
    Dim tableExists As Variant
    Dim permissionsQuery As String

    If _destinationTableName = "" Then
        Err.Raise vbObjectError + 514, "SqlBulkCopy", "DestinationTableName is not set."
    End If

    If _destinationConnectionString = "" Then
        Err.Raise vbObjectError + 515, "SqlBulkCopy", "DestinationConnectionString is not set."
    End If

    If _columnMappings.Count = 0 Then
        Err.Raise vbObjectError + 516, "SqlBulkCopy", "No column mappings defined."
    End If

    If dataTable Is Nothing Then
        Err.Raise vbObjectError + 517, "SqlBulkCopy", "DataTable is not provided."
    End If

    If dataTable.Rows.Count = 0 Then
        Err.Raise vbObjectError + 518, "SqlBulkCopy", "DataTable contains no rows."
    End If

    If _batchSize <= 0 Then
        _batchSize = DEFAULT_BATCH_SIZE
    End If

    ' Ensure table name is valid
    If Not ValidateSqlIdentifier(_destinationTableName) Then
        Err.Raise vbObjectError + 523, "SqlBulkCopy", "Destination table name '" & _destinationTableName & "' is not a valid SQL identifier."
    End If

    ' Ensure all column mappings are valid SQL identifiers
    For Each sourceColumn In _columnMappings.Keys
        destinationColumn = _columnMappings(sourceColumn)

        If Not ValidateSqlIdentifier(destinationColumn) Then
            Err.Raise vbObjectError + 524, "SqlBulkCopy", "Destination column name '" & destinationColumn & "' is not a valid SQL identifier."
        End If
    Next sourceColumn

    ' Ensure the number of identity columns is valid
    If Not IsArrayEmpty(_identityColumns) Then
        For Each identityColumn In _identityColumns
            If Not _columnMappings.Exists(identityColumn) Then
                Err.Raise vbObjectError + 520, "SqlBulkCopy", "Identity column '" & identityColumn & "' is not mapped in the column mappings."
            End If
        Next identityColumn
    End If

    ' Create and open the destination connection
    Set conn = New ADODB.Connection
    conn.CommandTimeout = _bulkCopyTimeout
    conn.Open _destinationConnectionString

    'Start Transaction
    conn.BeginTrans

    On Error GoTo RollbackAndError

    ' Ensure the user has permission to perform bulk copy
    permissionsQuery = "SELECT " & _
        "HAS_PERMS_BY_NAME('" & _destinationTableName & "', 'OBJECT', 'INSERT') AS CanInsert"

    Set permissionsRecordset = GetRecordset(conn, permissionsQuery)

    If permissionsRecordset Is Nothing _
        Or IsNull(permissionsRecordset.Fields(0).Value) _
        Or CBool(permissionsRecordset.Fields(0).Value) = False Then
        Err.Raise vbObjectError + 519, "SqlBulkCopy", "The user does not have permission to insert data in the database. The INSERT permission is required."
    End If

    permissionsRecordset.Close
    Set permissionsRecordset = Nothing

    ' Ensure the destination table exists
    tableExists = ExecuteScalar(conn, "SELECT CASE WHEN OBJECT_ID('[" & _destinationTableName & "]', 'U') IS NOT NULL THEN 1 ELSE 0 END")

    If IsNull(tableExists) Or CInt(tableExists) = 0 Then
        Err.Raise vbObjectError + 521, "SqlBulkCopy", "Destination table '" & _destinationTableName & "' does not exist in the database."
    End If

    'Ensure the destination table has the required columns
    For Each sourceColumn In _columnMappings.Keys
        destinationColumn = _columnMappings(sourceColumn)
        tableExists = ExecuteScalar(conn, "SELECT CASE WHEN COL_LENGTH('[" & _destinationTableName & "]', '" & destinationColumn & "') IS NOT NULL THEN 1 ELSE 0 END")

        If IsNull(tableExists) Or CInt(tableExists) = 0 Then
            Err.Raise vbObjectError + 522, "SqlBulkCopy", "Destination table '" & _destinationTableName & "' does not have the required column '" & destinationColumn & "'."
        End If
    Next sourceColumn

    ' Ensure identity insert if there are identity columns
    If Not IsArrayEmpty(_identityColumns) Then
        If UBound(_identityColumns) >= LBound(_identityColumns) Then
            ExecuteNonQuery conn, "SET IDENTITY_INSERT " & _destinationTableName & " ON"
        End If
    End If

    ' Write data to server
    For i = 0 To dataTable.Rows.Count - 1 Step _batchSize
        endRow = 0
        j = 0
        insertColumns = ""
        insertValues = ""
        insertQuery = ""

        ' Determine the end row for the current batch
        endRow = Application.Min(i + _batchSize - 1, dataTable.Rows.Count - 1)

        ' Build the insert query for the batch
        For j = i To endRow
            Set row = dataTable.Rows(j)
            valueList = ""

            For Each sourceColumn In _columnMappings.Keys
                destinationColumn = _columnMappings(sourceColumn)

                If insertColumns = "" Then
                    insertColumns = destinationColumn
                ElseIf InStr(insertColumns, destinationColumn) = 0 Then
                    insertColumns = insertColumns & ", " & destinationColumn
                End If

                If valueList = "" Then
                    valueList = FormatValueForSql(row(sourceColumn))
                Else
                    valueList = valueList & ", " & FormatValueForSql(row(sourceColumn))
                End If
            Next sourceColumn

            If insertValues = "" Then
                insertValues = "(" & valueList & ")"
            Else
                insertValues = insertValues & ", (" & valueList & ")"
            End If
        Next j

        insertQuery = "INSERT INTO " & _destinationTableName & " (" & insertColumns & ") VALUES " & insertValues

        ' Execute the insert query
        ExecuteNonQuery conn, insertQuery

    Next i

    ' If identity insert was enabled, turn it off
    If Not IsArrayEmpty(_identityColumns) Then
        If UBound(_identityColumns) >= LBound(_identityColumns) Then
            ExecuteNonQuery conn, "SET IDENTITY_INSERT " & _destinationTableName & " OFF"
        End If
    End If

    ' Commit transaction if everything is successful
    On Error Resume Next
    conn.CommitTrans
    On Error GoTo 0

    ' Clean up
    conn.Close
    Set conn = Nothing

    Exit Sub

RollbackAndError:
    If Err.Number <> 0 Then
        ' Rollback transaction on error
        If Not conn Is Nothing Then
            On Error Resume Next
            conn.RollbackTrans
            On Error GoTo 0
        End If

        ' Clean up
        If Not conn Is Nothing Then
            conn.Close
        End If
        Set conn = Nothing

        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Sub

''' <summary>
''' Validates that a SQL identifier (table or column name) contains only alphanumeric characters and underscores
''' </summary>
''' <param name="identifier">The SQL identifier to validate</param>
''' <returns>True if valid, False otherwise</returns>
''' <remarks>
''' Ensures identifier starts with a letter or underscore.
''' Private helper method for SQL injection prevention.
''' </remarks>
Private Function ValidateSqlIdentifier(identifier As String) As Boolean
    Dim regex As Object

    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "^[a-zA-Z_][a-zA-Z0-9_]*$"

    ValidateSqlIdentifier = regex.Test(identifier)
End Function

''' <summary>
''' Formats a VBA value for SQL insertion
''' </summary>
''' <param name="value">The Variant value to format</param>
''' <returns>The SQL-formatted value as a string</returns>
''' <remarks>
''' Handles NULL, numeric, date, and string values with proper escaping.
''' Escapes single quotes in strings for basic SQL injection prevention.
''' Private helper method.
''' </remarks>
Private Function FormatValueForSql(value As Variant) As String
    If IsNull(value) Then
        FormatValueForSql = "NULL"
    ElseIf IsNumeric(value) Then
        FormatValueForSql = CStr(value)
    ElseIf IsDate(value) Then
        FormatValueForSql = "'" & Format(value, "yyyy-mm-dd hh:nn:ss") & "'"
    Else
        ' Escape single quotes in strings
        FormatValueForSql = "'" & Replace(CStr(value), "'", "''") & "'"
    End If
End Function

''' <summary>
''' Executes a SQL command that does not return a result set
''' </summary>
''' <param name="conn">The active ADODB.Connection object</param>
''' <param name="query">The SQL query string to execute</param>
''' <remarks>
''' Used for INSERT, UPDATE, DELETE, and DDL commands.
''' Private helper method.
''' </remarks>
Private Sub ExecuteNonQuery(conn As ADODB.Connection, query As String)
    Dim cmd As ADODB.Command
    
    Set cmd = New ADODB.Command

    With cmd
        .ActiveConnection = conn
        .CommandText = query
        .CommandType = adCmdText
        .Execute
    End With

    Set cmd = Nothing
End Sub

''' <summary>
''' Executes a SQL query and returns the first column of the first row as a Variant
''' </summary>
''' <param name="conn">The active ADODB.Connection object</param>
''' <param name="query">The SQL query string to execute</param>
''' <returns>The scalar value from the query, or Null if no rows returned</returns>
''' <remarks>Private helper method</remarks>
Private Function ExecuteScalar(conn As ADODB.Connection, query As String) As Variant
    Dim cmd As ADODB.Command
    Dim result As Variant
    Dim rs As ADODB.Recordset
    
    Set cmd = New ADODB.Command

    With cmd
        .ActiveConnection = conn
        .CommandText = query
        .CommandType = adCmdText
        Set rs = .Execute
    End With

    If Not rs.EOF Then
        result = rs.Fields(0).Value
    Else
        result = Null
    End If

    rs.Close
    Set rs = Nothing
    Set cmd = Nothing

    ExecuteScalar = result
End Function

''' <summary>
''' Executes a SQL query and returns the result as an ADODB.Recordset
''' </summary>
''' <param name="conn">The active ADODB.Connection object</param>
''' <param name="query">The SQL query string to execute</param>
''' <returns>The recordset containing query results</returns>
''' <remarks>Private helper method</remarks>
Private Function GetRecordset(conn As ADODB.Connection, query As String) As ADODB.Recordset
    Dim cmd As ADODB.Command
    
    Set cmd = New ADODB.Command

    With cmd
        .ActiveConnection = conn
        .CommandText = query
        .CommandType = adCmdText
        Set GetRecordset = .Execute 
    End With

    Set cmd = Nothing
End Function
