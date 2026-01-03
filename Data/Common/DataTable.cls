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

'========================================
' Properties
'========================================

Private name As String
Private columns() As DataColumn
Private rows() As DataRow

''' <summary>
''' Sets the name of the DataTable
''' </summary>
''' <param name="value">The name to assign to the table</param>
Public Property Let Name(value As String)
    name = value
End Property

''' <summary>
''' Gets the name of the DataTable
''' </summary>
''' <returns>The table name</returns>
Public Property Get Name() As String
    Name = name
End Property

''' <summary>
''' Gets a copy of the columns array
''' </summary>
''' <returns>Array of DataColumn objects</returns>
Public Property Get Columns() As DataColumn()
    Columns = columns
End Property

''' <summary>
''' Gets a copy of the rows array
''' </summary>
''' <returns>Array of DataRow objects</returns>
Public Property Get Rows() As DataRow()
    Rows = rows
End Property

'========================================
' Column Operations
'========================================

''' <summary>
''' Gets the total number of columns in the table
''' </summary>
''' <returns>The count of columns</returns>
Public Function ColumnCount() As Long
    On Error Resume Next
    If UBound(columns) < LBound(columns) Then
        ColumnCount = 0
    Else
        ColumnCount = UBound(columns) - LBound(columns) + 1
    End If
    On Error GoTo 0
End Function

''' <summary>
''' Gets the ordinal position of a column by its name
''' </summary>
''' <param name="columnName">The name of the column</param>
''' <returns>The zero-based index of the column, or -1 if not found</returns>
Public Function GetColumnOrdinalByName(columnName As String) As Long
    Dim i As Long

    If ColumnCount() = 0 Then
        GetColumnOrdinalByName = -1
        Exit Function
    End If

    For i = LBound(columns) To UBound(columns)
        If columns(i).Name = columnName Then
            GetColumnOrdinalByName = i
            Exit Function
        End If
    Next i

    ' If not found, return -1
    GetColumnOrdinalByName = -1
End Function

''' <summary>
''' Removes a column from the table by its name
''' </summary>
''' <param name="columnName">The name of the column to remove</param>
''' <returns>True if successful, False otherwise</returns>
''' <remarks>Also removes the corresponding data from all rows</remarks>
Public Function RemoveColumnByName(columnName As String) As Boolean
    Dim i As Long
    Dim j As Long
    Dim currentCount As Long
    Dim columnOrdinal As Long

    currentCount = ColumnCount()

    If currentCount = 0 Then
        RemoveColumnByName = False
        Exit Function
    End If

    columnOrdinal = GetColumnOrdinalByName(columnName)

    If columnOrdinal = -1 Then
        RemoveColumnByName = False
        Exit Function
    End If

    ' Shift remaining columns down in the array to fill the gap
    For j = columnOrdinal To UBound(columns) - 1
        Set columns(j) = columns(j + 1)
    Next j

    ' Resize the array to trim off the excess element at the end of the array.
    ReDim Preserve columns(0 To currentCount - 2)

    ' Clean up row data
    Call RemoveRowColumnDataByColumnOrdinal(columnOrdinal)

    RemoveColumnByName = True
End Function

''' <summary>
''' Removes a column from the table by its ordinal position
''' </summary>
''' <param name="ordinal">The zero-based index of the column to remove</param>
''' <returns>True if successful, False otherwise</returns>
Public Function RemoveColumnByOrdinal(ordinal As Long) As Boolean
    Dim i As Long
    Dim currentCount As Long

    ' Get the current number of columns in the array
    currentCount = ColumnCount()

    If currentCount = 0 Then
        RemoveColumnByOrdinal = False
        Exit Function
    End If

    ' Check if ordinal is valid
    If ordinal < LBound(columns) Or ordinal > UBound(columns) Then
        RemoveColumnByOrdinal = False
        Exit Function
    End If

    ' Shift remaining columns down in the array to fill the gap
    For i = ordinal To UBound(columns) - 1
        Set columns(i) = columns(i + 1)
    Next i

    ' Resize the array to trim off the excess element at the end of the array.
    ReDim Preserve columns(0 To currentCount - 2)

    RemoveColumnByOrdinal = True

End Function

''' <summary>
''' Removes all columns from the table
''' </summary>
''' <returns>True if successful</returns>
''' <remarks>Also clears all rows since they would have no data</remarks>
Public Function ClearColumns() As Boolean
    ReDim columns(0 To -1)

    ' A table with no columns should also have no rows since any rows would have no data.
    Call ClearRows
    
    ClearColumns = True
End Function

''' <summary>
''' Checks if a column with the specified name exists in the table
''' </summary>
''' <param name="columnName">The name of the column to search for</param>
''' <returns>True if the column exists, False otherwise</returns>
Public Function ContainsColumn(columnName As String) As Boolean
    Dim col As DataColumn
    Dim i As Long

    If ColumnCount() = 0 Then
        ContainsColumn = False
        Exit Function
    End If
    
    For i = LBound(columns) To UBound(columns)
        Set col = columns(i)

        If col.Name = columnName Then
            ContainsColumn = True
            Exit Function
        End If

    Next i

    ContainsColumn = False
End Function

Public Function AddColumn(ByRef column As DataColumn) As DataColumn
    Dim newCol As DataColumn
    Dim currentCount As Long

    If column Is Nothing Then
        Err.Raise vbObjectError + 511, "DataTable.AddColumn", "Cannot add a null column reference."
    End If

    If GetColumnOrdinalByName(column.Name) <> -1 Then
        Err.Raise vbObjectError + 512, "DataTable.AddColumn", "A column with the name '" & column.Name & "' already exists."
    End If

    'Get the current number of columns in the array
    currentCount = ColumnCount()

    ' Extend the array by one
    ReDim Preserve columns(0 To currentCount)

    'Add the new column at the end of the array
    Set columns(currentCount) = column

    Set AddColumn = column

End Function

'========================================
' Row Operations
'========================================

''' <summary>
''' Gets the total number of rows in the table
''' </summary>
''' <returns>The count of rows</returns>
Public Function RowCount() As Long
    On Error Resume Next
    If UBound(rows) < LBound(rows) Then
        RowCount = 0
    Else
        RowCount = UBound(rows) - LBound(rows) + 1
    End If
    On Error GoTo 0
End Function

''' <summary>
''' Adds a row to the table
''' </summary>
''' <param name="row">The DataRow to add</param>
''' <returns>The added DataRow reference</returns>
''' <remarks>
''' Raises error 513 if row is Nothing.
''' Raises error 514 if row column count doesn't match table column count.
''' </remarks>
Public Function AddRow(ByRef row As DataRow) As DataRow
    Dim newRow As DataRow
    Dim currentCount As Long

    If row Is Nothing Then
        Err.Raise vbObjectError + 513, "DataTable.AddRow", "Cannot add a null row reference."
    End If
    
    If row.ItemCount() <> Me.ColumnCount() Then
        Err.Raise vbObjectError + 514, "DataTable.AddRow", "Row column count does not match table column count."
    End If

    'Get the current number of rows in the array
    currentCount = RowCount()

    ' Extend the array by one
    ReDim Preserve rows(0 To currentCount)

    'Add the new row at the end of the array
    Set rows(currentCount) = row

    Set AddRow = row

End Function

''' <summary>
''' Gets a row by its ordinal position
''' </summary>
''' <param name="ordinal">The zero-based index of the row</param>
''' <returns>The DataRow object if found, Nothing if out of bounds</returns>
Public Function GetRowByOrdinal(ordinal As Long) As DataRow
    If RowCount() = 0 Then
        Set GetRowByOrdinal = Nothing
        Exit Function
    End If

    If ordinal >= LBound(rows) And ordinal <= UBound(rows) Then
        Set GetRowByOrdinal = rows(ordinal)
    Else
        ' If out of bounds, return Nothing
        Set GetRowByOrdinal = Nothing
    End If
End Function

''' <summary>
''' Removes a row from the table by its ordinal position
''' </summary>
''' <param name="ordinal">The zero-based index of the row to remove</param>
''' <returns>True if successful, False otherwise</returns>
Public Function RemoveRowByOrdinal(ordinal As Long) As Boolean
    Dim i As Long
    Dim currentCount As Long

    ' Get the current number of rows in the array
    currentCount = RowCount()

    If currentCount = 0 Then
        RemoveRowByOrdinal = False
        Exit Function
    End If

    ' Check if the ordinal is valid
    If ordinal < LBound(rows) Or ordinal > UBound(rows) Then
        RemoveRowByOrdinal = False
        Exit Function
    End If

    ' Shift remaining rows down in the array to fill the gap
    For i = ordinal To UBound(rows) - 1
        Set rows(i) = rows(i + 1)
    Next i

    ' Resize the array to trim off the excess element at the end of the array.
    ReDim Preserve rows(0 To currentCount - 2)

    RemoveRowByOrdinal = True

End Function

''' <summary>
''' Removes all rows from the table
''' </summary>
''' <returns>True if successful</returns>
Public Function ClearRows() As Boolean
    ReDim rows(0 To -1)
    ClearRows = True
End Function

''' <summary>
''' Removes data from a specific column in all rows
''' </summary>
''' <param name="columnOrdinal">The zero-based index of the column</param>
''' <returns>True if successful, False if invalid ordinal</returns>
''' <remarks>Private helper method used when removing columns</remarks>
Private Function RemoveRowColumnDataByColumnOrdinal(columnOrdinal As Long) As Boolean
    Dim i As Long

    If RowCount() = 0 Then
        RemoveRowColumnDataByColumnOrdinal = True
        Exit Function
    End If

    If ColumnCount() = 0 Then
        RemoveRowColumnDataByColumnOrdinal = True
        Exit Function
    End If

    If columnOrdinal < LBound(columns) Or columnOrdinal > UBound(columns) Then
        RemoveRowColumnDataByColumnOrdinal = False
        Exit Function
    End If

    For i = LBound(rows) To UBound(rows)
        Call rows(i).RemoveByColumnOrdinal(columnOrdinal)
    Next i

    RemoveRowColumnDataByColumnOrdinal = True
End Function

''' <summary>
''' Gets the value at a specific row and column position
''' </summary>
''' <param name="rowOrdinal">The zero-based row index</param>
''' <param name="columnOrdinal">The zero-based column index</param>
''' <returns>The value at the specified position</returns>
''' <remarks>
''' Raises error 515 if no columns or column ordinal out of bounds.
''' Raises error 516 if no rows or row ordinal out of bounds.
''' </remarks>
Public Function GetValueAt(rowOrdinal As Long, columnOrdinal As Long) As Variant
    Dim targetRow As DataRow
    Dim targetValue As Variant

    If ColumnCount() = 0 Then
        Err.Raise vbObjectError + 515, "DataTable.GetValueAt", "No columns in the table."
    End If

    If RowCount() = 0 Then
        Err.Raise vbObjectError + 516, "DataTable.GetValueAt", "No rows in the table."
    End If

    If columnOrdinal < LBound(columns) Or columnOrdinal > UBound(columns) Then
        Err.Raise vbObjectError + 515, "DataTable.GetValueAt", "Column ordinal " & columnOrdinal & " is out of bounds."
    End If

    If rowOrdinal < LBound(rows) Or rowOrdinal > UBound(rows) Then
        Err.Raise vbObjectError + 516, "DataTable.GetValueAt", "Row ordinal " & rowOrdinal & " is out of bounds."
    End If

    Set targetRow = GetRowByOrdinal(rowOrdinal)

    If targetRow Is Nothing Then
        Err.Raise vbObjectError + 516, "DataTable.GetValueAt", "Row ordinal " & rowOrdinal & " is out of bounds."
    End If

    targetValue = targetRow.GetValue(columnOrdinal)

    GetValueAt = targetValue
End Function

''' <summary>
''' Sets the value at a specific row and column position
''' </summary>
''' <param name="rowOrdinal">The zero-based row index</param>
''' <param name="columnOrdinal">The zero-based column index</param>
''' <param name="value">The value to set</param>
''' <remarks>
''' Raises error 515 if no columns or column ordinal out of bounds.
''' Raises error 516 if no rows or row ordinal out of bounds.
''' </remarks>
Public Sub SetValueAt(rowOrdinal As Long, columnOrdinal As Long, value As Variant)
    Dim targetRow As DataRow

    If ColumnCount() = 0 Then
        Err.Raise vbObjectError + 515, "DataTable.SetValueAt", "No columns in the table."
    End If

    If RowCount() = 0 Then
        Err.Raise vbObjectError + 516, "DataTable.SetValueAt", "No rows in the table."
    End If

    If columnOrdinal < LBound(columns) Or columnOrdinal > UBound(columns) Then
        Err.Raise vbObjectError + 515, "DataTable.SetValueAt", "Column ordinal " & columnOrdinal & " is out of bounds."
    End If

    If rowOrdinal < LBound(rows) Or rowOrdinal > UBound(rows) Then
        Err.Raise vbObjectError + 516, "DataTable.SetValueAt", "Row ordinal " & rowOrdinal & " is out of bounds."
    End If

    Set targetRow = GetRowByOrdinal(rowOrdinal)

    If targetRow Is Nothing Then
        Err.Raise vbObjectError + 516, "DataTable.SetValueAt", "Row ordinal " & rowOrdinal & " is out of bounds."
    End If

    Call targetRow.SetValue(columnOrdinal, value)

End Sub

'========================================
' Table Operations
'========================================

''' <summary>
''' Clears all columns and rows from the table
''' </summary>
''' <returns>True if successful</returns>
Public Function Clear() As Boolean
    ClearColumns
    ClearRows
    Clear = True
End Function

''' <summary>
''' Creates a copy of the table structure without data
''' </summary>
''' <returns>A new DataTable with the same columns but no rows</returns>
''' <remarks>Only clones column structure, not row data</remarks>
Public Function Clone() As DataTable
    Dim newTable As DataTable
    Dim i As Long
    Dim col As DataColumn

    ' Create a new DataTable instance
    Set newTable = New DataTable
    newTable.Name = Me.Name

    ' Clone columns
    If ColumnCount() > 0 Then
        For i = LBound(columns) To UBound(columns)
            Set col = columns(i).Clone()
            Call newTable.AddColumn(col)
        Next i
    End If

    ' Note: Rows are not cloned in this operation

    Set Clone = newTable
End Function

''' <summary>
''' Initializes a new DataTable instance
''' </summary>
''' <remarks>Sets default table name to "Table1"</remarks>
Private Sub Class_Initialize()
    ReDim columns(0 To -1)
    ReDim rows(0 To -1)

    name = "Table1"
End Sub

''' <summary>
''' Cleans up resources when the DataTable is destroyed
''' </summary>
''' <remarks>Releases all column and row object references</remarks>
Private Sub Class_Terminate()
    Dim i As Long

    ' Clean up columns
    On Error Resume Next

    If UBound(columns) >= LBound(columns) Then
        For i = LBound(columns) To UBound(columns)
            Set columns(i) = Nothing
        Next i
    End If
    
    ' Clean up rows
    If UBound(rows) >= LBound(rows) Then
        For i = LBound(rows) To UBound(rows)
            Set rows(i) = Nothing
        Next i
    End If

    On Error GoTo 0

    ReDim columns(0 To -1)
    ReDim rows(0 To -1)
End Sub

''' <summary>
''' Converts the table data to a 2D Variant array
''' </summary>
''' <returns>A 2D array containing all table values</returns>
Public Function ToArray() As Variant
    Dim result() As Variant
    Dim row As Long
    Dim col As Long
    Dim totalRows As Long
    Dim totalCols As Long

    totalRows = RowCount()
    totalCols = ColumnCount()

    ReDim result(0 To totalRows - 1, 0 To totalCols - 1)

    For row = 0 To totalRows - 1
        For col = 0 To totalCols - 1
            result(row, col) = Me.GetValueAt(row, col)
        Next col
    Next row

    ToArray = result
End Function

''' <summary>
''' Populates the table from a 2D Variant array
''' </summary>
''' <param name="data">The 2D array containing the data to import</param>
''' <returns>True if successful</returns>
''' <remarks>
''' Clears existing table data before importing.
''' Creates columns named "Column1", "Column2", etc.
''' </remarks>
Public Function FromArray(data As Variant) As Boolean
    Dim totalRows As Long
    Dim totalCols As Long
    Dim row As Long
    Dim col As Long
    Dim newRow As DataRow
    Dim newCol As DataColumn
    Dim i As Long

    ' Clear existing data
    Call Me.Clear()

    ' Determine dimensions
    totalRows = UBound(data, 1) - LBound(data, 1) + 1
    totalCols = UBound(data, 2) - LBound(data, 2) + 1

    ' Create columns
    For col = 0 To totalCols - 1
        Set newCol = New DataColumn
        newCol.Name = "Column" & CStr(col + 1)

        Call Me.AddColumn(newCol)
    Next col

    ' Create rows and populate data
    For row = 0 To totalRows - 1
        Set newRow = New DataRow

        ' Add default values for each column
        For col = 0 To totalCols - 1
            Call newRow.AddValue(data(LBound(data, 1) + row, LBound(data, 2) + col))
        Next col

        Call Me.AddRow(newRow)
    Next row

    FromArray = True
End Function

''' <summary>
''' Fills the DataTable with data from an ADO Recordset
''' </summary>
''' <param name="recordSet">The ADO Recordset to read from</param>
''' <remarks>
''' Clears existing table data and creates columns based on recordset fields.
''' Preserves the recordset's current position after filling.
''' Raises error 519 if recordset is Nothing.
''' </remarks>
Public Sub Fill(ByVal recordSet As ADODB.Recordset)
    Dim newCol As DataColumn
    Dim newRow As DataRow
    Dim i As Long
    Dim originalPosition As Variant
    Dim hasBookmark As Boolean

    ' Clear existing data
    Call Me.Clear()

    If recordSet Is Nothing Then
        Err.Raise vbObjectError + 519, "DataTable.Fill", "Recordset reference is null."
    End If

    If recordSet.EOF And recordSet.BOF Then
        ' Empty recordset, nothing to fill
        Exit Sub
    End If

    ' Save current position if recordset supports bookmarks
    On Error Resume Next

    hasBookmark = recordSet.Supports(adBookmark)
    
    If hasBookmark And (Not recordSet.BOF And Not recordSet.EOF) Then
        originalPosition = recordSet.Bookmark
    End If
    
    On Error GoTo 0

    ' Move to first record
    If Not recordSet.BOF Then
        recordSet.MoveFirst
    End If

    ' Create columns based on recordset fields
    For i = 0 To recordSet.Fields.Count - 1
        Set newCol = New DataColumn
        newCol.Name = recordSet.Fields(i).Name

        Call Me.AddColumn(newCol)
    Next i

    ' Populate rows from recordset
    Do While Not recordSet.EOF
        Set newRow = New DataRow

        ' Add values for each column
        For i = 0 To recordSet.Fields.Count - 1
            Call newRow.AddValue(recordSet.Fields(i).Value)
        Next i

        Call Me.AddRow(newRow)

        recordSet.MoveNext
    Loop

    ' Restore original position if possible
    On Error Resume Next

    If hasBookmark And (Not IsEmpty(originalPosition)) Then
        recordSet.Bookmark = originalPosition
    End If

    On Error GoTo 0
End Sub