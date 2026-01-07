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


'========================================
' DataSet Class
'========================================
' Provides a collection of DataTables
'
' Dependencies:
'  - DataTable.cls
'
' Error Codes:
'  512 - Cannot add a Nothing table to the DataSet
'  513 - A table with the specified name already exists in the DataSet
'  514 - A table with the specified name already exists after renaming attempt
'
' Usage:
'  Dim ds As DataSet
'  Set ds = New DataSet
'  ds.Name = "MyDataSet"
'  Dim tbl As DataTable
'  Set tbl = New DataTable
'  tbl.Name = "MyTable"
'  Call ds.AddTable(tbl)
'
' Author: MegaByteMark
' Version: 2.0
'========================================

Option Explicit

'========================================
' Properties
'========================================

Private name As String
Private tables() As DataTable

''' <summary>
''' Sets the name of the DataSet
''' </summary>
''' <param name="value">The name to assign</param>
Public Property Let Name(value As String)
    name = value
End Property

''' <summary>
''' Gets the name of the DataSet
''' </summary>
''' <returns>The DataSet name</returns>
Public Property Get Name() As String
    Name = name
End Property

''' <summary>
''' Gets a copy of the tables array
''' </summary>
''' <returns>Array of DataTable objects</returns>
Public Property Get Tables() As DataTable()
    Tables = tables
End Property

'========================================
' DataSet Operations
'========================================

''' <summary>
''' Removes all tables from the DataSet
''' </summary>
''' <returns>True if successful</returns>
Public Function ClearTables() As Boolean
    Dim i As Long
    
    ' Clean up object references
    On Error Resume Next

    If UBound(tables) >= LBound(tables) Then
        For i = LBound(tables) To UBound(tables)
            Set tables(i) = Nothing
        Next i
    End If
    
    On Error GoTo 0
    
    ReDim tables(0 To -1)
    ClearTables = True
End Function

''' <summary>
''' Clears all tables from the DataSet
''' </summary>
''' <returns>True if successful</returns>
Public Function Clear() As Boolean
    Call ClearTables
    Clear = True
End Function

''' <summary>
''' Creates a deep copy of the DataSet structure including all tables
''' </summary>
''' <returns>A new DataSet with cloned tables</returns>
Public Function Clone() As DataSet
    Dim newDataSet As DataSet
    Dim i As Long
    Dim tbl As DataTable
    Dim newTbl As DataTable

    Set newDataSet = New DataSet
    newDataSet.Name = Me.Name

    ' Clone each table
    On Error GoTo CloneError
    
    If UBound(tables) >= LBound(tables) Then
        For i = LBound(tables) To UBound(tables)
            Set tbl = tables(i)
            If Not tbl Is Nothing Then  ' Add this check
                Set newTbl = tbl.Clone()
                Call newDataSet.AddTable(newTbl)
            End If
        Next i
    End If

    Set Clone = newDataSet
    Exit Function
    
CloneError:
    Set Clone = Nothing
    Err.Raise Err.Number, "DataSet.Clone", "Error cloning dataset: " & Err.Description
End Function

''' <summary>
''' Gets the index position of a table by its name
''' </summary>
''' <param name="tableName">The name of the table to find</param>
''' <returns>The zero-based index, or -1 if not found</returns>
Public Function GetDataTableOrdinalByName(tableName As String) As Long
    Dim tbl As DataTable
    Dim i As Long

    On Error Resume Next

    If UBound(tables) < LBound(tables) Then
        On Error GoTo 0
        GetDataTableOrdinalByName = -1
        Exit Function
    End If
    
    On Error GoTo 0

    For i = LBound(tables) To UBound(tables)
        Set tbl = tables(i)

        If tbl.Name = tableName Then
            GetDataTableOrdinalByName = i
            Exit Function
        End If

    Next i

    ' If not found, return -1
    GetDataTableOrdinalByName = -1
End Function

'========================================
' Table Operations
'========================================

''' <summary>
''' Checks if a table with the specified name exists in the DataSet
''' </summary>
''' <param name="tableName">The name of the table to check</param>
''' <returns>True if the table exists, False otherwise</returns>
Public Function ContainsTable(tableName As String) As Boolean
    Dim tbl As DataTable
    Dim i As Long

    On Error Resume Next

    If UBound(tables) >= LBound(tables) Then
        For i = LBound(tables) To UBound(tables)
            Set tbl = tables(i)

            If tbl.Name = tableName Then
                ContainsTable = True
                Exit Function
            End If

        Next i
    End If

    On Error GoTo 0

    ContainsTable = False
End Function

''' <summary>
''' Gets a DataTable by its name
''' </summary>
''' <param name="tableName">The name of the table to retrieve</param>
''' <returns>The DataTable object if found, Nothing otherwise</returns>
Public Function GetTableByName(tableName As String) As DataTable
    Dim tbl As DataTable
    Dim i As Long

    On Error Resume Next

    If UBound(tables) < LBound(tables) Then
        On Error GoTo 0
        Set GetTableByName = Nothing
        Exit Function
    End If

    On Error GoTo 0

    For i = LBound(tables) To UBound(tables)
        Set tbl = tables(i)

        If tbl.Name = tableName Then
            Set GetTableByName = tbl
            Exit Function
        End If

    Next i

    Set GetTableByName = Nothing
End Function

''' <summary>
''' Gets a DataTable by its ordinal position
''' </summary>
''' <param name="ordinal">The zero-based index of the table to retrieve</param>
''' <returns>The DataTable object if found, Nothing if ordinal is out of bounds</returns>
Public Function GetTableByOrdinal(ordinal As Long) As DataTable
    On Error Resume Next
    If TableCount() = 0 Then  ' Add this check
        Set GetTableByOrdinal = Nothing
        Exit Function
    End If
    On Error GoTo 0
    
    If ordinal >= LBound(tables) And ordinal <= UBound(tables) Then
        Set GetTableByOrdinal = tables(ordinal)
    Else
        Set GetTableByOrdinal = Nothing
    End If
End Function

''' <summary>
''' Gets the number of tables in the DataSet
''' </summary>
''' <returns>The count of tables</returns>
Public Function TableCount() As Long
    On Error Resume Next

    If UBound(tables) < LBound(tables) Then
        TableCount = 0
    Else
        TableCount = UBound(tables) - LBound(tables) + 1
    End If
    
    On Error GoTo 0
End Function

''' <summary>
''' Adds a DataTable to the DataSet. If a table with the same name exists,
''' attempts to rename tables with default names (e.g., "Table1")
''' </summary>
''' <param name="tbl">The DataTable to add</param>
''' <returns>The added DataTable reference, or Nothing if failed</returns>
''' <remarks>
''' Raises error 512 if tbl is Nothing
''' Raises error 513 if duplicate name exists (non-default name)
''' Raises error 514 if renaming fails to resolve conflict
''' </remarks>
Public Function AddTable(ByRef tbl As DataTable) As DataTable
    Dim newTable As DataTable
    Dim currentCount As Long
    Dim suffix As String
    Dim newIndex As Long

    If tbl Is Nothing Then
        ' Cannot add a Nothing table
        Err.Raise vbObjectError + 512, "DataSet.AddTable", "Cannot add a Nothing table to the DataSet."
        Set AddTable = Nothing
        
        Exit Function
    End If

    ' Make sure the table doesn't already contain an entry for this table name
    If Not GetTableByName(tbl.Name) Is Nothing Then
        ' Table with this name already exists
        ' If the table name starts with "Table" followed by a number, this is a default name 
        ' generated by the DataTable class. In this case, we can rename the table to avoid the conflict.
        If LCase(Left(tbl.Name, 5)) = "table" Then
            suffix = Mid(tbl.Name, 6)

            If IsNumeric(suffix) Then
                newIndex = 1

                Do While Not GetTableByName("Table" & newIndex) Is Nothing
                    newIndex = newIndex + 1
                Loop

                tbl.Name = "Table" & newIndex
            End If
        Else
            ' Otherwise, raise an error
            Err.Raise vbObjectError + 513, "DataSet.AddTable", "A table with the name '" & tbl.Name & "' already exists in the DataSet."
            Set AddTable = Nothing
            Exit Function
        End If

    End If

    ' Check table name again after renaming attempt, edge case
    If Not GetTableByName(tbl.Name) Is Nothing Then
        ' Still a conflict, cannot add
        Err.Raise vbObjectError + 514, "DataSet.AddTable", "A table with the name '" & tbl.Name & "' already exists in the DataSet after renaming attempt."
        Set AddTable = Nothing
        Exit Function
    End If

    'Get the current number of tables in the array
    currentCount = TableCount()

    ' Extend the array by one
    ReDim Preserve tables(0 To currentCount)

    Set tables(currentCount) = tbl

    Set AddTable = tbl
End Function

''' <summary>
''' Removes a table from the DataSet by its name
''' </summary>
''' <param name="tableName">The name of the table to remove</param>
''' <returns>True if removed successfully, False if not found</returns>
Public Function RemoveTableByName(tableName As String) As Boolean
    Dim i As Long
    Dim j As Long
    Dim currentCount As Long
    Dim tableOrdinal As Long

    currentCount = TableCount()
    
    ' Check for empty array
    If currentCount = 0 Then
        RemoveTableByName = False
        Exit Function
    End If
    
    tableOrdinal = GetDataTableOrdinalByName(tableName)

    If tableOrdinal = -1 Then
        ' Table not found
        RemoveTableByName = False
        Exit Function
    End If

    ' Shift remaining tables down in the array to fill the gap
    For j = tableOrdinal To UBound(tables) - 1
        Set tables(j) = tables(j + 1)
    Next j

    ' If found, resize the array to trim off the excess element at the end of the array.
    ReDim Preserve tables(0 To currentCount - 2)

    RemoveTableByName = True  ' Change this line
End Function

''' <summary>
''' Removes a table from the DataSet by its ordinal position
''' </summary>
''' <param name="ordinal">The zero-based index of the table to remove</param>
''' <returns>True if removed successfully, False if ordinal invalid</returns>
Public Function RemoveTableByOrdinal(ordinal As Long) As Boolean
    Dim i As Long
    Dim j As Long
    Dim found As Boolean
    Dim currentCount As Long

    found = False
    currentCount = TableCount()
    
    ' Check for empty array
    If currentCount = 0 Then
        RemoveTableByOrdinal = False
        Exit Function
    End If

    ' Check if the ordinal is valid
    If ordinal >= LBound(tables) And ordinal <= UBound(tables) Then
        found = True

        ' Shift remaining tables down in the array to fill the gap
        For j = ordinal To UBound(tables) - 1
            Set tables(j) = tables(j + 1)
        Next j

        ' Resize the array to trim off the excess element at the end of the array.
        ReDim Preserve tables(0 To currentCount - 2)
    End If

    RemoveTableByOrdinal = found
End Function

''' <summary>
''' Initializes the DataSet with an empty tables array
''' </summary>
Private Sub Class_Initialize()
    ReDim tables(0 To -1)
End Sub

''' <summary>
''' Cleans up object references when the DataSet is destroyed
''' </summary>
Private Sub Class_Terminate()
    Dim i As Long

    ' Clean up each table
    On Error Resume Next

    If UBound(tables) >= LBound(tables) Then
        For i = LBound(tables) To UBound(tables)
            Set tables(i) = Nothing
        Next i
    End If
    
    On Error GoTo 0

    ReDim tables(0 To -1)
End Sub