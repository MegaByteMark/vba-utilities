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
' JsonSerializer Class
'========================================
' Provides serialization and deserialization of DataTable objects to/from JSON format.
'
' Dependencies:
'   - DataTable.cls (Data/Common/DataTable.cls)
'   - DataRow.cls (Data/Common/DataRow.cls)
'   - DataColumn.cls (Data/Common/DataColumn.cls)
'   - DataCell.cls (Data/Common/DataCell.cls)
'
' Error Codes:
'   513 - DataTable cannot be Nothing
'   514 - DataTable must have at least one column
'   515 - JSON string cannot be Nothing
'   516 - JSON string cannot be empty
'   517 - JSON string must contain at least one comma
'   518 - JSON row field count mismatch
'
' Usage:
'   Dim serializer As New JsonSerializer
'   serializer.IncludeHeaders = True
'   Dim json As String
'   json = serializer.Serialize(myDataTable)
'
' Author: MegaByteMark
' Version: 2.0
'========================================

Option Explicit

''' <summary>
''' Serializes a DataTable into a JSON string
''' </summary>
''' <param name="dataTable">The DataTable to serialize</param>
''' <returns>A JSON string representing the DataTable</returns>
Public Function Serialize(dataTable As DataTable) As String        
    Dim json As String
    Dim i As Long, j As Long
    Dim row As DataRow
    Dim cell As DataCell
    Dim columnCount As Long

    ' Error Handling
    If dataTable Is Nothing Then
        Err.Raise 513, "JsonSerializer.Serialize", "DataTable cannot be Nothing."
    End If
    
    If dataTable.Columns.Count = 0 Then
        Err.Raise 514, "JsonSerializer.Serialize", "DataTable must have at least one column."
    End If

    If dataTable.Rows.Count > 0 Then
        columnCount = dataTable.ColumnCount()

        ' Validate row field counts
        For i = 0 To dataTable.Rows.Count - 1
            Set row = dataTable.Rows(i)

            If row.Cells.Count <> columnCount Then
                Err.Raise 518, "JsonSerializer.Serialize", "JSON row field count mismatch, row " & CStr(i) & "."
                Exit Function
            End If
        Next i
    End If
    
    json = "{" & vbCrLf
    json = json & " ""name"": """ & EscapeJson(dataTable.Name) & """" & vbCrLf
    json = json & ", ""columns"": " & SerializeColumns(dataTable) & vbCrLf
    json = json & ", ""rows"": " & SerializeRows(dataTable) & vbCrLf
    json = json & "}"

    Serialize = json
End Function

Public Function Deserialize(json As String) As DataTable
    Dim dataTable As DataTable
    Dim nameStart As Long
    Dim nameEnd As Long
    Dim tableName As String
    Dim columnsStart As Long
    Dim columnsEnd As Long
    Dim rowsStart As Long
    Dim rowsEnd As Long
    Dim columnsJson As String
    Dim rowsJson As String
    Dim columns() As DataColumn
    Dim rows() As DataRow

    If IsNull(json) Then
        Err.Raise 515, "JsonSerializer.Deserialize", "JSON string cannot be Nothing."
    End If

    If Trim(json) = "" Then
        Err.Raise 516, "JsonSerializer.Deserialize", "JSON string cannot be empty."
    End If

    If InStr(json, "{") = 0 Then
        Err.Raise 517, "JsonSerializer.Deserialize", "JSON string must be valid JSON format."
    End If

    If InStr(json, "rows:") = 0 Then
        Err.Raise 517, "JsonSerializer.Deserialize", "JSON string must be valid JSON format."
    End If

    If InStr(json, "columns:") = 0 Then
        Err.Raise 517, "JsonSerializer.Deserialize", "JSON string must be valid JSON format."
    End If

    ' Flatten string to remove any line breaks or extra spaces
    json = FlattenJsonString(json)

    'Has to start with an opening brace
    If Left(json, 1) <> "{" Then
        Err.Raise 517, "JsonSerializer.Deserialize", "JSON string must be valid JSON format."
    End If

    Set dataTable = New DataTable

    ' Chop off the opening brace
    json = Mid(json, InStr(json, "{") + 1)

    ' Process the datatable name first, if the name is not provided, default to "Table1"
    nameStart = InStr(json, """name"":")

    If nameStart > 0 Then
        nameStart = InStr(nameStart, json, """") + 1
        nameEnd = InStr(nameStart, json, """")
        tableName = Mid(json, nameStart, nameEnd - nameStart)
        dataTable.Name = tableName
    Else
        dataTable.Name = "Table1"
    End If

    ' Process the columns next
    columnsStart = InStr(json, """columns"": [")

    If columnsStart > 0 Then
        columnsStart = InStr(columnsStart, json, "[") + 1
        columnsEnd = InStr(columnsStart, json, "]")
        columnsJson = Mid(json, columnsStart, columnsEnd - columnsStart)
    Else
        Err.Raise 517, "JsonSerializer.Deserialize", "JSON string must be valid JSON format."
    End If

    Set columns = DeserializeColumns(columnsJson)

    For Each Dim col As DataColumn In columns
        dataTable.AddColumn col
    Next col

    ' Process the rows next
    rowsStart = InStr(json, """rows"": [")

    If rowsStart > 0 Then
        rowsStart = InStr(rowsStart, json, "[") + 1
        rowsEnd = InStr(rowsStart, json, "]")
        rowsJson = Mid(json, rowsStart, rowsEnd - rowsStart)
    Else
        Err.Raise 517, "JsonSerializer.Deserialize", "JSON string must be valid JSON format."
    End If

    Set rows = DeserializeRows(rowsJson, dataTable)

    For Each Dim rw As DataRow In rows
        dataTable.AddRow rw
    Next rw

    Set Deserialize = dataTable

End Function

'========================================
' Json Parsing Helpers
'========================================
'''<summary>
''' Escapes special characters in a string for JSON formatting
'''</summary>
'''<param name="value">The string to escape</param>
'''<returns>The escaped string</returns>
Private Function EscapeJson(value As String) As String
    If IsNull(value) Then
        EscapeJson = ""
        Exit Function
    End If
    
    value = Replace(value, "\", "\\")
    value = Replace(value, """", "\""")
    value = Replace(value, vbCrLf, "\n")
    value = Replace(value, vbLf, "\n")
    value = Replace(value, vbCr, "\r")

    EscapeJson = value
End Function

'''<summary>
''' Serializes the columns of a DataTable into JSON format
'''</summary>
'''<param name="dataTable">The DataTable containing the columns to serialize</param>
'''<returns>A JSON string representing the columns</returns>
Private Function SerializeColumns(dataTable As DataTable) As String
    Dim json As String
    Dim i As Long
    Dim column As DataColumn

    json = "["

    For i = 0 To dataTable.ColumnCount() - 1
        Set column = dataTable.Columns(i)

        json = json & vbCrLf & "  {"
        json = json & """name"": """ & EscapeJson(column.Name) & """," & vbCrLf
        json = json & """dataType"": """ & EscapeJson(column.DataType) & """" & vbCrLf
        json = json & "}"

        If i < dataTable.Columns.Count - 1 Then
            json = json & ","
        End If
    Next i

    json = json & vbCrLf & "]"

    SerializeColumns = json
End Function

'''<summary>
''' Serializes the rows of a DataTable into JSON format
'''</summary>
'''<param name="dataTable">The DataTable containing the rows to serialize</param>
'''<returns>A JSON string representing the rows</returns>
Private Function SerializeRows(dataTable As DataTable) As String
    Dim json As String
    Dim i As Long
    Dim j As Long
    Dim row As DataRow
    Dim cell As DataCell
    Dim column As DataColumn

    json = "["

    For i = 0 To dataTable.RowCount() - 1
        Set row = dataTable.Rows(i)

        json = json & vbCrLf & "  {"

        For j = 0 To dataTable.ColumnCount() - 1
            Set cell = row.GetValue(j)
            Set column = dataTable.Columns(j)

            json = json & """" & EscapeJson(column.Name) & """: "

            If IsNull(cell) Or IsEmpty(cell) Then
                json = json & "null"
            Else
                Select Case column.DataType
                    Case vbBoolean
                        If SafeBooleanConvert(cell) Then
                            json = json & "true"
                        Else
                            json = json & "false"
                        End If
                    Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
                        json = json & CStr(cell)
                    Case Else
                        json = json & """" & EscapeJson(CStr(cell)) & """"
                End Select
            End If

            If j < dataTable.ColumnCount() - 1 Then
                json = json & ", "
            End If
        Next j

        json = json & "}"

        If i < dataTable.RowCount() - 1 Then
            json = json & ","
        End If
    Next i

    json = json & vbCrLf & "]"

    SerializeRows = json
End Function

'''<summary>
''' Safely converts a variant to boolean, handling errors gracefully
'''</summary>
'''<param name="value">The variant value to convert</param>
'''<returns>Boolean representation of the value, or False on error</returns>
Private Function SafeBooleanConvert(value As Variant) As Boolean
    On Error Resume Next

    SafeBooleanConvert = CBool(value)

    If Err.Number <> 0 Then
        SafeBooleanConvert = False
        Err.Clear
    End If

    On Error GoTo 0
End Function

'''<summary>
''' Deserializes JSON string into an array of DataColumn objects
'''</summary>
'''<param name="columnsJson">The JSON string representing the columns</param>
'''<returns>An array of DataColumn objects</returns>
Private Function DeserializeColumns(columnsJson As String) As DataColumn()
    ' Columns Json:
    ' [
    '   {
    '       "name": "ColumnName",
    '       "dataType": "vbString"
    '   },
    ' ...
    ' ]

    Dim columns() As DataColumn
    Dim inColumnsArray As Boolean
    Dim inColumnObject As Boolean
    Dim inQuotedString As Boolean
    Dim currentColumn As DataColumn
    Dim i As Long
    Dim char As String
    Dim propName As String
    Dim propValue As String
    Dim propStart As Long
    Dim propEnd As Long
    Dim jsonBlock As String

    If Trim(columnsJson) = "" Then
        ' Return empty array if no columns
        ReDim columns(-1)
        Set DeserializeColumns = columns
        Exit Function
    End If

    'Flatten string to remove any line breaks or extra spaces
    columnsJson = FlattenJsonString(columnsJson)

    '  Must start with opening bracket
    If Left(columnsJson, 1) <> "[" Then
        Err.Raise 517, "JsonSerializer.DeserializeColumns", "JSON string must be valid JSON format."
    End If

    i = 1
    inColumnsArray = False
    inColumnObject = False
    inQuotedString = False
    Set currentColumn = Nothing
    jsonBlock = ""

    Do While i <= Len(columnsJson)
        char = Mid(columnsJson, i, 1)

        If char = "[" And Not inQuotedString Then
            inColumnsArray = True
        ElseIf char = "]" And Not inQuotedString Then
            inColumnsArray = False
        ElseIf char = "{" And Not inQuotedString Then
            inColumnObject = True
            Set currentColumn = New DataColumn
        ElseIf char = "}" And Not inQuotedString Then
            inColumnObject = False

            ' Add the completed column to the array
            If Not currentColumn Is Nothing Then
                If IsEmpty(columns) Then
                    ReDim columns(0)
                Else
                    ReDim Preserve columns(0 To UBound(columns) + 1)
                End If

                Set columns(UBound(columns)) = currentColumn
                Set currentColumn = Nothing
            End If
        ElseIf char = """" Then
            inQuotedString = Not inQuotedString
        Else 
            ' Add char to jsonBlock if needed
            If inColumnObject Then
                jsonBlock = jsonBlock & char
            End If
        End If
        
        If inColumnsArray And Not inColumnObject And Not inQuotedString And jsonBlock <> "" Then
            ' We've grabbed a column object block, now parse its properties
            If InStr(jsonBlock, """name""") Then
                propStart = InStr(jsonBlock, """name""") + 1
                propEnd = InStr(propStart, jsonBlock, ",")

                If propEnd = 0 Then 
                    propEnd = Len(jsonBlock) + 1
                End If
                
                propValue = Trim(Mid(columnsJson, propStart, propEnd - propStart))
                propValue = Replace(propValue, """", "")
                
                currentColumn.Name = propValue
                i = propEnd - 1
            ElseIf InStr(jsonBlock, """dataType""") Then
                propStart = InStr(jsonBlock, """dataType""") + 1
                propEnd = InStr(propStart, jsonBlock, ",")

                If propEnd = 0 Then 
                    propEnd = InStr(propStart, jsonBlock, "}")
                End If
                
                propValue = Trim(Mid(columnsJson, propStart, propEnd - propStart))
                propValue = Replace(propValue, """", "")    
                
                currentColumn.DataType = CLng(propValue)
                i = propEnd - 1
            End If
        End If

        i = i + 1
    Loop

    Set DeserializeColumns = columns
End Function

Private Function DeserializeRows(rowsJson As String, dataTable As DataTable) As DataRow()
    ' --- IGNORE ---
    ' WIP: Implementation of DeserializeRows function
    ' --- IGNORE ---
End Function

Private Function FlattenJsonString(json As String) As String
    ' Helper function to remove line breaks and extra spaces from JSON string
    json = Replace(json, vbCrLf, "")
    json = Replace(json, vbLf, "")
    json = Replace(json, vbCr, "")
    json = Trim(json)

    FlattenJsonString = json
End Function