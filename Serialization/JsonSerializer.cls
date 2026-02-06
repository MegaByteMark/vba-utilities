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
' Features:
'   - Custom date formatting via DateFormat property (default: "yyyy-MM-dd" ISO 8601)
'
' Usage:
'   Dim serializer As New JsonSerializer
'   serializer.DateFormat = "yyyy-MM-dd HH:mm:ss"
'   Dim json As String
'   json = serializer.Serialize(myDataTable)
'   
'   Dim deserialized As DataTable
'   Set deserialized = serializer.Deserialize(json)
'
' Author: MegaByteMark
' Version: 2.1
'========================================

Option Explicit

Private dateFormat As String

''' <summary>
''' Sets the date format string for serialization and deserialization.
''' Uses VBA Format function syntax (e.g., "yyyy-MM-dd", "dd/MM/yyyy HH:mm:ss")
''' Defaults to ISO 8601 format: "yyyy-MM-dd"
''' </summary>
''' <param name="value">The date format string</param>
Public Property Let DateFormat(value As String)
    dateFormat = value
End Property

''' <summary>
''' Gets the date format string for serialization and deserialization.
''' </summary>
''' <return>The date format string</return>
Public Property Get DateFormat() As String
    If dateFormat = "" Then
        DateFormat = "yyyy-MM-dd"
    Else
        DateFormat = dateFormat
    End If
End Property

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
        json = json & """dataType"": """ & EscapeJson(VbTypeToString(column.DataType)) & """" & vbCrLf
        json = json & "}"

        If i < dataTable.ColumnCount() - 1 Then
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
    Dim parts() as String
    Dim partCount As Long

    ReDim parts(0 To dataTable.RowCount() * (dataTable.ColumnCount() + 2))
    partCount = 0
    parts(partCount) = "["
    partCount = partCount + 1

    For i = 0 To dataTable.RowCount() - 1
        Set row = dataTable.Rows(i)

        parts(partCount) = vbCrLf & "  {"

        For j = 0 To dataTable.ColumnCount() - 1
            Set cell = row.GetValue(j)
            Set column = dataTable.Columns(j)

            parts(partCount) = parts(partCount) & """" & EscapeJson(column.Name) & """: "
            partCount = partCount + 1

            If IsNull(cell) Or IsEmpty(cell) Then
                parts(partCount) = parts(partCount) & "null"
                partCount = partCount + 1
            Else
                Select Case column.DataType
                    Case vbBoolean
                        If SafeBooleanConvert(cell) Then
                            parts(partCount) = "true"
                            partCount = partCount + 1
                        Else
                            parts(partCount) = "false"
                            partCount = partCount + 1
                        End If
                    Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
                        parts(partCount) = CStr(cell)
                        partCount = partCount + 1
                    Case vbDate
                        parts(partCount) = """" & EscapeJson(Format(CDate(cell), DateFormat)) & """"
                        partCount = partCount + 1
                    Case Else
                        parts(partCount) = """" & EscapeJson(CStr(cell)) & """"
                        partCount = partCount + 1
                End Select
            End If

            If j < dataTable.ColumnCount() - 1 Then
                parts(partCount) = ", "
                partCount = partCount + 1
            End If
        Next j

        parts(partCount) = vbCrLf & "  }"
        partCount = partCount + 1

        If i < dataTable.RowCount() - 1 Then
            parts(partCount) = ","
            partCount = partCount + 1
        End If
    Next i

    parts(partCount) = vbCrLf & "]"
    partCount = partCount + 1

    json = Join(parts, "", 0, partCount)

    SerializeRows = json
End Function

''' <summary>
''' Deserializes a JSON string into a DataTable
''' </summary>
''' <param name="json">The JSON string to deserialize</param>
''' <returns>A DataTable represented by the JSON string</returns>
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

    If InStr(json, """rows""") = 0 Then
        Err.Raise 517, "JsonSerializer.Deserialize", "JSON string must be valid JSON format."
    End If

    If InStr(json, """columns""") = 0 Then
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
        nameStart = InStr(nameStart, json, """")

        If nameStart = 0 Then
            Err.Raise 517, "JsonSerializer.Deserialize", "Invalid JSON: malformed name property"
        End If
        
        nameEnd = InStr(nameStart + 1, json, """")

        If nameEnd = 0 Then
            Err.Raise 517, "JsonSerializer.Deserialize", "Invalid JSON: malformed name property"
        End If

        tableName = Mid(json, nameStart + 1, nameEnd - nameStart - 1)
        dataTable.Name = tableName
    Else
        dataTable.Name = "Table1"
    End If

    ' Process the columns next
    columnsStart = InStr(json, """columns""")

    If columnsStart > 0 Then
        columnsStart = InStr(columnsStart, json, "[")

        If columnsStart = 0 Then
            Err.Raise 517, "JsonSerializer.Deserialize", "Invalid JSON: malformed columns property."
        End If

        columnsStart = columnsStart + 1
        columnsEnd = InStr(columnsStart, json, "]")

        If columnsEnd = 0 Then
            Err.Raise 517, "JsonSerializer.Deserialize", "Invalid JSON: malformed columns property."
        End If

        columnsJson = Mid(json, columnsStart, columnsEnd - columnsStart)
    Else
        Err.Raise 517, "JsonSerializer.Deserialize", "Invalid JSON: malformed columns property."
    End If

    Set columns = DeserializeColumns(columnsJson)

    For Each Dim col As DataColumn In columns
        dataTable.AddColumn col
    Next col

    ' Process the rows next
    rowsStart = InStr(json, """rows""")

    If rowsStart > 0 Then
        rowsStart = InStr(rowsStart, json, "[")

        If rowsStart = 0 Then
            Err.Raise 517, "JsonSerializer.Deserialize", "Invalid JSON: malformed rows property."
        End If

        rowsStart = rowsStart + 1
        rowsEnd = InStr(rowsStart, json, "]")

        If rowsEnd = 0 Then
            Err.Raise 517, "JsonSerializer.Deserialize", "Invalid JSON: malformed rows property."
        End If

        rowsJson = Mid(json, rowsStart, rowsEnd - rowsStart)
    Else
        Err.Raise 517, "JsonSerializer.Deserialize", "Invalid JSON: malformed rows property."
    End If

    Set rows = DeserializeRows(rowsJson, dataTable)

    For Each Dim rw As DataRow In rows
        dataTable.AddRow rw
    Next rw

    Set Deserialize = dataTable

End Function

'''<summary>
''' Deserializes JSON string into an array of DataColumn objects
'''</summary>
'''<param name="columnsJson">The JSON string representing the columns</param>
'''<returns>An array of DataColumn objects</returns>
Private Function DeserializeColumns(columnsJson As String) As DataColumn()
    Dim columns() As DataColumn
    Dim objects() As String
    Dim i As Long, j As Long
    Dim char As String
    Dim inQuote As Boolean
    Dim braceDepth As Integer
    Dim currentObject As String
    
    i = 1
    braceDepth = 0
    inQuote = False
    currentObject = ""
    
    Do While i <= Len(columnsJson)
        char = Mid(columnsJson, i, 1)
        
        If char = """" Then 
            inQuote = Not inQuote
        End If
        
        If Not inQuote Then
            If char = "{" Then
                braceDepth = braceDepth + 1
                currentObject = currentObject & char
            ElseIf char = "}" Then
                currentObject = currentObject & char
                braceDepth = braceDepth - 1
                
                If braceDepth = 0 Then
                    If IsEmpty(objects) Then
                        ReDim objects(0)
                    Else
                        ReDim Preserve objects(0 To UBound(objects) + 1)
                    End If

                    objects(UBound(objects)) = currentObject
                    currentObject = ""
                End If
            ElseIf braceDepth > 0 Then
                currentObject = currentObject & char
            End If
        ElseIf braceDepth > 0 Then
            currentObject = currentObject & char
        End If
        
        i = i + 1
    Loop
    
    If Not IsEmpty(objects) Then
        ReDim columns(0 To UBound(objects))
        
        For i = 0 To UBound(objects)
            Set columns(i) = DeserializeColumn(objects(i))
        Next i
    Else
        ReDim columns(-1)
    End If
    
    Set DeserializeColumns = columns
End Function

'''<summary>
''' Deserializes a single JSON column object into a DataColumn
'''</summary>
'''<param name="objectJson">The JSON string representing the column</param>
'''<returns>A DataColumn object</returns>
Private Function DeserializeColumn(objectJson As String) As DataColumn
    Dim col As DataColumn
    Dim nameStart As Long, nameEnd As Long
    Dim typeStart As Long, typeEnd As Long
    
    Set col = New DataColumn
    
    nameStart = InStr(objectJson, """name"":")

    If nameStart > 0 Then
        nameStart = InStr(nameStart, objectJson, """")

        If nameStart = 0 Then
            Err.Raise 517, "JsonSerializer.DeserializeColumn", _
                "Invalid column JSON: malformed name property"
        End If

        nameStart = nameStart + 1
        nameEnd = InStr(nameStart, objectJson, """")

        If nameEnd = 0 Then
            Err.Raise 517, "JsonSerializer.DeserializeColumn", _
                "Invalid column JSON: malformed name property"
        End If

        col.Name = Trim(Mid(objectJson, nameStart, nameEnd - nameStart))
    Else
        Err.Raise 517, "JsonSerializer.DeserializeColumn", _
            "Invalid column JSON: missing 'name' property"
    End If
    
    typeStart = InStr(objectJson, """dataType"":")

    If typeStart > 0 Then
        typeStart = InStr(typeStart, objectJson, """")
        
        If typeStart = 0 Then
            Err.Raise 517, "JsonSerializer.DeserializeColumn", _
                "Invalid column JSON: malformed dataType property"
        End If

        typeStart = typeStart + 1
        typeEnd = InStr(typeStart, objectJson, """")

        If typeEnd = 0 Then
            Err.Raise 517, "JsonSerializer.DeserializeColumn", _
                "Invalid column JSON: malformed dataType property"
        End If
        
        col.DataType = StringToVbType(Trim(Mid(objectJson, typeStart, typeEnd - typeStart)))
    Else
        Err.Raise 517, "JsonSerializer.DeserializeColumn", _
            "Invalid column JSON: missing 'dataType' property"
    End If
    
    Set DeserializeColumn = col
End Function

'''<summary>
''' Deserializes JSON string of DataRow objects, comma separated into an array of DataRow objects
'''</summary>
'''<param name="rowsJson">The JSON string representing the rows</param>
'''<param name="dataTable">The DataTable to which the rows belong</param>
'''<returns>An array of DataRow objects</returns>
Private Function DeserializeRows(rowsJson As String, dataTable As DataTable) As DataRow()
    Dim rows() As DataRow
    Dim objects() As String
    Dim inObject As Boolean
    Dim inQuotedString As Boolean
    Dim i As Long
    Dim char As String
    Dim currentObject As String
    Dim objectNesting As Integer
    Dim braceBalance As Integer 
    Dim row As DataRow

    If rowsJson = "" Or IsNull(rowsJson) Then
        ReDim rows(-1)
        Set DeserializeRows = rows
        
        Exit Function
    End If

    inObject = False
    inQuotedString = False
    currentObject = ""
    i = 1
    objectNesting = 0
    braceBalance = 0

    Do While i <= Len(rowsJson)
        char = Mid(rowsJson, i, 1)

        ' Handle quotes
        If char = """" Then
            inQuotedString = Not inQuotedString
            If inObject Then currentObject = currentObject & char
        End If

        If Not inQuotedString Then
            If char = "{" Then
                braceBalance = braceBalance + 1
                inObject = True
                currentObject = currentObject & char
                objectNesting = objectNesting + 1
                
            ElseIf char = "}" Then
                braceBalance = braceBalance - 1
                
                ' Check for negative balance
                If braceBalance < 0 Then
                    Err.Raise 517, "JsonSerializer.DeserializeRows", _
                        "Invalid JSON: Unexpected closing brace at position " & i
                End If
                
                currentObject = currentObject & char
                objectNesting = objectNesting - 1

                If objectNesting = 0 Then
                    inObject = False

                    ' Completed an object, add to objects array
                    If IsEmpty(objects) Then
                        ReDim objects(0)
                    Else
                        ReDim Preserve objects(0 To UBound(objects) + 1)
                    End If

                    objects(UBound(objects)) = currentObject
                    currentObject = ""
                End If
            Else
                If inObject Then
                    currentObject = currentObject & char
                End If
            End If
        End If

        i = i + 1
    Loop

    ' Validate final balance
    If braceBalance <> 0 Then
        Err.Raise 517, "JsonSerializer.DeserializeRows", _
            "Invalid JSON: Unclosed braces (missing " & Abs(braceBalance) & " closing brace(s))"
    End If

    ' Now parse each object into DataRow
    If Not IsEmpty(objects) Then
        If UBound(objects) < 0 Then
            ReDim rows(-1)
        Else
            ReDim rows(0 To UBound(objects))
            For i = 0 To UBound(objects)
                Set rows(i) = DeserializeRow(objects(i), dataTable)
            Next i
        End If
    Else
        ReDim rows(-1)
    End If

    DeserializeRows = rows
End Function

'''<summary>
''' Deserializes a single JSON row object into a DataRow
'''</summary>
'''<param name="rowJson">The JSON string representing the row</param>
'''<param name="dataTable">The DataTable to which the row belongs</param>
'''<returns>A DataRow object</returns>
Private Function DeserializeRow(rowJson As String, dataTable As DataTable) As DataRow    
    Dim row As DataRow
    Dim col As DataColumn
    Dim j As Long, k As Long
    Dim currentKey As String
    Dim currentValue As String
    Dim char As String
    Dim colIndex As Long
    Dim inKey As Boolean
    Dim inValue As Boolean
    Dim inValueQuoted As Boolean
    Dim backslashCount As Long

    Set row = New DataRow
    inKey = False
    inValue = False
    inValueQuoted = False
    currentKey = ""
    currentValue = ""
    colIndex = -1

    For j = 1 To Len(rowJson)
        char = Mid(rowJson, j, 1)

        ' Count consecutive backslashes before this quote
        If char = """" Then
            backslashCount = 0
            k = j - 1

            Do While k >= 1 And Mid(rowJson, k, 1) = "\"
                backslashCount = backslashCount + 1
                k = k - 1
            Loop
        End If
        
        ' Odd number of backslashes = escaped quote
        Dim isEscapedQuote As Boolean
        isEscapedQuote = (char = """" And (backslashCount Mod 2 = 1))

        If char = """" And Not isEscapedQuote Then
            If inValue Then
                ' This is a string delimiter (open or close)
                inValueQuoted = Not inValueQuoted
                currentValue = currentValue & char
            Else
                inKey = Not inKey

                If inKey Then
                    currentKey = ""
                Else
                    colIndex = dataTable.GetColumnOrdinalByName(Replace(currentKey, """", ""))

                    If colIndex = -1 Then
                        Err.Raise 518, "JsonSerializer.DeserializeRow", _
                            "Column '" & Replace(currentKey, """", "") & "' not found in DataTable."
                    End If
                End If
            End If
        ElseIf char = ":" And Not inKey And Not inValueQuoted Then
            inValue = True
            currentValue = ""
        ElseIf (char = "," Or char = "}") And Not inKey And Not inValueQuoted Then
            If colIndex >= 0 Then
                Set col = dataTable.Columns(colIndex)
                
                ' Remove whitespace
                currentValue = Trim(currentValue)

                If LCase(currentValue) = "null" Or currentValue = "" Then
                    row.SetValue colIndex, Null
                Else
                    Select Case col.DataType
                        Case vbBoolean
                            If LCase(currentValue) = "true" Then
                                row.SetValue colIndex, True
                            ElseIf LCase(currentValue) = "false" Then
                                row.SetValue colIndex, False
                            Else
                                Err.Raise 518, "JsonSerializer.DeserializeRow", _
                                    "Invalid boolean value '" & currentValue & "' for column '" & col.Name & "'."
                            End If
                        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
                            ' Strip quotes if present
                            If Left(currentValue, 1) = """" Then
                                currentValue = Mid(currentValue, 2)
                            End If

                            If Right(currentValue, 1) = """" Then
                                currentValue = Left(currentValue, Len(currentValue) - 1)
                            End If

                            row.SetValue colIndex, CDbl(currentValue)
                        Case vbDate
                            ' Strip quotes if present
                            If Left(currentValue, 1) = """" Then
                                currentValue = Mid(currentValue, 2)
                            End If

                            If Right(currentValue, 1) = """" Then
                                currentValue = Left(currentValue, Len(currentValue) - 1)
                            End If

                            ' Unescape any escaped quotes in the date string
                            currentValue = Replace(currentValue, "\""", """")
                            
                            row.SetValue colIndex, SafeDateConvert(currentValue)
                        Case Else
                            ' Trim outer quotes
                            If Left(currentValue, 1) = """" Then
                                currentValue = Mid(currentValue, 2)
                            End If

                            If Right(currentValue, 1) = """" Then
                                currentValue = Left(currentValue, Len(currentValue) - 1)
                            End If

                            ' Unescape
                            currentValue = Replace(currentValue, "\""", """")
                            currentValue = Replace(currentValue, "\\", "\")

                            row.SetValue colIndex, currentValue
                    End Select
                End If
            End If

            currentKey = ""
            currentValue = ""
            inValue = False
        ElseIf inKey Then
            currentKey = currentKey & char
        ElseIf inValue Then
            currentValue = currentValue & char
        End If
    Next j

    Set DeserializeRow = row
    
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
''' Converts a VB data type to its string representation
''' </summary>
''' <param name="vbType">The VB data type</param>
''' <returns>The string representation of the data type</returns>
Private Function VbTypeToString(vbType As VbVarType) As String
    'Reference: https://learn.microsoft.com/en-us/office/vba/language/concepts/getting-started/vartype-constants

    Select Case vbType
        Case vbEmpty: VbTypeToString = "object"
        Case vbNull: VbTypeToString = "object"
        Case vbInteger: VbTypeToString = "int"
        Case vbLong: VbTypeToString = "long"
        Case vbSingle: VbTypeToString = "decimal"
        Case vbDouble: VbTypeToString = "decimal"
        Case vbCurrency: VbTypeToString = "decimal"
        Case vbDate: VbTypeToString = "date"
        Case vbString: VbTypeToString = "string"
        Case vbObject: VbTypeToString = "object"
        Case vbBoolean: VbTypeToString = "boolean"
        Case vbVariant: VbTypeToString = "object"
        Case vbDecimal: VbTypeToString = "decimal"
        Case vbByte: VbTypeToString = "int"
        Case vbLongLong: VbTypeToString = "long"
        Case vbUserDefinedType: VbTypeToString = "object"
        Case vbArray: VbTypeToString = "object"
        Case Else: VbTypeToString = "object"
    End Select
End Function

'''<summary>
''' Converts a string representation of a data type to its VB data type
''' </summary>
''' <param name="typeString">The string representation of the data type</param>
''' <returns>The VB data type</returns>
Private Function StringToVbType(typeString As String) As VbVarType
    Select Case LCase(typeString)
        Case "object": StringToVbType = vbVariant
        Case "int": StringToVbType = vbInteger
        Case "long": StringToVbType = vbLong
        Case "decimal": StringToVbType = vbDouble
        Case "date": StringToVbType = vbDate
        Case "string": StringToVbType = vbString
        Case "boolean": StringToVbType = vbBoolean
        Case Else: StringToVbType = vbVariant
    End Select
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
''' Parses a date string according to the custom date format (locale-independent)
'''</summary>
'''<param name="dateString">The string to parse</param>
'''<param name="format">The date format string</param>
'''<returns>A Date value</returns>
Private Function ParseDateByFormat(dateString As String, format As String) As Variant
    Dim year As Long, month As Long, day As Long
    Dim hour As Long, minute As Long, second As Long
    Dim formatPos As Long, datePos As Long
    Dim currentDateValue As String
    
    ' Initialize defaults
    year = 1900
    month = 1
    day = 1
    hour = 0
    minute = 0
    second = 0
    
    On Error GoTo ErrorHandler
    
    ' Process format string character by character, matching date string
    formatPos = 1
    datePos = 1
    
    Do While formatPos <= Len(format) And datePos <= Len(dateString)
        ' Check for 4-digit year pattern
        If Mid(format, formatPos, 4) = "yyyy" Then
            If datePos + 3 > Len(dateString) Then GoTo ErrorHandler
            currentDateValue = Mid(dateString, datePos, 4)
            year = CLng(currentDateValue)
            formatPos = formatPos + 4
            datePos = datePos + 4
        
        ' Check for 2-digit month pattern
        ElseIf Mid(format, formatPos, 2) = "MM" Then
            If datePos + 1 > Len(dateString) Then GoTo ErrorHandler
            currentDateValue = Mid(dateString, datePos, 2)
            month = CLng(currentDateValue)
            formatPos = formatPos + 2
            datePos = datePos + 2
        
        ' Check for 2-digit day pattern
        ElseIf Mid(format, formatPos, 2) = "dd" Then
            If datePos + 1 > Len(dateString) Then GoTo ErrorHandler
            currentDateValue = Mid(dateString, datePos, 2)
            day = CLng(currentDateValue)
            formatPos = formatPos + 2
            datePos = datePos + 2
        
        ' Check for 2-digit hour pattern (24-hour)
        ElseIf Mid(format, formatPos, 2) = "HH" Then
            If datePos + 1 > Len(dateString) Then GoTo ErrorHandler
            currentDateValue = Mid(dateString, datePos, 2)
            hour = CLng(currentDateValue)
            formatPos = formatPos + 2
            datePos = datePos + 2
        
        ' Check for 2-digit hour pattern (12-hour)
        ElseIf Mid(format, formatPos, 2) = "hh" Then
            If datePos + 1 > Len(dateString) Then GoTo ErrorHandler
            currentDateValue = Mid(dateString, datePos, 2)
            hour = CLng(currentDateValue)
            formatPos = formatPos + 2
            datePos = datePos + 2
        
        ' Check for 2-digit minute pattern
        ElseIf Mid(format, formatPos, 2) = "mm" Then
            If datePos + 1 > Len(dateString) Then GoTo ErrorHandler
            currentDateValue = Mid(dateString, datePos, 2)
            minute = CLng(currentDateValue)
            formatPos = formatPos + 2
            datePos = datePos + 2
        
        ' Check for 2-digit second pattern
        ElseIf Mid(format, formatPos, 2) = "ss" Then
            If datePos + 1 > Len(dateString) Then GoTo ErrorHandler
            currentDateValue = Mid(dateString, datePos, 2)
            second = CLng(currentDateValue)
            formatPos = formatPos + 2
            datePos = datePos + 2
        
        ' Non-pattern character (separator like -, /, :, space), skip in both
        Else
            formatPos = formatPos + 1
            datePos = datePos + 1
        End If
    Loop
    
    ' Validate extracted values
    If month < 1 Or month > 12 Then GoTo ErrorHandler
    If day < 1 Or day > 31 Then GoTo ErrorHandler
    If hour < 0 Or hour > 23 Then GoTo ErrorHandler
    If minute < 0 Or minute > 59 Then GoTo ErrorHandler
    If second < 0 Or second > 59 Then GoTo ErrorHandler
    
    ' Create the date/time value
    ParseDateByFormat = DateSerial(year, month, day) + TimeSerial(hour, minute, second)
    Exit Function
    
ErrorHandler:
    Err.Raise 518, "JsonSerializer.ParseDateByFormat", _
        "Unable to parse date string '" & dateString & "' using format '" & format & "'."
    ParseDateByFormat = Null
End Function

'''<summary>
''' Safely converts a string to a date using the configured date format (locale-independent)
'''</summary>
'''<param name="dateString">The string to convert to a date</param>
'''<returns>A Date value</returns>
Private Function SafeDateConvert(dateString As String) As Variant
    On Error Resume Next
    
    Dim result As Variant
    result = ParseDateByFormat(dateString, DateFormat)
    
    If Err.Number <> 0 Then
        Err.Raise 518, "JsonSerializer.SafeDateConvert", _
            "Unable to convert '" & dateString & "' to date using format '" & DateFormat & "'."
        SafeDateConvert = Null
    Else
        SafeDateConvert = result
    End If
    
    On Error GoTo 0
End Function

'''<summary>
''' Flattens a JSON string by removing line breaks and extra spaces
'''</summary>
'''<param name="json">The JSON string to flatten</param>
'''<returns>The flattened JSON string</returns>
Private Function FlattenJsonString(json As String) As String
    ' Helper function to remove line breaks and extra spaces from JSON string
    json = Replace(json, vbCrLf, "")
    json = Replace(json, vbLf, "")
    json = Replace(json, vbCr, "")
    json = Trim(json)

    FlattenJsonString = json
End Function