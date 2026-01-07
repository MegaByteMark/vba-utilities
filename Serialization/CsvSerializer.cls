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
' CsvSerializer Class
'========================================
' Provides serialization and deserialization of DataTable objects to/from CSV format.
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
'   515 - DataTable must have at least one row
'   516 - CSV string cannot be Nothing
'   517 - CSV string cannot be empty
'   519 - CSV string must contain at least one comma
'   520 - CSV row field count mismatch
'
' Usage:
'   Dim serializer As New CsvSerializer
'   serializer.IncludeHeaders = True
'   Dim csv As String
'   csv = serializer.Serialize(myDataTable)
'
' Author: MegaByteMark
' Version: 1.0
'========================================

Option Explicit

'========================================
' Properties
'========================================

Private _includeHeaders As Boolean

''' <summary>
''' Gets or sets a value indicating whether to include headers in the CSV output.
''' </summary>
Public Property Get IncludeHeaders() As Boolean
    IncludeHeaders = _includeHeaders
End Property

''' <summary>
''' Gets or sets a value indicating whether to include headers in the CSV output.
''' </summary>
''' <param name="value">Boolean value to set.</param>
Public Property Let IncludeHeaders(ByVal value As Boolean)
    _includeHeaders = value
End Property

''' <summary>
''' Sets the IncludeHeaders property and returns the CsvSerializer instance for method chaining.
''' </summary>
''' <param name="value">Boolean value to set.</param>
''' <returns>CsvSerializer instance.</returns>
Public Function HasIncludeHeaders(ByVal value As Boolean) As CsvSerializer
    Me.IncludeHeaders = value

    Set HasIncludeHeaders = Me
End Function

''' <summary>
''' Initializes a new instance of the CsvSerializer class.
''' </summary>
Private Sub Class_Initialize()
    ' Default to including headers
    _includeHeaders = True
End Sub

''' <summary>
''' Escapes a CSV field according to CSV rules.
''' </summary>
''' <param name="field">The field to escape.</param>
''' <returns>Escaped CSV field.</returns>
Private Function EscapeCsvField(ByVal field As String) As String
    Dim needsQuotes As Boolean
    Dim escapedField As String
    
    ' Check if the field contains special characters
    needsQuotes = InStr(field, ",") > 0 Or InStr(field, """") > 0 Or InStr(field, vbCr) > 0 Or InStr(field, vbLf) > 0
    
    ' Escape double quotes by doubling them
    escapedField = Replace(field, """", """""")
    
    ' Enclose in double quotes if necessary
    If needsQuotes Then
        escapedField = """" & escapedField & """"
    End If
    
    EscapeCsvField = escapedField
End Function

Private Function UnescapeCsvField(ByVal field As String) As String
    Dim unescapedField As String
    
    ' Remove enclosing double quotes if present
    If Left(field, 1) = """" And Right(field, 1) = """" Then
        unescapedField = Mid(field, 2, Len(field) - 2)
    Else
        unescapedField = field
    End If
    
    ' Replace doubled double quotes with single double quotes
    unescapedField = Replace(unescapedField, """""", """")
    
    UnescapeCsvField = unescapedField
End Function

''' <summary>
''' Serializes a DataTable to a CSV formatted string.
''' </summary>
''' <param name="dataTable">The DataTable to serialize.</param>
''' <returns>CSV formatted string.</returns>
Public Function Serialize(ByVal dataTable As DataTable) As String
    Dim csvLines As Collection
    Dim row As DataRow
    Dim col As DataColumn
    Dim line As String
    Dim i As Long

    If dataTable Is Nothing Then
        Err.Raise vbObjectError + 513, "CsvSerializer", "DataTable cannot be Nothing."
        Exit Function
    End If

    If dataTable.Columns.Count = 0 Then
        Err.Raise vbObjectError + 514, "CsvSerializer", "DataTable must have at least one column."
        Exit Function
    End If

    If dataTable.Rows.Count = 0 Then
        Err.Raise vbObjectError + 515, "CsvSerializer", "DataTable must have at least one row."
        Exit Function
    End If
    
    Set csvLines = New Collection
    
    If (IncludeHeaders) Then
        ' Create header line
        line = ""

        For i = 0 To dataTable.Columns.Count - 1
            line = line & EscapeCsvField(dataTable.Columns(i).Name)

            If i < dataTable.Columns.Count - 1 Then
                line = line & ","
            End If
        Next i

        csvLines.Add line
    End If
    
    ' Create data lines
    For Each row In dataTable.Rows
        line = ""

        For i = 0 To dataTable.Columns.Count - 1
            ' Handle null/empty values
            If IsNull(row.Cells(i).Value) Or IsEmpty(row.Cells(i).Value) Then
                line = line & ""  ' Empty field
            Else
                line = line & EscapeCsvField(CStr(row.Cells(i).Value))
            End If

            If i < dataTable.Columns.Count - 1 Then
                line = line & ","
            End If
        Next i

        csvLines.Add line
    Next row
    
    ' Combine all lines into a single string
    Serialize = Join(CollectionToArray(csvLines), vbCrLf)
End Function

''' <summary>
''' Converts a Collection of strings to an array of strings.
''' </summary>
''' <param name="col">The Collection to convert.</param>
''' <returns>Array of strings.</returns>
Private Function CollectionToArray(col As Collection) As String()
    Dim arr() As String
    Dim i As Long

    If col Is Nothing Then
        Err.Raise 5, "CollectionToArray", "Collection cannot be Nothing."
    End If
    
    If col.Count = 0 Then
        ' Return properly sized empty array
        ReDim arr(0 To -1)  ' Creates a truly empty array
        CollectionToArray = arr
        Exit Function
    End If

    ReDim arr(0 To col.Count - 1)

    For i = 1 To col.Count
        arr(i - 1) = col(i)
    Next i

    CollectionToArray = arr
End Function

''' <summary>
''' Deserializes a CSV formatted string into a DataTable.
''' </summary>
''' <param name="csvString">The CSV formatted string to deserialize.</param>
''' <returns>DataTable instance.</returns>
Public Function Deserialize(ByVal csvString As String) As DataTable
    Dim dataTable As DataTable
    Dim lines() As String
    Dim i As Long, j As Long
    Dim fields() As String
    Dim newRow As DataRow

    If csvString Is Nothing Then
        Err.Raise vbObjectError + 516, "CsvSerializer", "CSV string cannot be Nothing."
        Exit Function
    End If

    If Trim(csvString) = "" Then
        Err.Raise vbObjectError + 517, "CsvSerializer", "CSV string cannot be empty."
        Exit Function
    End If

    If InStr(csvString, ",") = 0 Then
        Err.Raise vbObjectError + 519, "CsvSerializer", "CSV string must contain at least one comma."
        Exit Function
    End If

    Set dataTable = New DataTable
    lines = ReadCsvFile(csvString)
    
    ' Validate we got at least one line
    If UBound(lines) < LBound(lines) Then
        Err.Raise vbObjectError + 521, "CsvSerializer", "CSV contains no valid lines."
        Exit Function
    End If
    
    fields = ReadCsvLine(lines(0))

    If IncludeHeaders Then
        ' Process header line
        For j = LBound(fields) To UBound(fields)
            dataTable.AddColumn UnescapeCsvField(fields(j))  ' Remove Trim()
        Next j
    Else        
        ' Assume first line defines number of columns
        For j = LBound(fields) To UBound(fields)
            dataTable.AddColumn "Column" & (j + 1)
        Next j
    End If

    ' Process data lines
    For i = IIf(IncludeHeaders, 1, 0) To UBound(lines)
        ' Skip empty lines
        If Trim(lines(i)) <> "" Then
            fields = ReadCsvLine(lines(i))

            If (UBound(fields) - LBound(fields) + 1) <> dataTable.Columns.Count Then
                Err.Raise vbObjectError + 520, "CsvSerializer", "CSV row " & (i + 1) & " has " & _
                    (UBound(fields) - LBound(fields) + 1) & " fields but expected " & dataTable.Columns.Count & "."
                Exit Function
            End If

            Set newRow = dataTable.NewRow()

            For j = LBound(fields) To UBound(fields)
                newRow.Cells(j).Value = UnescapeCsvField(fields(j))  ' Remove Trim()
            Next j

            dataTable.AddRow newRow
        End If
    Next i

    Set Deserialize = dataTable
End Function

''' <summary>
''' Reads a single CSV line and splits it into fields, handling quoted fields.
''' </summary>
''' <param name="line">The CSV line to read.</param>
''' <returns>Array of field values.</returns>
Private Function ReadCsvLine(line As String) As String()
    Dim fields As Collection
    Dim currentField As String
    Dim inQuotes As Boolean
    Dim i As Long
    Dim char As String
    Dim nextChar As String
    Dim result() As String
    
    Set fields = New Collection
    currentField = ""
    inQuotes = False
    i = 1

    Do While i <= Len(line)
        char = Mid(line, i, 1)
        
        If char = """" Then
            ' Check for escaped quote (two consecutive quotes)
            If i < Len(line) Then
                nextChar = Mid(line, i + 1, 1)

                If nextChar = """" And inQuotes Then
                    ' Escaped quote - add one quote to field
                    currentField = currentField & """"
                    i = i + 1 ' Skip next quote - now works correctly!
                Else
                    ' Toggle quote state
                    inQuotes = Not inQuotes
                End If
            Else
                ' Last character is a quote
                inQuotes = Not inQuotes
            End If
        ElseIf char = "," And Not inQuotes Then
            ' Field separator - found comma outside quotes
            fields.Add currentField
            currentField = ""
        Else
            ' Regular character
            currentField = currentField & char
        End If
        
        i = i + 1
    Loop
    
    ' Add last field
    fields.Add currentField

    ' Convert collection to array
    result = CollectionToArray(fields)
    
    ReadCsvLine = result
End Function

''' <summary>
''' Reads a CSV file content and returns it as an array of lines.
''' </summary>
''' <param name="csvFile">The CSV file content as a string.</param>
''' <returns>Array of lines from the CSV file.</returns>
Private Function ReadCsvFile(csvFile As String) As String()

    ' Reads the entire content of a CSV file and returns it as an array of lines.
    ' The problem we are trying to solve here is that a CSV value which may be quoted could include line breaks.
    ' If we split on the line before processing the quoted values in ReadCsvLine, we could end up breaking a single value into multiple lines and corrupt the data.
    ' Therefore, we read the entire file content as a single string and return it as an array of string one per line correctly handling quoted line breaks.

    Dim lines As Collection
    Dim currentLine As String
    Dim inQuotes As Boolean
    Dim i As Long
    Dim char As String
    Dim nextChar As String
    Dim result() As String

    Set lines = New Collection
    currentLine = ""
    inQuotes = False
    i = 1

    Do While i <= Len(csvFile)
        char = Mid(csvFile, i, 1)
        
        If char = """" Then
            ' Check for escaped quote (two consecutive quotes)
            If i < Len(csvFile) Then
                nextChar = Mid(csvFile, i + 1, 1)

                If nextChar = """" And inQuotes Then
                    ' Escaped quote - add both quotes to line
                    currentLine = currentLine & """"""
                    i = i + 1 ' Skip next quote
                Else
                    ' Toggle quote state and add quote to line
                    inQuotes = Not inQuotes
                    currentLine = currentLine & """"
                End If
            Else
                ' Last character is a quote
                inQuotes = Not inQuotes
                currentLine = currentLine & """"
            End If
        ElseIf (char = vbCr Or char = vbLf) And Not inQuotes Then
            ' Line separator - found line break outside quotes
            lines.Add currentLine
            currentLine = ""
            
            ' Handle CRLF sequence - skip the LF if next char is LF
            If char = vbCr And i < Len(csvFile) Then
                If Mid(csvFile, i + 1, 1) = vbLf Then
                    i = i + 1 ' Skip the LF in CRLF
                End If
            End If
        Else
            ' Regular character (including line breaks inside quotes)
            currentLine = currentLine & char
        End If
        
        i = i + 1
    Loop
    
    ' Add last line if not empty
    If Len(currentLine) > 0 Or lines.Count = 0 Then
        lines.Add currentLine
    End If

    ' Convert collection to array
    result = CollectionToArray(lines)

    ReadCsvFile = result

End Function