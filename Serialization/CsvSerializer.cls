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
            line = line & EscapeCsvField(CStr(row.Cells(i).Value))

            If i < dataTable.Columns.Count - 1 Then
                line = line & ","
            End If
        Next i

        csvLines.Add line
    Next row
    
    ' Combine all lines into a single string
    Serlialize = Join(CollectionToArray(csvLines), vbCrLf)
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

    If InStr(csvString, vbCrLf) = 0 Then
        Err.Raise vbObjectError + 518, "CsvSerializer", "CSV string must contain at least one line."
        Exit Function
    End If

    If InStr(csvString, ",") = 0 Then
        Err.Raise vbObjectError + 519, "CsvSerializer", "CSV string must contain at least one comma."
        Exit Function
    End If

    Set dataTable = New DataTable
    lines = Split(csvString, vbCrLf)

    'TODO - This simple Split approach does not handle quoted commas or line breaks within fields.
    fields = Split(lines(0), ",")

    If IncludeHeaders Then
        ' Process header line
        For j = LBound(fields) To UBound(fields)
            dataTable.Columns.Add UnescapeCsvField(Trim(fields(j)))
        Next j
    Else        
        ' Assume first line defines number of columns
        For j = LBound(fields) To UBound(fields)
            dataTable.Columns.Add "Column" & (j + 1)
        Next j
    End If

    ' Process data lines
    For i = IIf(IncludeHeaders, 1, 0) To UBound(lines)
        'TODO - This simple Split approach does not handle quoted commas or line breaks within fields.
        fields = Split(lines(i), ",")
        Set newRow = dataTable.NewRow()

        If(fields.Length <> dataTable.Columns.Count) Then
            Err.Raise vbObjectError + 520, "CsvSerializer", "CSV row " & (i) & " does not match number of columns."
            Exit Function
        End If

        For j = LBound(fields) To UBound(fields)
            newRow.Add.Value = UnescapeCsvField(Trim(fields(j)))
        Next j

        dataTable.Rows.Add newRow
    Next i

    Set Deserialize = dataTable
End Function