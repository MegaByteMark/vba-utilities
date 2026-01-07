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
' DataRow Class
'========================================
' Provides a definition for a row in a DataTable
'
' Dependencies:
' None
'
' Error Codes:
'  None
'
' Usage:
'  Dim row As DataRow
'  Set row = New DataRow
'  Call row.AddValue 123
'  Call row.AddValue "Sample Text"
'
' Author: MegaByteMark
' Version: 2.0
'========================================

Option Explicit

'========================================
' Properties
'========================================

Private items() As Variant

''' <summary>
''' Initializes a new DataRow instance
''' </summary>
''' <remarks>Creates an empty items array</remarks>
Private Sub Class_Initialize()
    ReDim items(0 To -1)
End Sub

''' <summary>
''' Cleans up resources when the DataRow is destroyed
''' </summary>
''' <remarks>Releases all item references</remarks>
Private Sub Class_Terminate()
    On Error Resume Next
    Erase items
    On Error GoTo 0
End Sub

'========================================
' Item Operations
'========================================

''' <summary>
''' Gets the total number of items in the row
''' </summary>
''' <returns>The count of items</returns>
Public Function ItemCount() As Long
    On Error Resume Next
    If UBound(items) < LBound(items) Then
        ItemCount = 0
    Else
        ItemCount = UBound(items) - LBound(items) + 1
    End If
    On Error GoTo 0
End Function

''' <summary>
''' Adds a value to the row
''' </summary>
''' <param name="value">The value to add</param>
''' <returns>The index where the value was added</returns>
Public Function AddValue(ByVal value As Variant) As Long
    Dim currentCount As Long
    
    currentCount = ItemCount()
    
    ' Resize array to accommodate new item
    ReDim Preserve items(0 To currentCount)

    If IsEmpty(value) Or IsNull(value) Then
        items(currentCount) = value
    ElseIf IsObject(value) Then
        If value Is Nothing Then
            items(currentCount) = Nothing
        Else
            Set items(currentCount) = value
        End If
    Else
        items(currentCount) = value
    End If
    
    AddValue = currentCount
End Function

''' <summary>
''' Gets a value from the row by its ordinal position
''' </summary>
''' <param name="ordinal">The zero-based index of the value</param>
''' <returns>The value at the specified position</returns>
''' <remarks>Raises error 521 if ordinal is out of bounds</remarks>
Public Function GetValue(ByVal ordinal As Long) As Variant
    If ItemCount() = 0 Then
        Err.Raise vbObjectError + 521, "DataRow.GetValue", "Row has no items."
    End If
    
    If ordinal < LBound(items) Or ordinal > UBound(items) Then
        Err.Raise vbObjectError + 521, "DataRow.GetValue", "Item ordinal is out of bounds."
    End If
    
    If IsObject(items(ordinal)) Then
        Set GetValue = items(ordinal)
    Else
        GetValue = items(ordinal)
    End If
End Function

''' <summary>
''' Sets a value in the row at a specific ordinal position
''' </summary>
''' <param name="ordinal">The zero-based index where to set the value</param>
''' <param name="value">The value to set</param>
''' <remarks>Raises error 522 if ordinal is out of bounds</remarks>
Public Sub SetValue(ByVal ordinal As Long, ByVal value As Variant)
    If ItemCount() = 0 Then
        Err.Raise vbObjectError + 522, "DataRow.SetValue", "Row has no items."
    End If
    
    If ordinal < LBound(items) Or ordinal > UBound(items) Then
        Err.Raise vbObjectError + 522, "DataRow.SetValue", "Item ordinal is out of bounds."
    End If
    
    ' Handle Empty, Null, and Nothing
    If IsEmpty(value) Or IsNull(value) Then
        items(ordinal) = value
    ElseIf IsObject(value) Then
        If value Is Nothing Then
            Set items(ordinal) = Nothing
        Else
            Set items(ordinal) = value
        End If
    Else
        items(ordinal) = value
    End If
End Sub

''' <summary>
''' Removes a value from the row by its ordinal position
''' </summary>
''' <param name="ordinal">The zero-based index of the value to remove</param>
''' <returns>True if successful, False otherwise</returns>
Public Function RemoveValue(ByVal ordinal As Long) As Boolean
    Dim i As Long
    Dim currentCount As Long
    
    currentCount = ItemCount()
    
    If currentCount = 0 Then
        RemoveValue = False
        Exit Function
    End If
    
    If ordinal < LBound(items) Or ordinal > UBound(items) Then
        RemoveValue = False
        Exit Function
    End If
    
    ' Clean up object reference if applicable
    On Error Resume Next

    If IsObject(items(ordinal)) Then
        Set items(ordinal) = Nothing
    End If
    
    On Error GoTo 0
    
    ' Shift remaining items down to fill the gap
    For i = ordinal To UBound(items) - 1
        If IsObject(items(i + 1)) Then
            Set items(i) = items(i + 1)
        Else
            items(i) = items(i + 1)
        End If
    Next i
    
    ' Resize array to trim off the last element
    If currentCount > 1 Then
        ReDim Preserve items(0 To currentCount - 2)
    Else
        ReDim items(0 To -1)  ' Empty array if removing last item
    End If
    
    RemoveValue = True
End Function

''' <summary>
''' Removes all values from the row
''' </summary>
''' <returns>True if successful</returns>
Public Function Clear() As Boolean
    On Error Resume Next

    Erase items
    ReDim items(0 To -1)
    Clear = True
    
    On Error GoTo 0
End Function

''' <summary>
''' Gets all items as an array
''' </summary>
''' <returns>Array of all values in the row, or Empty if row has no items</returns>
Public Function ToArray() As Variant
    If ItemCount() = 0 Then
        ToArray = Array()  ' Return empty array instead of uninitialized array
    Else
        ToArray = items
    End If
End Function

''' <summary>
''' Creates a copy of the row with all its values
''' </summary>
''' <returns>A new DataRow with the same values</returns>
Public Function Clone() As DataRow
    Dim newRow As DataRow
    Dim i As Long
    
    Set newRow = New DataRow
    
    On Error GoTo CloneError
    
    If ItemCount() > 0 Then
        For i = LBound(items) To UBound(items)
            Call newRow.AddValue(items(i))
        Next i
    End If
    
    Set Clone = newRow
    Exit Function
    
CloneError:
    Set Clone = Nothing
    Err.Raise Err.Number, "DataRow.Clone", "Error cloning row: " & Err.Description
End Function