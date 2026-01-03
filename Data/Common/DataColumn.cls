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
Private dataType As VbVarType  
Private allowNull As Boolean   
Private defaultValue As Variant

''' <summary>
''' Gets the name of the column
''' </summary>
''' <returns>The name of the column</returns>
Public Property Get Name() As String
    Name = name
End Property

''' <summary>
''' Sets the name of the column
''' </summary>
''' <param name="value">The name to set</param>
Public Property Let Name(ByVal value As String)
    name = value
End Property

''' <summary>
''' Gets the data type of the column
''' </summary>
''' <returns>VbVarType enumeration value</returns>
Public Property Get DataType() As VbVarType
    DataType = dataType
End Property

''' <summary>
''' Sets the data type of the column
''' </summary>
''' <param name="value">VbVarType enumeration value</param>
Public Property Let DataType(ByVal value As VbVarType)
    dataType = value
End Property

''' <summary>
''' Gets whether the column allows null values
''' </summary>
''' <returns>True if nulls are allowed, False otherwise</returns>
Public Property Get AllowNull() As Boolean
    AllowNull = allowNull
End Property

''' <summary>
''' Sets whether the column allows null values
''' </summary>
''' <param name="value">True to allow nulls, False otherwise</param>
Public Property Let AllowNull(ByVal value As Boolean)
    allowNull = value
End Property

''' <summary>
''' Gets the default value for the column
''' </summary>
''' <returns>The default value</returns>
Public Property Get DefaultValue() As Variant
    If IsObject(defaultValue) Then
        Set DefaultValue = defaultValue
    Else
        DefaultValue = defaultValue
    End If
End Property

''' <summary>
''' Sets the default value for the column
''' </summary>
''' <param name="value">The default value to set</param>
Public Property Let DefaultValue(ByVal value As Variant)
    defaultValue = value
End Property

''' <summary>
''' Sets the default value for the column (object version)
''' </summary>
''' <param name="value">The default value to set</param>
Public Property Set DefaultValue(ByVal value As Variant)
    Set defaultValue = value
End Property

''' <summary>
''' Initializes a new DataColumn instance
''' </summary>
''' <remarks>Sets default values: AllowNull=True, DataType=vbVariant, DefaultValue=Empty</remarks>
Private Sub Class_Initialize()
    allowNull = True
    dataType = vbVariant
    defaultValue = Empty
End Sub

''' <summary>
''' Validates if a value is compatible with the column's data type
''' </summary>
''' <param name="testValue">The value to validate</param>
''' <returns>True if valid, False otherwise</returns>
Public Function CanAcceptValue(ByVal testValue As Variant) As Boolean
    ' Check for null
    If IsNull(testValue) Or IsEmpty(testValue) Then
        CanAcceptValue = allowNull
        Exit Function
    End If
    
    ' If dataType is Variant, accept anything
    If dataType = vbVariant Then
        CanAcceptValue = True
        Exit Function
    End If
    
    ' Check type compatibility
    On Error Resume Next

    CanAcceptValue = (VarType(testValue) = dataType)

    On Error GoTo 0
End Function

''' <summary>
''' Creates a copy of the column definition
''' </summary>
''' <returns>A new DataColumn with the same properties</returns>
''' <remarks>Does not copy the value, only the schema</remarks>
Public Function Clone() As DataColumn
    Dim newColumn As DataColumn
    Set newColumn = New DataColumn
    
    newColumn.Name = Me.Name
    newColumn.DataType = Me.DataType
    newColumn.AllowNull = Me.AllowNull
    
    If IsObject(Me.DefaultValue) Then
        Set newColumn.DefaultValue = Me.DefaultValue
    Else
        newColumn.DefaultValue = Me.DefaultValue
    End If
    
    Set Clone = newColumn
End Function