VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WindowsEventLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Private Declare PtrSafe Function RegisterEventSource Lib "advapi32.dll" Alias "RegisterEventSourceW" ( _
    ByVal lpUNCServerName As LongPtr, _
    ByVal lpSourceName As LongPtr) As LongPtr

Private Declare PtrSafe Function ReportEvent Lib "advapi32.dll" Alias "ReportEventW" ( _
    ByVal hEventLog As LongPtr, _
    ByVal wType As Integer, _
    ByVal wCategory As Integer, _
    ByVal dwEventID As Long, _
    ByVal lpUserSid As LongPtr, _
    ByVal wNumStrings As Integer, _
    ByVal dwDataSize As Long, _
    ByRef lpStrings As LongPtr, _
    ByVal lpRawData As LongPtr) As Long

Private logName As String
Private appName As String
Private verboseLogging As Boolean

Public Property Let Log(value As String)
    logName = value
End Property

Public Property Get Log() As String
    Log = logName
End Property

Public Property Let Application(value As String)
    appName = value
End Property

Public Property Get Application() As String
    Application = appName
End Property

Public Property Let Verbose(value As Boolean)
    verboseLogging = value
End Property

Public Property Get Verbose() As Boolean
    Verbose = verboseLogging
End Property

Public Sub Initialize(ByVal ApplicationName As String, Optional ByVal EventLogName As String = "Application", Optional ByVal VerboseEventLogging As Boolean = False)
    appName = ApplicationName
    logName = EventLogName
    verboseLogging = VerboseEventLogging
End Sub

Public Sub WriteEntry(ByVal Message As String, Optional ByVal EntryType As String = "I")
    On Error GoTo ErrHandler

    Dim hEventLog As LongPtr
    Dim messagePtr As LongPtr
    
    Dim wType As Integer
    Dim result As Long

    Select Case UCase(EntryType)
        Case "E": wType = 1  ' Error
        Case "W": wType = 2  ' Warning
        Case "I": wType = 4  ' Information
        Case Else: wType = 4
    End Select
    
    ' Register with custom source name
    hEventLog = RegisterEventSource(0, StrPtr(Me.Application))
    
    If hEventLog = 0 Then
        Debug.Print "Failed to register event source: " & Me.Application
        Exit Sub
    End If
    
    messagePtr = StrPtr(Message)
    result = ReportEvent(hEventLog, wType, 0, 1000, 0, 1, 0, messagePtr, 0)
    
    If result = 0 Then
        Debug.Print "Failed to write event"
    End If
    
    Exit Sub

ErrHandler:
    Debug.Print Err.Number & "-" & Err.Description
    Err.Clear
    
    Exit Sub

End Sub