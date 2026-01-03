VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SqlConnectionStringBuilder"
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

Private ServerName As String
Private DatabaseName As String
Private integratedSec As Boolean
Private userName As String
Private pwd As String
Private appName As String
Private trustServerCert As Boolean
Private trustedConn As Boolean
Private encryptConn As Boolean
Private driverName As String
Private providerName As String

Public Property Let Server(value As String)
    ServerName = value
End Property

Public Property Get Server() As String
    Server = ServerName
End Property

Public Property Let Database(value As String)
    DatabaseName = value
End Property

Public Property Get Database() As String
    Database = DatabaseName
End Property

Public Property Let IntegratedSecurity(value As Boolean)
    integratedSec = value
End Property

Public Property Get IntegratedSecurity() As Boolean
    IntegratedSecurity = integratedSec
End Property

Public Property Let UserId(value As String)
    userName = value
End Property

Public Property Get UserId() As String
    UserId = userName
End Property

Public Property Let Password(value As String)
    pwd = value
End Property

Public Property Get Password() As String
    Password = pwd
End Property

Public Property Let ApplicationName(value As String)
    appName = value
End Property

Public Property Get ApplicationName() As String
    ApplicationName = appName
End Property

Public Property Let TrustServerCertificate(value As Boolean)
    trustServerCert = value
End Property

Public Property Get TrustServerCertificate() As Boolean
    TrustServerCertificate = trustServerCert
End Property

Public Property Let TrustedConnection(value As Boolean)
    trustedConn = value
End Property

Public Property Get TrustedConnection() As Boolean
    TrustedConnection = trustedConn
End Property

Public Property Let Encrypt(value As Boolean)
    encryptConn = value
End Property

Public Property Get Encrypt() As Boolean
    Encrypt = encryptConn
End Property

Public Property Let Driver(value As String)
    driverName = value
End Property

Public Property Get Driver() As String
    Driver = driverName
End Property

Public Property Let Provider(value As String)
    providerName = value
End Property

Public Property Get Provider() As String
    Provider = providerName
End Property

Public Function ToString() As String

    Dim connString As String
    
    connString = ""
    AppendString connString, "Provider", providerName
    AppendString connString, "Driver", driverName
    AppendString connString, "Data Source", ServerName
    AppendString connString, "Initial Catalog", DatabaseName
    AppendString connString, "Integrated Security", IIf(integratedSec, "SSPI", "False")
    
    If integratedSec Then
        AppendString connString, "User Id", userName
        AppendString connString, "Password", pwd
    End If
    
    If trustedConn Then
        AppendBoolean connString, "Trusted_Connection", trustedConn
    End If
    
    If trustServerCert Then
        AppendBoolean connString, "TrustServerCertificate", trustServerCert
    End If
    
    If Encrypt Then
        AppendString connString, "Encrypt", IIf(Encrypt, "yes", "no")
    End If
    
    AppendString connString, "Application Name", appName
    
    ToString = connString

End Function

Private Sub AppendString(ByRef TheString As String, ByVal TheKey As String, ByVal TheValue As String)

    If Not IsNull(TheValue) And Len(TheValue) > 0 Then
    
        TheString = TheString & TheKey & "=" & TheValue & ";"
    
    End If

End Sub

Private Sub AppendBoolean(ByRef TheString As String, ByVal TheKey As String, ByVal TheValue As Boolean)

    If Not IsNull(TheValue) And Len(TheValue) > 0 Then
    
        If TheValue Then
            TheString = TheString & TheKey & "=True;"
        Else
            TheString = TheString & TheKey & "=False;"
        End If
    
    End If

End Sub
