VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Authentication"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
    Dim http As New MSXML2.ServerXMLHTTP

    Public Function Token(url As String, user As String, pwd As String) As String
    
        With http
            .open "POST", url & "/security/authenticate", False
            .setRequestHeader "password", pwd
            .setRequestHeader "user", user
            .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
            .send
        End With

        Token = http.responseText
        
    End Function

