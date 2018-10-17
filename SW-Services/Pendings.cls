VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Pendings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
    Dim http As New MSXML2.ServerXMLHTTP
    
    Public Function Pendings(url As String, Token As String, RFC As String) As String
    
        With http
            .open "GET", url & "/pendings/" & RFC, False
            .setRequestHeader "Authorization", "bearer " & Token
            .setRequestHeader "Content-Type", "application/json"
            .send
        End With

        Pendings = http.responseText
        
    End Function



