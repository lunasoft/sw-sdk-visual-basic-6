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
Attribute VB_Exposed = True
 
Dim http As New MSXML2.ServerXMLHTTP

    Public Function Token(url As String, user As String, pwd As String) As String
        


       



        http.open "POST", url, False
        http.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"

        http.setRequestHeader "user", user

        http.setRequestHeader "password", pwd

        http.send

        Token = http.responseText
        
        'http = Nothing

    End Function

