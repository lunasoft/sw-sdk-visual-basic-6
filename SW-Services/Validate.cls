VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Validate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
    Dim sUrl As String
    Dim http As MSXML2.ServerXMLHTTP
    Dim body As String
    Dim bound As String
    Dim sp As String
    
    Public Function Validate(url As String, Token As String, XML As String) As String
        Set http = New MSXML2.ServerXMLHTTP
        sp = "\r\n"
        bound = "AaB03x"
        sUrl = url & "/validate/cfdi33/"
        
        body = ""
        body = body & "--" & bound & vbCrLf & "Content-Disposition: form-data; name=xml; filename=xml" & vbCrLf & "Content-Transfer-Encoding: binary" & vbCrLf & "Content-Type: text/xml" & vbCrLf & vbCrLf & XML & vbCrLf & "--" + bound & "--" & vbCrLf
        
        With http
            .open "POST", sUrl, False
            .setRequestHeader "Content-Type", "multipart/form-data; boundary=" & bound
            .setRequestHeader "Authorization", "bearer " & Token
            .setRequestHeader "Content-Length", Len(body)
            .send (body)
            
        End With
        
        Validate = http.responseText
        
    End Function
    
   

