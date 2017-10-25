VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Stamp"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
  Dim sUrl As String
    Dim http As MSXML2.ServerXMLHTTP
    Dim body As String
    Dim bound As String
    Dim sp As String
    Public Function Stamp(url As String, xml As String, version As String, Token As String) As String
        Set http = New MSXML2.ServerXMLHTTP
        sp = "\r\n"
        bound = "AaB03x"
        sUrl = url & "/cfdi33/stamp/" & version
        
        
        http.open "POST", sUrl, False
        http.setRequestHeader "Content-Type", "multipart/form-data; boundary=" & bound

        http.setRequestHeader "Authorization", "bearer " & Token
        http.setRequestHeader "Content-Length", Len(body)

        
        body = ""
        
        body = body & "--" & bound & vbCrLf & "Content-Disposition: form-data; name=xml; filename=xml" & vbCrLf & "Content-Transfer-Encoding: binary" & vbCrLf & "Content-Type: text/xml" & vbCrLf & vbCrLf & xml & vbCrLf & "--" + bound & "--" & vbCrLf
        log (body)

        http.send (body)

        StampV1 = http.responseText

       
    End Function
    
   

    

Private Sub Class_Initialize()

End Sub
