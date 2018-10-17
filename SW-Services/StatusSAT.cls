VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StatusSAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
 
    Public Function StatusCFDI(url As String, Action As String, RfcEmisor As String, RfcReceptor As String, total As String, UUID As String) As String
    Set http = New MSXML2.ServerXMLHTTP
    Dim body As String
    body = "<soapenv:Envelope xmlns:soapenv=" & Chr(34) & "http://schemas.xmlsoap.org/soap/envelope/" & Chr(34) & " xmlns:tem=" & Chr(34) & "http://tempuri.org/" & Chr(34) & ">" _
        & "<soapenv:Header/>" _
        & "<soapenv:Body>" _
        & "<tem:Consulta>" _
        & "<!--Optional:-->" _
        & "<tem:expresionImpresa><![CDATA[?re=" & RfcEmisor & "&rr=" & RfcReceptor & "&tt=" & total & "&id=" & UUID & "]]></tem:expresionImpresa>" _
        & "</tem:Consulta>" _
        & "</soapenv:Body>" _
        & "</soapenv:Envelope>"
        With http
            .open "POST", url, False
            .setRequestHeader "SOAPAction", Action
            .setRequestHeader "Content-Type", "text/xml; charset=utf-8"
            .send (body)
        End With
        MsgBox (http.responseText)
        StatusCFDI = (http.responseText)
        
    End Function

Public Function procesarRespuesta(ByVal s As String) As String
    ' Poner los datos en el analizador de XML
    Dim parser As DOMDocument
    Set parser = New DOMDocument
    Dim respond As String
    
    parser.loadXML s
    respond = parser.selectSingleNode("/s:Envelope/s:Body/ConsultaResponse/ConsultaResult/a:CodigoEstatus").Text
    On Error Resume Next
     procesarRespuesta = parser.selectSingleNode("/s:Envelope/s:Body/s:Fault/faultstring").Text
End Function


