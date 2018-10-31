VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cancelation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
    Dim sUrl As String
    Dim http As MSXML2.ServerXMLHTTP
    Dim body As String
    Dim quote As String
    Dim sp As String
    
    Public Function CancelationByCSD(url As String, b64Cer As String, b64Key As String, PasswordKey As String, RFC As String, UUID As String, Token As String) As String
        Set http = New MSXML2.ServerXMLHTTP
        quote = """"
        sUrl = url & "/cfdi33/cancel/csd"
body = "{ " & Chr(34) & "uuid" & Chr(34) & ": " & Chr(34) & UUID & Chr(34) & "," _
            & Chr(34) & "password" & Chr(34) & ": " & Chr(34) & PasswordKey & Chr(34) & "," _
            & Chr(34) & "rfc" & Chr(34) & ": " & Chr(34) & RFC & Chr(34) & "," _
            & Chr(34) & "b64Cer" & Chr(34) & ": " & Chr(34) & b64Cer & Chr(34) & "," _
            & Chr(34) & "b64Key" & Chr(34) & ": " & Chr(34) & b64Key & Chr(34) & "}"
            
        With http
            .open "POST", sUrl, False
            .setRequestHeader "Authorization", "bearer " & Token
            .setRequestHeader "Cache-Control", "no-cache"
            .setRequestHeader "Content-Type", "application/json"
            .send (body)

        End With
        
        CancelationByCSD = http.responseText
        
    End Function

    Public Function CancelationByPFX(url As String, b64pfx As String, password As String, RFC As String, UUID As String, Token As String) As String
        Set http = New MSXML2.ServerXMLHTTP
        sUrl = url & "/cfdi33/cancel/pfx"
        Dim quote As String

     body = "{ " & Chr(34) & "uuid" & Chr(34) & ": " & Chr(34) & UUID & Chr(34) & "," _
                 & Chr(34) & "password" & Chr(34) & ": " & Chr(34) & password & Chr(34) & "," _
                 & Chr(34) & "rfc" & Chr(34) & ": " & Chr(34) & RFC & Chr(34) & "," _
                 & Chr(34) & "b64Pfx" & Chr(34) & ": " & Chr(34) & b64pfx & Chr(34) & "}"

        With http
            .open "POST", sUrl, False
            .setRequestHeader "Content-Type", "application/json"
            .setRequestHeader "Cache-Control", "no-cache"
            .setRequestHeader "Authorization", "bearer " & Token
            .send (body)

        End With

        CancelationByPFX = http.responseText

    End Function

    Public Function CancelationByUUID(url As String, RFC As String, UUID As String, Token As String) As String
        Set http = New MSXML2.ServerXMLHTTP
        sUrl = url & "/cfdi33/cancel/" & RFC & "/" & UUID
        
        With http
            .open "POST", sUrl, False
            .setRequestHeader "Content-Type", "application/json"
            .setRequestHeader "Authorization", "bearer " & Token
            .send

        End With
        
        CancelationByUUID = http.responseText
        
    End Function
    
    Public Function CancelationByXML(url As String, XML As String, Token As String) As String
        Set http = New MSXML2.ServerXMLHTTP
        sp = "\r\n"
        bound = "AaB03x"
        XML = Replace(XML, "﻿", "")
        XML = Replace(XML, "ï»¿", "")
        sUrl = url & "/cfdi33/cancel/xml"

        body = ""
        body = body & "--" & bound & vbCrLf & "Content-Disposition: form-data; name=xml; filename=xml" & vbCrLf & "Content-Transfer-Encoding: binary" & vbCrLf & "Content-Type: text/xml" & vbCrLf & vbCrLf & XML & vbCrLf & "--" + bound & "--" & vbCrLf

        With http
            .open "POST", sUrl, False
            .setRequestHeader "Content-Type", "multipart/form-data; boundary=" & bound
            .setRequestHeader "Authorization", "bearer " & Token
            .setRequestHeader "Content-Length", Len(body)
            .send (body)

        End With

        CancelationByXML = http.responseText

    End Function

