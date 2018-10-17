VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Relations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
    Dim sUrl As String
    Dim http As MSXML2.ServerXMLHTTP
    Dim body As String
    
    Public Function RelationsByCSD(url As String, Token As String, b64Cer As String, b64Key As String, PasswordKey As String, RFC As String, UUID As String) As String
        Set http = New MSXML2.ServerXMLHTTP
        sUrl = url & "/relations/csd"
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
        
        RelationsByCSD = http.responseText
        
    End Function

    Public Function RelationsByPFX(url As String, Token As String, b64pfx As String, password As String, RFC As String, UUID As String) As String
        Set http = New MSXML2.ServerXMLHTTP
        sUrl = url & "/relations/pfx"
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

        RelationsByPFX = http.responseText

    End Function

    Public Function RelationsByUUID(url As String, Token As String, RFC As String, UUID As String) As String
        Set http = New MSXML2.ServerXMLHTTP
        sUrl = url & "/relations/" & RFC & "/" & UUID
        
        With http
            .open "POST", sUrl, False
            .setRequestHeader "Content-Type", "application/json"
            .setRequestHeader "Authorization", "bearer " & Token
            .send

        End With
        
        RelationsByUUID = http.responseText
        
    End Function
    
    Public Function RelationsByXML(url As String, Token As String, XML As String) As String
        Set http = New MSXML2.ServerXMLHTTP
        Dim bound As String
        bound = "AaB03x"
        XML = Replace(XML, "﻿", "")
        XML = Replace(XML, "ï»¿", "")
        sUrl = url & "/relations/xml"

        body = ""
        body = body & "--" & bound & vbCrLf & "Content-Disposition: form-data; name=xml; filename=xml" & vbCrLf & "Content-Transfer-Encoding: binary" & vbCrLf & "Content-Type: text/xml" & vbCrLf & vbCrLf & XML & vbCrLf & "--" + bound & "--" & vbCrLf

        With http
            .open "POST", sUrl, False
            .setRequestHeader "Content-Type", "multipart/form-data; boundary=" & bound
            .setRequestHeader "Authorization", "bearer " & Token
            .setRequestHeader "Content-Length", Len(body)
            .send (body)

        End With

        RelationsByXML = http.responseText

    End Function



