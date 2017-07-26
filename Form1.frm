VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9195
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   22560
   LinkTopic       =   "Form1"
   ScaleHeight     =   9195
   ScaleWidth      =   22560
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox body 
      Height          =   8415
      Left            =   4200
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   240
      Width           =   18015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "stamp"
      Height          =   615
      Left            =   960
      TabIndex        =   1
      Top             =   1920
      Width           =   2415
   End
   Begin VB.CommandButton auth 
      Caption         =   "auth"
      Height          =   1095
      Left            =   720
      TabIndex        =   0
      Top             =   0
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub auth_Click()
Dim auth As New Authentication
MsgBox auth.token("http://services.test.sw.com.mx", "demo", "123456789")

End Sub

Private Sub Command1_Click()
Dim stamp As New stamp
Dim xml As String
Dim token As String
token = "T2lYQ0t4L0RHVkR4dHZ5Nkk1VHNEakZ3Y0J4Nk9GODZuRyt4cE1wVm5tbXB3YVZxTHdOdHAwVXY2NTdJb1hkREtXTzE3dk9pMmdMdkFDR2xFWFVPUTQyWFhnTUxGYjdKdG8xQTZWVjFrUDNiOTVrRkhiOGk3RHladHdMaEM0cS8rcklzaUhJOGozWjN0K2h6R3gwQzF0c0g5aGNBYUt6N2srR3VoMUw3amtvPQ.T2lYQ0t4L0RHVkR4dHZ5Nkk1VHNEakZ3Y0J4Nk9GODZuRyt4cE1wVm5tbFlVcU92YUJTZWlHU3pER1kySnlXRTF4alNUS0ZWcUlVS0NhelhqaXdnWTRncklVSWVvZlFZMWNyUjVxYUFxMWFxcStUL1IzdGpHRTJqdS9Zakw2UGRiMTFPRlV3a2kyOWI5WUZHWk85ODJtU0M2UlJEUkFTVXhYTDNKZVdhOXIySE1tUVlFdm1jN3kvRStBQlpLRi9NeWJrd0R3clhpYWJrVUMwV0Mwd3FhUXdpUFF5NW5PN3J5cklMb0FETHlxVFRtRW16UW5ZVjAwUjdCa2g0Yk1iTExCeXJkVDRhMGMxOUZ1YWlIUWRRVC8yalFTNUczZXdvWlF0cSt2UW0waFZKY2gyaW5jeElydXN3clNPUDNvU1J2dm9weHBTSlZYNU9aaGsvalpQMUx2ckVwVHFKd2ZFUmZ4dFhMSXdIdWFySXh2amlTcFlvTEh2VHk1RWVYTDVGTVEwdVhmMzJZSmo5VStvUk9vT01iaVNOSGtGd1FnbzJ3RDZBbXFPRDgxQllOU3E5djR1Z0NrQWdpbjRVTWk2RFBZa0Naa21qR1UxTHVhUmprVVhCU0NzbmNpN3BCVXRsT1RueHZpZkNxZU09.w5yUO_iK0f4yo_3rRtmp_b9tOt93lL6Wb45nGdiugCQ"
xml = FileToString("C:\Users\asalvio\Documents\WORKSPACE\VISUAL BASIC 6\native\sw.services\33.xml")

MsgBox stamp.stampV1("http://services.test.sw.com.mx", xml, "v1", token)
End Sub
Function FileToString(strFilename As String) As String
  iFile = FreeFile
  Open strFilename For Input As #iFile
    FileToString = StrConv(InputB(LOF(iFile), iFile), vbUnicode)
  Close #iFile
End Function
