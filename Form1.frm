VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   13275
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14970
   LinkTopic       =   "Form1"
   ScaleHeight     =   13275
   ScaleWidth      =   14970
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   855
      Left            =   360
      TabIndex        =   22
      Top             =   12240
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   1508
      _Version        =   393216
      Appearance      =   1
      Max             =   5000
   End
   Begin VB.Frame Frame3 
      Caption         =   "Pruebas Unitarias"
      Height          =   6015
      Left            =   240
      TabIndex        =   23
      Top             =   7200
      Width           =   13455
      Begin VB.TextBox txtLogs 
         Height          =   1815
         Left            =   240
         TabIndex        =   27
         Text            =   "Logs: "
         Top             =   3120
         Width           =   12975
      End
      Begin VB.TextBox txtXmlResult 
         Height          =   1335
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   26
         Top             =   1560
         Width           =   11175
      End
      Begin VB.TextBox txtXmlRequest 
         Height          =   1215
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   25
         Top             =   240
         Width           =   11175
      End
      Begin VB.CommandButton btnStartTests 
         Caption         =   "Comenzar ejecución de pruebas unitarias"
         Height          =   1335
         Left            =   11520
         TabIndex        =   24
         Top             =   840
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Autenticacion"
      Height          =   6855
      Left            =   240
      TabIndex        =   10
      Top             =   120
      Width           =   6495
      Begin VB.CommandButton btnToken 
         Caption         =   "Obtener Token"
         Height          =   1455
         Left            =   4320
         TabIndex        =   18
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox txtToken 
         Height          =   3135
         Left            =   1680
         MultiLine       =   -1  'True
         TabIndex        =   17
         Top             =   3480
         Width           =   4575
      End
      Begin VB.TextBox txtUrlToken 
         Height          =   735
         Left            =   1680
         TabIndex        =   13
         Top             =   480
         Width           =   2415
      End
      Begin VB.TextBox txtPasswordToken 
         Height          =   735
         Left            =   1680
         TabIndex        =   12
         Top             =   2400
         Width           =   2415
      End
      Begin VB.TextBox txtUserToken 
         Height          =   855
         Left            =   1680
         TabIndex        =   11
         Top             =   1320
         Width           =   2415
      End
      Begin VB.Label Label9 
         Caption         =   "Resultado Token"
         Height          =   615
         Left            =   360
         TabIndex        =   21
         Top             =   4680
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "URL:"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Password:"
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Usuario:"
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   1560
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Timbrado V4"
      Height          =   6855
      Left            =   7080
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      Begin VB.TextBox txtXmlB64 
         Height          =   1575
         Left            =   1200
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   3000
         Width           =   5295
      End
      Begin VB.TextBox txtUsuario 
         Height          =   615
         Left            =   1200
         TabIndex        =   5
         Top             =   480
         Width           =   3135
      End
      Begin VB.TextBox txtPass 
         Height          =   615
         Left            =   1200
         TabIndex        =   4
         Top             =   1320
         Width           =   3135
      End
      Begin VB.TextBox txtUrl 
         Height          =   735
         Left            =   1200
         TabIndex        =   3
         Top             =   2160
         Width           =   3135
      End
      Begin VB.TextBox txtstampResult 
         Height          =   1695
         Left            =   1200
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   4800
         Width           =   5295
      End
      Begin VB.CommandButton btnStamp 
         Caption         =   "Timbrar"
         Height          =   1215
         Left            =   4680
         TabIndex        =   1
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label8 
         Caption         =   "Resultado"
         Height          =   615
         Left            =   240
         TabIndex        =   20
         Top             =   5400
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Documento a timbrar"
         Height          =   495
         Left            =   120
         TabIndex        =   19
         Top             =   3600
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Usuario:"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Password:"
         Height          =   495
         Left            =   240
         TabIndex        =   8
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "URL:"
         Height          =   615
         Left            =   240
         TabIndex        =   7
         Top             =   2400
         Width           =   1215
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function AuthenticationVB Lib "sw-sdk-cpp.dll" (ByVal Url As String, ByVal User As String, ByVal Pass As String, ByVal Token As String) As Long
Private Declare Function StampByTokenVB Lib "sw-sdk-cpp.dll" (ByVal Url As String, ByVal Token As String, ByVal xml As String, ByVal tfd As String) As Long
Private Declare Function StampVB Lib "sw-sdk-cpp.dll" (ByVal Url As String, ByVal User As String, ByVal Pass As String, ByVal xml As String, ByVal tfd As String) As Long
Private Declare Function StampVBV2 Lib "sw-sdk-cpp.dll" (ByVal Url As String, ByVal User As String, ByVal Pass As String, ByVal xml As String, ByVal tfd As String) As Long
Private Declare Function StampVBV3 Lib "sw-sdk-cpp.dll" (ByVal Url As String, ByVal User As String, ByVal Pass As String, ByVal xml As String, ByVal tfd As String) As Long
Private Declare Function StampVBV4 Lib "sw-sdk-cpp.dll" (ByVal Url As String, ByVal User As String, ByVal Pass As String, ByVal xml As String, ByVal tfd As String) As Long
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()


Private Sub btnSeleccionarFactura_Click()
Dim filename As String
Dim adoStream As ADODB.Stream
Dim var_String As String
CommonDialog1.DialogTitle = "Seleccionar Archivo"
CommonDialog1.ShowOpen
filename = CommonDialog1.filename
Set adoStream = New ADODB.Stream
    adoStream.Charset = "UTF-8"
    adoStream.Open
    adoStream.LoadFromFile CommonDialog1.filename
    var_String = adoStream.ReadText
    txtXmlB64.Text = var_String
End Sub

Private Sub btnOpenXml_Click()
CommonDialog1.ShowOpen
filename = CommonDialog1.filename
    Set adoStream = New ADODB.Stream
    adoStream.Charset = "UTF-8"
    adoStream.Open
    adoStream.LoadFromFile CommonDialog1.filename
    var_String = adoStream.ReadText
    txtXmlB64.Text = var_String
End Sub

Private Sub btnStartTests_Click()
Dim x As Long
Dim path As String
Dim xmlB64(0 To 4) As String
Dim i As Long

path = App.path + "\UnitTest\"
xmlB64(3) = path + "comercioexterior.txt"
xmlB64(0) = path + "conceptos1024.txt"
xmlB64(2) = path + "nomina.txt"
xmlB64(1) = path + "pago10.txt"


MsgBox ("Se ejecutaran las siguientes pruebas" + Chr(10) + xmlB64(0) + Chr(10) + xmlB64(1) + Chr(10) + xmlB64(2) + Chr(10) + xmlB64(3))
Dim y As Long
y = ProgressBar1.Max / 4
y = y / 4
Dim j As Long

For i = 0 To 3
Dim line As String, total As String
Open xmlB64(i) For Input As #1
Do Until EOF(1)
Line Input #1, Linea
total = ""
total = total + Linea + vbCrLf
Loop
Close #1
txtXmlRequest.Text = total
txtXmlResult.Text = StampXml(total, 1)
txtLogs.Text = txtLogs.Text + stringContain(txtXmlResult.Text, "error", xmlB64(i), "1")
ProgressBar1.Value = y * (j + 1)
txtXmlResult.Text = StampXml(total, 2)
txtLogs.Text = txtLogs.Text + stringContain(txtXmlResult.Text, "error", xmlB64(i), "2")
ProgressBar1.Value = y * (j + 2)
txtXmlResult.Text = StampXml(total, 3)
txtLogs.Text = txtLogs.Text + stringContain(txtXmlResult.Text, "error", xmlB64(i), "3")
ProgressBar1.Value = y * (j + 3)
txtXmlResult.Text = StampXml(total, 4)
txtLogs.Text = txtLogs.Text + stringContain(txtXmlResult.Text, "error", xmlB64(i), "4")
ProgressBar1.Value = y * (j + 4)
j = j + 4
Next i


End Sub
Public Function stringContain(ByVal result As String, ByVal contain As String, ByVal path As String, ByVal versionStamp As String) As String
Dim FountIt
FountIt = InStr(1, result, contain)
FountIt2 = InStr(1, result, "Error")
If FountIt <> 0 Or FountIt2 <> 0 Then
    MsgBox ("Error en la versión de timbrado " + versionStamp + result + " en el archivo" + path + " Detalles: " + result)
    stringContain = "Error en la versión de timbrado " + versionStamp + result + " en el archivo" + path + " Detalles: " + result
End If

End Function
Public Function encodeBase64(ByRef arrData() As Byte) As String
   Dim objXML As MSXML2.DOMDocument
   Dim objNode As MSXML2.IXMLDOMElement
   Set objXML = New MSXML2.DOMDocument
   Set objNode = objXML.createElement("b64")
   objNode.dataType = "bin.base64"
   objNode.nodeTypedValue = arrData
   encodeBase64 = objNode.Text
   Set objNode = Nothing
   Set objXML = Nothing
End Function



Private Sub btnToken_Click()
Dim Token As String
Dim User As String
Dim Pass As String
Dim Url As String
Dim nLen As String
Pass = Form1.txtPasswordToken.Text
User = Form1.txtUserToken.Text
Url = Form1.txtUrlToken.Text
Token = Space$(1024)
If Pass = "" Or User = "" Or Url = "" Then
    MsgBox ("Debes tener los datos de Url, usuario y contraseña")
Else
    nLen = AuthenticationVB(Url, User, Pass, Token)
    Form1.txtToken.Text = Token
End If
End Sub

Private Sub btnStamp_Click()
Form1.txtstampResult.Text = StampXml(Form1.txtXmlB64.Text, 4)
End Sub

Public Function StampXml(ByVal xml As String, ByVal version As Integer) As String
Dim User As String
Dim Password As String
Dim Url As String
Dim xml_ As String
Dim nLen As String
Dim tfd As String
tfd = Space$(2000000)
Url = Form1.txtUrl.Text
User = Form1.txtUsuario.Text
Password = Form1.txtPass.Text
xml_ = xml

If Url = "" Or User = "" Or Password = "" Or xml = "" Then
MsgBox ("Debes tener los datos de Url, usuario, Password y Xml")
Else
Select Case version
    Case 1
        nLen = StampVB(Url, User, Password, xml, tfd)
    Case 2
        nLen = StampVBV2(Url, User, Password, xml, tfd)
    Case 3
        nLen = StampVBV3(Url, User, Password, xml, tfd)
    Case 4
        nLen = StampVBV4(Url, User, Password, xml, tfd)
End Select
StampXml = tfd
End If

End Function

Private Sub Form_Initialize()
    InitCommonControls
    ChDir App.path
    Form1.txtUsuario.Text = "demo"
    Form1.txtPass.Text = "123456789"
    Form1.txtUrl.Text = "http://swservicestest-rc.azurewebsites.net"
    Form1.txtPasswordToken.Text = "123456789"
    Form1.txtUserToken.Text = "demo"
    Form1.txtUrlToken.Text = "http://swservicestest-rc.azurewebsites.net"
End Sub

Private Sub Label18_Click()
End Sub

Private Sub Label12_Click()
End Sub
