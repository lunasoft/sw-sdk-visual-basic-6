VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form UNIT_TEST 
   Caption         =   "Form1"
   ClientHeight    =   9180
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15045
   LinkTopic       =   "Form1"
   ScaleHeight     =   9180
   ScaleWidth      =   15045
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   9015
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   15901
      _Version        =   393216
      Tabs            =   5
      Tab             =   2
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Autenticación"
      TabPicture(0)   =   "Form1.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Auth"
      Tab(0).Control(1)=   "body"
      Tab(0).Control(2)=   "Respuesta"
      Tab(0).Control(3)=   "Label1"
      Tab(0).Control(4)=   "Label3"
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Timbrado"
      TabPicture(1)   =   "Form1.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "StampResponse"
      Tab(1).Control(1)=   "XMLATimbrar"
      Tab(1).Control(2)=   "cmbStamp"
      Tab(1).Control(3)=   "Command1"
      Tab(1).Control(4)=   "CommonDialog1"
      Tab(1).Control(5)=   "Label4"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Cancelación"
      TabPicture(2)   =   "Form1.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label6"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label7"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label8"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label9"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label10"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Label11"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Label12"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Label13"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "txtTipoCancelacion"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "btnCancelar"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "txtCer"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "txtResultCancelacion"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "txtKey"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "txtPFX"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "txtPswrdKey"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "txtUUID"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "btnCer"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "btnKey"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "btnPFX"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).Control(19)=   "TipoCancelacion"
      Tab(2).Control(19).Enabled=   0   'False
      Tab(2).Control(20)=   "txtXML"
      Tab(2).Control(20).Enabled=   0   'False
      Tab(2).Control(21)=   "txtRFC"
      Tab(2).Control(21).Enabled=   0   'False
      Tab(2).ControlCount=   22
      TabCaption(3)   =   "Estado de Cuenta"
      TabPicture(3)   =   "Form1.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "AccountBalanceParse"
      Tab(3).Control(1)=   "AccountBalanceResponse"
      Tab(3).Control(2)=   "Command2"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "Válidar"
      TabPicture(4)   =   "Form1.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "ValidateResponse"
      Tab(4).Control(1)=   "ValidateXML"
      Tab(4).Control(2)=   "Command3"
      Tab(4).Control(3)=   "Label5"
      Tab(4).Control(4)=   "XML"
      Tab(4).ControlCount=   5
      Begin VB.TextBox txtRFC 
         Height          =   375
         Left            =   4680
         TabIndex        =   44
         Top             =   2760
         Width           =   3255
      End
      Begin VB.TextBox txtXML 
         Enabled         =   0   'False
         Height          =   615
         Left            =   4080
         MultiLine       =   -1  'True
         TabIndex        =   42
         Text            =   "Form1.frx":008C
         Top             =   5400
         Width           =   10455
      End
      Begin VB.Frame TipoCancelacion 
         Caption         =   "Tipo de Cancelación"
         Height          =   2535
         Left            =   360
         TabIndex        =   36
         Top             =   600
         Width           =   2895
         Begin VB.OptionButton CancelarporPFX 
            Caption         =   "Cancelar por PFX"
            Height          =   375
            Left            =   240
            TabIndex        =   40
            Top             =   720
            Value           =   -1  'True
            Width           =   2055
         End
         Begin VB.OptionButton CancelarPorXML 
            Caption         =   "Cancelar por XML"
            Height          =   375
            Left            =   240
            TabIndex        =   39
            Top             =   1080
            Width           =   2175
         End
         Begin VB.OptionButton CancelarPorCSD 
            Caption         =   "Cancelar por CSD"
            Height          =   375
            Left            =   240
            TabIndex        =   38
            Top             =   1440
            Width           =   1800
         End
         Begin VB.OptionButton CancelarPorUUID 
            Caption         =   "Cancelar por UUID"
            Height          =   375
            Left            =   240
            TabIndex        =   37
            Top             =   1800
            Width           =   2055
         End
      End
      Begin VB.CommandButton btnPFX 
         Caption         =   "Importar .Pfx"
         Height          =   975
         Left            =   13200
         TabIndex        =   35
         Top             =   3360
         Width           =   1335
      End
      Begin VB.CommandButton btnKey 
         Caption         =   "Importar .Key"
         Enabled         =   0   'False
         Height          =   615
         Left            =   13320
         TabIndex        =   34
         Top             =   1800
         Width           =   1335
      End
      Begin VB.CommandButton btnCer 
         Caption         =   "Importar .Cer"
         Enabled         =   0   'False
         Height          =   495
         Left            =   13200
         TabIndex        =   33
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox txtUUID 
         Height          =   375
         Left            =   5400
         TabIndex        =   30
         Top             =   4560
         Width           =   9255
      End
      Begin VB.TextBox txtPswrdKey 
         Height          =   375
         Left            =   11160
         TabIndex        =   27
         Top             =   2760
         Width           =   3495
      End
      Begin VB.TextBox txtPFX 
         Height          =   855
         Left            =   4080
         MultiLine       =   -1  'True
         TabIndex        =   24
         Top             =   3480
         Width           =   8895
      End
      Begin VB.TextBox txtKey 
         Enabled         =   0   'False
         Height          =   735
         Left            =   4080
         MultiLine       =   -1  'True
         TabIndex        =   23
         Top             =   1800
         Width           =   8895
      End
      Begin VB.TextBox txtResultCancelacion 
         Height          =   2295
         Left            =   4080
         MultiLine       =   -1  'True
         TabIndex        =   22
         Top             =   6480
         Width           =   10575
      End
      Begin VB.TextBox txtCer 
         Enabled         =   0   'False
         Height          =   615
         Left            =   4080
         MultiLine       =   -1  'True
         TabIndex        =   21
         Top             =   840
         Width           =   8895
      End
      Begin VB.TextBox ValidateResponse 
         Height          =   4095
         Left            =   -71160
         MultiLine       =   -1  'True
         TabIndex        =   18
         Top             =   3840
         Width           =   8055
      End
      Begin VB.TextBox AccountBalanceParse 
         Height          =   3255
         Left            =   -69720
         MultiLine       =   -1  'True
         TabIndex        =   17
         Top             =   4680
         Width           =   8535
      End
      Begin VB.TextBox AccountBalanceResponse 
         Height          =   2895
         Left            =   -69720
         MultiLine       =   -1  'True
         TabIndex        =   16
         Top             =   840
         Width           =   8655
      End
      Begin VB.TextBox StampResponse 
         Height          =   3135
         Left            =   -70320
         MultiLine       =   -1  'True
         TabIndex        =   15
         Top             =   5520
         Width           =   9495
      End
      Begin VB.TextBox XMLATimbrar 
         Height          =   2895
         Left            =   -70440
         MultiLine       =   -1  'True
         TabIndex        =   14
         Text            =   "Form1.frx":1082
         Top             =   1320
         Width           =   9495
      End
      Begin VB.ComboBox cmbStamp 
         Height          =   315
         ItemData        =   "Form1.frx":1D77
         Left            =   -74640
         List            =   "Form1.frx":1D87
         TabIndex        =   12
         Top             =   1440
         Width           =   2775
      End
      Begin VB.TextBox ValidateXML 
         Height          =   1935
         Left            =   -71280
         MultiLine       =   -1  'True
         TabIndex        =   11
         Text            =   "Form1.frx":1D9B
         Top             =   960
         Width           =   8055
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Validate"
         Height          =   855
         Left            =   -74640
         TabIndex        =   10
         Top             =   840
         Width           =   2775
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Account Balance"
         Height          =   855
         Left            =   -74760
         TabIndex        =   9
         Top             =   960
         Width           =   2775
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar"
         Height          =   975
         Left            =   360
         TabIndex        =   8
         Top             =   7800
         Width           =   3135
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Timbrar"
         Height          =   855
         Left            =   -74640
         TabIndex        =   7
         Top             =   2280
         Width           =   2775
      End
      Begin VB.CommandButton Auth 
         Caption         =   "Autorización"
         Height          =   735
         Left            =   -74760
         TabIndex        =   6
         Top             =   840
         Width           =   2775
      End
      Begin VB.TextBox body 
         Height          =   1335
         Left            =   -67920
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   840
         Width           =   6615
      End
      Begin VB.TextBox Respuesta 
         Height          =   1335
         Left            =   -67920
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   2880
         Width           =   6735
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   -73680
         Top             =   5640
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox txtTipoCancelacion 
         Height          =   375
         Left            =   1080
         TabIndex        =   41
         Text            =   "PFX"
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label13 
         Caption         =   "RFC"
         Height          =   255
         Left            =   4080
         TabIndex        =   45
         Top             =   2880
         Width           =   615
      End
      Begin VB.Label Label12 
         Caption         =   "XML"
         Height          =   255
         Left            =   4080
         TabIndex        =   43
         Top             =   5040
         Width           =   1455
      End
      Begin VB.Label Label11 
         Caption         =   "Resultado"
         Height          =   255
         Left            =   4080
         TabIndex        =   32
         Top             =   6240
         Width           =   2655
      End
      Begin VB.Label Label10 
         Caption         =   "UUID"
         Height          =   255
         Left            =   4200
         TabIndex        =   31
         Top             =   4680
         Width           =   615
      End
      Begin VB.Label Label9 
         Caption         =   "PFX"
         Height          =   255
         Left            =   4080
         TabIndex        =   29
         Top             =   3240
         Width           =   1935
      End
      Begin VB.Label Label8 
         Caption         =   "Contraseña Key"
         Height          =   255
         Left            =   9840
         TabIndex        =   28
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Key"
         Height          =   255
         Left            =   4080
         TabIndex        =   26
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Label Label6 
         Caption         =   "Certificado"
         Height          =   255
         Left            =   4080
         TabIndex        =   25
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label Label5 
         Caption         =   "Resultado Validación"
         Height          =   255
         Left            =   -71160
         TabIndex        =   20
         Top             =   3360
         Width           =   1575
      End
      Begin VB.Label XML 
         Caption         =   "XML"
         Height          =   375
         Left            =   -71160
         TabIndex        =   19
         Top             =   480
         Width           =   3375
      End
      Begin VB.Label Label4 
         Caption         =   "Versión"
         Height          =   375
         Left            =   -74640
         TabIndex        =   13
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Respuesta"
         Height          =   255
         Left            =   -67800
         TabIndex        =   5
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Parseado"
         Height          =   255
         Left            =   -67680
         TabIndex        =   4
         Top             =   2520
         Width           =   975
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Parseado"
      Height          =   855
      Left            =   3480
      TabIndex        =   0
      Top             =   2160
      Width           =   2175
   End
End
Attribute VB_Name = "UNIT_TEST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub auth_Click()
Dim Authentication As New Authentication

body.Text = Authentication.Token("http://services.test.sw.com.mx", "demo", "123456789")

    Dim oScriptEngine As ScriptControl
    Set oScriptEngine = New ScriptControl
    oScriptEngine.Language = "JScript"

    Dim sJsonString As String
    sJsonString = body.Text

    Dim objJSON As Object
    Set objJSON = oScriptEngine.Eval("(" + sJsonString + ")")

Respuesta.Text = VBA.CallByName(VBA.CallByName(objJSON, "data", VbGet), "token", VbGet)

End Sub

Private Sub btnCancelar_Click()
Dim Cancelation As New Cancelation

Dim url As String
url = "http://services.test.sw.com.mx"


Dim Token As String
Token = Respuesta.Text

Dim Result As String

Dim TipoCancelacion As String
TipoCancelacion = txtTipoCancelacion.Text

Dim Cer As String
Cer = txtCer.Text

Dim Key As String
Key = txtKey.Text

Dim RFC As String
RFC = txtRFC.Text

Dim password As String
password = txtPswrdKey.Text

Dim PFX As String
PFX = txtPFX.Text

Dim UUID As String
UUID = txtUUID.Text

Dim XML As String
XML = txtXML.Text

txtResultCancelacion.Text = ""


MsgBox (TipoCancelacion)

If TipoCancelacion = "UUID" Then

Result = Cancelation.CancelationByUUID(url, RFC, UUID, Token)

ElseIf TipoCancelacion = "PFX" Then

Result = Cancelation.CancelationByPFX(url, PFX, UUID, password, RFC, Token)

ElseIf TipoCancelacion = "CSD" Then

Result = Cancelation.CancelationByCSD(url, Token, Cer, Key, password, UUID)

ElseIf TipoCancelacion = "XML" Then

Result = Cancelation.CancelationByXML(url, XML, Token)


End If

txtResultCancelacion.Text = Result

End Sub

Private Sub btnCer_Click()

   Dim mybase64 As String
    CommonDialog1.Filter = "Cer files (*.cer)|*.cer"
    CommonDialog1.ShowOpen
    If CommonDialog1.FileName <> "" Then
    Open CommonDialog1.FileName For Input As #1
    Do
    Input #1, linetext
    txtCer.Text = txtCer.Text & linetext
    Loop Until EOF(1)
    End If
    Close #1
    

    strAttachment = Base64EncodeString(txtCer.Text)
    txtCer.Text = strAttachment
End Sub

Private Sub btnKey_Click()
 Dim mybase64 As String
    CommonDialog1.Filter = "Key files (*.key)|*.key"
    CommonDialog1.ShowOpen
    If CommonDialog1.FileName <> "" Then
    Open CommonDialog1.FileName For Input As #1
    Do
    Input #1, linetext
    txtKey.Text = txtKey.Text & linetext
    Loop Until EOF(1)
    End If
    Close #1
    

    strAttachment = Base64EncodeString(txtKey.Text)
    txtKey.Text = strAttachment
End Sub

Private Sub btnPFX_Click()
Dim mybase64 As String
    CommonDialog1.Filter = "Pfx files (*.pfx)|*.pfx"
    CommonDialog1.ShowOpen
    If CommonDialog1.FileName <> "" Then
    Open CommonDialog1.FileName For Input As #1
    Do
    Input #1, linetext
    txtPFX.Text = txtPFX.Text & linetext
    Loop Until EOF(1)
    End If
    Close #1
    

    strAttachment = Base64EncodeString(txtPFX.Text)
    txtPFX.Text = strAttachment
End Sub

Private Sub btnXML_Click()
CommonDialog1.Filter = "XML files (*.xml)|*.xml"
    CommonDialog1.ShowOpen
    If CommonDialog1.FileName <> "" Then
    Open CommonDialog1.FileName For Input As #1
    Do
    Input #1, linetext
    txtXML.Text = txtXML.Text & linetext
    Loop Until EOF(1)
    End If
    Close #1
    
    txtXML.Text = UTF8ENCODE(txtXML.Text)
    
End Sub

Private Sub CancelarPorCSD_Click()

txtCer.Enabled = True
txtCer.Text = ""
btnCer.Enabled = True

txtKey.Enabled = True
txtKey.Text = ""
btnKey.Enabled = True

txtPswrdKey.Enabled = True
txtPswrdKey.Text = ""

txtUUID.Enabled = False
txtUUID.Text = ""

txtXML.Enabled = False

txtPFX.Enabled = False
txtPFX.Text = ""
btnPFX.Enabled = False

txtResultCancelacion.Text = ""
txtTipoCancelacion.Text = "CSD"

End Sub

Private Sub CancelarporPFX_Click()

txtCer.Enabled = False
txtCer.Text = ""
btnCer.Enabled = False

txtKey.Enabled = False
txtKey.Text = ""
btnKey.Enabled = False

txtPswrdKey.Enabled = True
txtPswrdKey.Text = ""

txtUUID.Enabled = True
txtUUID.Text = ""

txtXML.Enabled = False

txtPFX.Enabled = True
txtPFX.Text = ""


txtRFC.Enabled = True
txtRFC.Text = ""

txtResultCancelacion.Text = ""
txtTipoCancelacion.Text = "PFX"

End Sub


Private Sub CancelarPorUUID_Click()

txtCer.Enabled = False
txtCer.Text = ""
btnCer.Enabled = False

txtKey.Enabled = False
txtKey.Text = ""
btnKey.Enabled = False

txtPswrdKey.Enabled = False
txtPswrdKey.Text = ""

txtUUID.Enabled = True
txtUUID.Text = ""

txtXML.Enabled = False

txtPFX.Enabled = False
txtPFX.Text = ""
btnPFX.Enabled = False

txtResultCancelacion.Text = ""
txtTipoCancelacion.Text = "UUID"

End Sub

Private Sub CancelarPorXML_Click()

txtCer.Enabled = False
txtCer.Text = ""
btnCer.Enabled = False

txtKey.Enabled = False
txtKey.Text = ""
btnKey.Enabled = False

txtPswrdKey.Enabled = False
txtPswrdKey.Text = ""

txtUUID.Enabled = False
txtUUID.Text = ""

txtXML.Enabled = True

txtPFX.Enabled = False
txtPFX.Text = ""
btnPFX.Enabled = False

txtResultCancelacion.Text = ""
txtTipoCancelacion.Text = "XML"

End Sub

Private Sub Command1_Click()
Dim Stamp As New Stamp
Dim XML As String
Dim Token As String
Dim url As String
Dim Result As String

Token = Respuesta.Text
url = "http://services.test.sw.com.mx"
XML = XMLATimbrar.Text

If cmbStamp.Text = "V1" Then
    Result = Stamp.StampV1(url, XML, Token)
ElseIf cmbStamp.Text = "V2" Then
    Result = Stamp.StampV2(url, XML, Token)
ElseIf cmbStamp.Text = "V3" Then
    Result = Stamp.StampV3(url, XML, Token)
ElseIf cmbStamp.Text = "V4" Then
    Result = Stamp.StampV4(url, XML, Token)
End If

StampResponse.Text = Result
End Sub
Function FileToString(strFilename As String) As String
  iFile = FreeFile
  Open strFilename For Input As #iFile
    FileToString = StrConv(InputB(LOF(iFile), iFile), vbUnicode)
  Close #iFile
End Function

Private Sub Command10_Click()
    Dim mybase64 As String
    CommonDialog1.Filter = "Pfx files (*.pfx)|*.pfx"
    CommonDialog1.ShowOpen
    If CommonDialog1.FileName <> "" Then
    Open CommonDialog1.FileName For Input As #1
    Do
    Input #1, linetext
    txtPFX.Text = txtPFX.Text & linetext
    Loop Until EOF(1)
    End If
    Close #1

    strAttachment = Base64EncodeString(txtPFX.Text)
    txtPFX.Text = strAttachment
End Sub

Private Sub Command2_Click()
Dim Token As String
Dim url As String

Token = Respuesta.Text
url = "http://services.test.sw.com.mx"
Dim AccountBalance As New AccountBalance

AccountBalanceResponse.Text = AccountBalance.AccountBalance(url, Token)
End Sub

Private Sub Command3_Click()
Dim Validate As New Validate
Dim url As String
Dim Token As String
Dim XML As String

url = "http://services.test.sw.com.mx"
Token = Respuesta.Text
XML = ValidateXML.Text

ValidateResponse.Text = Validate.Validate(url, Token, XML)


End Sub

Function FToString(strFilename As String) As String
  
  Open strFilename For Input As #iFile
    FileToString = StrConv(InputB(LOF(iFile), iFile), vbUnicode)
  Close #iFile
  
End Function

Private Sub Command4_Click()



End Sub

Private Sub Command5_Click()

    XMLATimbrar.Text = ""
    CommonDialog1.Filter = "XML files (*.xml)|*.xml"
    CommonDialog1.ShowOpen
    If CommonDialog1.FileName <> "" Then
    Open CommonDialog1.FileName For Input As #1
    Do
    Input #1, linetext
    XMLATimbrar.Text = XMLATimbrar.Text & linetext
    Loop Until EOF(1)
    End If
    Close #1

    XMLATimbrar.Text = UTF8ENCODE(XMLATimbrar.Text)
    
End Sub

Private Sub Command6_Click()

    CommonDialog1.Filter = "XML files (*.xml)|*.xml"
    CommonDialog1.ShowOpen
    If CommonDialog1.FileName <> "" Then
    Open CommonDialog1.FileName For Input As #1
    Do
    Input #1, linetext
    ValidateXML.Text = ValidateXML.Text & linetext
    Loop Until EOF(1)
    End If
    Close #1
    
    
    ValidateXML.Text = UTF8ENCODE(ValidateXML.Text)
  
 
End Function
 


Private Sub Command8_Click()

   Dim mybase64 As String
    CommonDialog1.Filter = "Cer files (*.cer)|*.cer"
    CommonDialog1.ShowOpen
    If CommonDialog1.FileName <> "" Then
    Open CommonDialog1.FileName For Input As #1
    Do
    Input #1, linetext
    txtCer.Text = txtCer.Text & linetext
    Loop Until EOF(1)
    End If
    Close #1
    

    strAttachment = Base64EncodeString(txtCer.Text)
    txtCer.Text = strAttachment
    
End Sub


Public Function UTF8ENCODE(ByVal sStr As String) As String

  For l& = 1 To Len(sStr)
        lChar& = AscW(Mid(sStr, l&, 1))
        If lChar& < 128 Then
        sUtf8$ = sUtf8$ + Mid(sStr, l&, 1)
    ElseIf ((lChar& > 127) And (lChar& < 2048)) Then
        sUtf8$ = sUtf8$ + Chr(((lChar& \ 64) Or 192))
        sUtf8$ = sUtf8$ + Chr(((lChar& And 63) Or 128))
    Else
        sUtf8$ = sUtf8$ + Chr(((lChar& \ 144) Or 234))
        sUtf8$ = sUtf8$ + Chr((((lChar& \ 64) And 63) Or 128))
        sUtf8$ = sUtf8$ + Chr(((lChar& And 63) Or 128))
    End If
  Next l&
  
    UTF8ENCODE = sUtf8$

End Function

Private Sub Command9_Click()
 Dim mybase64 As String
    CommonDialog1.Filter = "Key files (*.key)|*.key"
    CommonDialog1.ShowOpen
    If CommonDialog1.FileName <> "" Then
    Open CommonDialog1.FileName For Input As #1
    Do
    Input #1, linetext
    txtKey.Text = txtKey.Text & linetext
    Loop Until EOF(1)
    End If
    Close #1
    

    strAttachment = Base64EncodeString(txtKey.Text)
    txtKey.Text = strAttachment
End Sub

