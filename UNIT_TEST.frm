VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form UNIT_TEST 
   Caption         =   "Form1"
   ClientHeight    =   8970
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14970
   LinkTopic       =   "Form1"
   ScaleHeight     =   8970
   ScaleWidth      =   14970
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   9015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   15901
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Autenticación"
      TabPicture(0)   =   "UNIT_TEST.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Auth"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "body"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Respuesta"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Timbrado"
      TabPicture(1)   =   "UNIT_TEST.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label4"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "CommonDialog1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmbStamp"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Command1"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "XMLATimbrar"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "StampResponse"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Command4"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "Cancelación"
      TabPicture(2)   =   "UNIT_TEST.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label13"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label12"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label11"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label10"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label9"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Label8"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Label7"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Label6"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "txtTipoCancelacion"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "txtRFC"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "TipoCancelacion"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "btnPFX"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "btnKey"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "btnCer"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "txtUUID"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "txtPswrdKey"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "txtPFX"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "txtKey"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "txtResultCancelacion"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).Control(19)=   "txtCer"
      Tab(2).Control(19).Enabled=   0   'False
      Tab(2).Control(20)=   "btnCancelar"
      Tab(2).Control(20).Enabled=   0   'False
      Tab(2).Control(21)=   "txtXML"
      Tab(2).Control(21).Enabled=   0   'False
      Tab(2).Control(22)=   "Command5"
      Tab(2).Control(22).Enabled=   0   'False
      Tab(2).ControlCount=   23
      TabCaption(3)   =   "Estado de Cuenta"
      TabPicture(3)   =   "UNIT_TEST.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "AccountBalanceParse"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "AccountBalanceResponse"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Command2"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "Válidar"
      TabPicture(4)   =   "UNIT_TEST.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label5"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "XML"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "ValidateResponse"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "ValidateXML"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "Command3"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).ControlCount=   5
      Begin VB.CommandButton Command5 
         Caption         =   "Importar XML"
         Height          =   735
         Left            =   -61440
         TabIndex        =   46
         Top             =   5400
         Width           =   1215
      End
      Begin RichTextLib.RichTextBox txtXML 
         Height          =   735
         Left            =   -70920
         TabIndex        =   45
         Top             =   5400
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   1296
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"UNIT_TEST.frx":008C
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Importar XML"
         Height          =   975
         Left            =   -74520
         TabIndex        =   44
         Top             =   2160
         Width           =   2775
      End
      Begin RichTextLib.RichTextBox StampResponse 
         Height          =   3855
         Left            =   -71520
         TabIndex        =   43
         Top             =   4800
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   6800
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"UNIT_TEST.frx":010E
      End
      Begin RichTextLib.RichTextBox XMLATimbrar 
         Height          =   3375
         Left            =   -71520
         TabIndex        =   42
         Top             =   1080
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   5953
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"UNIT_TEST.frx":0190
      End
      Begin VB.TextBox Respuesta 
         Height          =   1335
         Left            =   3840
         MultiLine       =   -1  'True
         TabIndex        =   27
         Top             =   2880
         Width           =   9975
      End
      Begin VB.TextBox body 
         Height          =   1335
         Left            =   3720
         MultiLine       =   -1  'True
         TabIndex        =   26
         Top             =   840
         Width           =   9975
      End
      Begin VB.CommandButton Auth 
         Caption         =   "Autorización"
         Height          =   735
         Left            =   240
         TabIndex        =   25
         Top             =   840
         Width           =   2775
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Timbrar"
         Height          =   855
         Left            =   -74760
         TabIndex        =   24
         Top             =   6240
         Width           =   2775
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar"
         Height          =   975
         Left            =   -74640
         TabIndex        =   23
         Top             =   7800
         Width           =   3135
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Account Balance"
         Height          =   855
         Left            =   -74760
         TabIndex        =   22
         Top             =   960
         Width           =   2775
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Validate"
         Height          =   855
         Left            =   -74640
         TabIndex        =   21
         Top             =   840
         Width           =   2775
      End
      Begin VB.TextBox ValidateXML 
         Height          =   1935
         Left            =   -71280
         MultiLine       =   -1  'True
         TabIndex        =   20
         Text            =   "UNIT_TEST.frx":0212
         Top             =   960
         Width           =   8055
      End
      Begin VB.ComboBox cmbStamp 
         Height          =   315
         ItemData        =   "UNIT_TEST.frx":0F51
         Left            =   -74640
         List            =   "UNIT_TEST.frx":0F61
         TabIndex        =   19
         Top             =   1440
         Width           =   2775
      End
      Begin VB.TextBox AccountBalanceResponse 
         Height          =   2895
         Left            =   -69720
         MultiLine       =   -1  'True
         TabIndex        =   18
         Top             =   840
         Width           =   8655
      End
      Begin VB.TextBox AccountBalanceParse 
         Height          =   3255
         Left            =   -69720
         MultiLine       =   -1  'True
         TabIndex        =   17
         Top             =   4680
         Width           =   8535
      End
      Begin VB.TextBox ValidateResponse 
         Height          =   4095
         Left            =   -71160
         MultiLine       =   -1  'True
         TabIndex        =   16
         Top             =   3840
         Width           =   8055
      End
      Begin VB.TextBox txtCer 
         Enabled         =   0   'False
         Height          =   615
         Left            =   -70920
         MultiLine       =   -1  'True
         TabIndex        =   15
         Top             =   840
         Width           =   8895
      End
      Begin VB.TextBox txtResultCancelacion 
         Height          =   2295
         Left            =   -70920
         MultiLine       =   -1  'True
         TabIndex        =   14
         Top             =   6480
         Width           =   10575
      End
      Begin VB.TextBox txtKey 
         Enabled         =   0   'False
         Height          =   735
         Left            =   -70920
         MultiLine       =   -1  'True
         TabIndex        =   13
         Top             =   1800
         Width           =   8895
      End
      Begin VB.TextBox txtPFX 
         Height          =   855
         Left            =   -70920
         MultiLine       =   -1  'True
         TabIndex        =   12
         Top             =   3480
         Width           =   8895
      End
      Begin VB.TextBox txtPswrdKey 
         Height          =   375
         Left            =   -63840
         TabIndex        =   11
         Top             =   2760
         Width           =   3495
      End
      Begin VB.TextBox txtUUID 
         Height          =   375
         Left            =   -69600
         TabIndex        =   10
         Top             =   4560
         Width           =   9255
      End
      Begin VB.CommandButton btnCer 
         Caption         =   "Importar .Cer"
         Enabled         =   0   'False
         Height          =   495
         Left            =   -61800
         TabIndex        =   9
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton btnKey 
         Caption         =   "Importar .Key"
         Enabled         =   0   'False
         Height          =   615
         Left            =   -61680
         TabIndex        =   8
         Top             =   1800
         Width           =   1335
      End
      Begin VB.CommandButton btnPFX 
         Caption         =   "Importar .Pfx"
         Height          =   975
         Left            =   -61800
         TabIndex        =   7
         Top             =   3360
         Width           =   1335
      End
      Begin VB.Frame TipoCancelacion 
         Caption         =   "Tipo de Cancelación"
         Height          =   2535
         Left            =   -74640
         TabIndex        =   2
         Top             =   600
         Width           =   2895
         Begin VB.OptionButton CancelarPorUUID 
            Caption         =   "Cancelar por UUID"
            Height          =   375
            Left            =   240
            TabIndex        =   6
            Top             =   1800
            Width           =   2055
         End
         Begin VB.OptionButton CancelarPorCSD 
            Caption         =   "Cancelar por CSD"
            Height          =   375
            Left            =   240
            TabIndex        =   5
            Top             =   1440
            Width           =   1800
         End
         Begin VB.OptionButton CancelarPorXML 
            Caption         =   "Cancelar por XML"
            Height          =   375
            Left            =   240
            TabIndex        =   4
            Top             =   1080
            Width           =   2175
         End
         Begin VB.OptionButton CancelarporPFX 
            Caption         =   "Cancelar por PFX"
            Height          =   375
            Left            =   240
            TabIndex        =   3
            Top             =   720
            Value           =   -1  'True
            Width           =   2055
         End
      End
      Begin VB.TextBox txtRFC 
         Height          =   375
         Left            =   -70320
         TabIndex        =   1
         Top             =   2760
         Width           =   3255
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
         Left            =   -73560
         TabIndex        =   28
         Text            =   "PFX"
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Parseado"
         Height          =   255
         Left            =   3960
         TabIndex        =   41
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Respuesta"
         Height          =   255
         Left            =   3720
         TabIndex        =   40
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Versión"
         Height          =   375
         Left            =   -74640
         TabIndex        =   39
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label XML 
         Caption         =   "XML"
         Height          =   375
         Left            =   -71160
         TabIndex        =   38
         Top             =   480
         Width           =   3375
      End
      Begin VB.Label Label5 
         Caption         =   "Resultado Validación"
         Height          =   255
         Left            =   -71160
         TabIndex        =   37
         Top             =   3360
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Certificado"
         Height          =   255
         Left            =   -70920
         TabIndex        =   36
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label Label7 
         Caption         =   "Key"
         Height          =   255
         Left            =   -70920
         TabIndex        =   35
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Label Label8 
         Caption         =   "Contraseña Key"
         Height          =   255
         Left            =   -65160
         TabIndex        =   34
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "PFX"
         Height          =   255
         Left            =   -70920
         TabIndex        =   33
         Top             =   3240
         Width           =   1935
      End
      Begin VB.Label Label10 
         Caption         =   "UUID"
         Height          =   255
         Left            =   -70800
         TabIndex        =   32
         Top             =   4680
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "Resultado"
         Height          =   255
         Left            =   -70920
         TabIndex        =   31
         Top             =   6240
         Width           =   2655
      End
      Begin VB.Label Label12 
         Caption         =   "XML"
         Height          =   255
         Left            =   -70920
         TabIndex        =   30
         Top             =   5040
         Width           =   1455
      End
      Begin VB.Label Label13 
         Caption         =   "RFC"
         Height          =   255
         Left            =   -70920
         TabIndex        =   29
         Top             =   2880
         Width           =   615
      End
   End
End
Attribute VB_Name = "UNIT_TEST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Auth_Click()
Dim Authentication As New Authentication

body.Text = Authentication.Token("http://services.test.sw.com.mx", "demo", "123456789")
' En caso de que presente un error el Script Control procederemos a agregarlo:
' Project/Components/Microsoft Script Control 1.0

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

Private Sub Command2_Click()
Dim AccountBalance As New AccountBalance
Dim url As String
Dim Token As String
Dim XML As String

url = "http://services.test.sw.com.mx"
Token = Respuesta.Text

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

Private Sub Command4_Click()
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
    
    txtXML.Text = XMLATimbrar.Text
End Sub

Private Sub Command5_Click()
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
    
    txtXML.Text = txtXML.Text


End Sub
