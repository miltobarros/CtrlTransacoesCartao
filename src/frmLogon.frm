VERSION 5.00
Begin VB.Form frmLogon 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Logon"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4950
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogon.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4950
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSenha 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   3120
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1260
      Width           =   1635
   End
   Begin VB.TextBox txtUsuario 
      Height          =   285
      Left            =   3120
      MaxLength       =   15
      TabIndex        =   0
      Top             =   750
      Width           =   1635
   End
   Begin VB.Frame fraLinha 
      Height          =   30
      Left            =   2220
      TabIndex        =   11
      Top             =   1890
      Width           =   2445
   End
   Begin VB.PictureBox pctCadeado 
      AutoSize        =   -1  'True
      Height          =   2865
      Left            =   60
      Picture         =   "frmLogon.frx":030A
      ScaleHeight     =   2805
      ScaleWidth      =   1860
      TabIndex        =   2
      Top             =   120
      Width           =   1920
   End
   Begin VB.CommandButton cmdCancelar 
      Appearance      =   0  'Flat
      Caption         =   "&Cancelar"
      Height          =   495
      Left            =   3660
      TabIndex        =   10
      Top             =   2160
      Width           =   1155
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   495
      Left            =   2160
      TabIndex        =   9
      Top             =   2160
      Width           =   1155
   End
   Begin VB.Label lblUsuario 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "USUÁRIO:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   1
      Left            =   2100
      TabIndex        =   6
      Top             =   810
      Width           =   930
   End
   Begin VB.Label lblSenha 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "SENHA:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   0
      Left            =   2280
      TabIndex        =   7
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label lblSenha 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SENHA:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   240
      Index           =   1
      Left            =   2310
      TabIndex        =   8
      Top             =   1350
      Width           =   735
   End
   Begin VB.Label lblUsuario 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "USUÁRIO:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   240
      Index           =   0
      Left            =   2130
      TabIndex        =   5
      Top             =   840
      Width           =   930
   End
   Begin VB.Label lblCabecalho 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Banco de Teste"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   360
      Index           =   0
      Left            =   2280
      TabIndex        =   3
      Top             =   210
      Width           =   2220
   End
   Begin VB.Label lblCabecalho 
      AutoSize        =   -1  'True
      Caption         =   "Banco de Teste"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Index           =   1
      Left            =   2250
      TabIndex        =   4
      Top             =   240
      Width           =   2220
   End
End
Attribute VB_Name = "frmLogon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private bOk As Boolean

Private Sub cmdCancelar_Click()
    Unload Me

End Sub

Private Sub cmdOK_Click()
    On Error GoTo Trata_Erro
    bOk = True
    If UCase(txtSenha.Text) <> "" And UCase(txtUsuario.Text) <> "" Then
        'Conecta com um usuário Existente no Banco de Dados
        Set dbBancoDados = CreateObject("ADODB.Connection")
        'dbBancoDados.ConnectionString = ("DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & G_Servidor & ";PORT=3306;DATABASE=" & G_NomeBanco & ";USER=" & txtUsuario.Text & ";PASSWORD=" & txtSenha.Text)
        dbBancoDados.ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & G_NomeBanco & ";Data Source=" & G_Servidor
        
        dbBancoDados.ConnectionTimeout = 60
        dbBancoDados.CommandTimeout = 400
        dbBancoDados.CursorLocation = adUseClient
        dbBancoDados.Open

        If dbBancoDados.State = 1 Then
            If UCase$(txtUsuario.Text) <> UCase$("desafio") And txtSenha.Text <> "desafio" Then
                dbBancoDados.Close
                GoTo Trata_Erro
            End If
        Else
            MsgBox "Erro na abertura do Banco de Dados!", vbInformation
            End
        End If
        
        Unload Me
    End If
    Exit Sub

Trata_Erro:
    bOk = False
    MsgBox "Usuário ou Senha Inválida.", vbInformation, GC_NomeSistema
    txtUsuario.SetFocus

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Call PG_Muda_Teclas(KeyAscii)

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case FG_Tecla_Atalho(KeyCode, Shift)
        Case "ESC" ' Fechar
            End

        Case "F1" 'Ajuda

    End Select

End Sub

Private Sub Form_Load()
  Me.Top = Screen.Height / 2 - Me.Height / 2
  Me.Left = Screen.Width / 2 - Me.Width / 2
  lblCabecalho(0).Caption = GC_NomeSistema
  lblCabecalho(1).Caption = lblCabecalho(0).Caption

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not bOk Then
        End
    End If

End Sub

Private Sub txtSenha_GotFocus()
    txtSenha.SelStart = 0
    txtSenha.SelLength = Len(txtSenha)

End Sub

Private Sub txtSenha_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call cmdOK_Click

End Sub

Private Sub txtUsuario_GotFocus()
    txtUsuario.SelStart = 0
    txtUsuario.SelLength = Len(txtUsuario)

End Sub
