VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.OCX"
Begin VB.MDIForm mdiPrincipal 
   BackColor       =   &H8000000C&
   Caption         =   "Gerenciamento de Transações de Cartão de Crédito"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   5460
   Icon            =   "mdiPrincipal.frx":0000
   LinkTopic       =   "MDIForm1"
   Moveable        =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog cdlDialogo 
      Left            =   2940
      Top             =   2130
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgMDI 
      Left            =   1440
      Top             =   2100
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiPrincipal.frx":08CA
            Key             =   "TRANSACAO"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbBotoes 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5460
      _ExtentX        =   9631
      _ExtentY        =   1058
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      Style           =   1
      ImageList       =   "imgMDI"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "TRANSACAO"
            Description     =   "Cadastro de Transações de Cartão de Crédito"
            Object.ToolTipText     =   "Cadastro de Transações de Cartão de Crédito"
            ImageIndex      =   1
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgCadastros 
      Left            =   2070
      Top             =   2100
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiPrincipal.frx":0BE4
            Key             =   "Novo"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiPrincipal.frx":135E
            Key             =   "Salvar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiPrincipal.frx":1AD8
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiPrincipal.frx":2252
            Key             =   "Procurar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiPrincipal.frx":29CC
            Key             =   "Imprimir"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiPrincipal.frx":3146
            Key             =   "Fechar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiPrincipal.frx":38C0
            Key             =   "Ajuda"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiPrincipal.frx":403A
            Key             =   "Selecionar"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuSistema 
      Caption         =   "&Sistema"
      Begin VB.Menu mnuSair 
         Caption         =   "Sai&r"
      End
   End
   Begin VB.Menu mnuCadastro 
      Caption         =   "&Cadastro"
      Begin VB.Menu itmCadastro 
         Caption         =   "&Transações"
         Index           =   2
      End
   End
   Begin VB.Menu mnuRelatorio 
      Caption         =   "&Relatório"
   End
   Begin VB.Menu mnuJanela 
      Caption         =   "&Janela"
      WindowList      =   -1  'True
   End
End
Attribute VB_Name = "mdiPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private i As Integer

Private Sub itmCadastro_Click(Index As Integer)
    Select Case FG_Retirar_Atalho(itmCadastro(Index).Caption)

    Case "TRANSACOES"
        With frmTransacaoCartao
            If FG_Form_Carregado(frmTransacaoCartao) Then
                .SetFocus
            Else
                .Show
            End If
        End With

    End Select

End Sub

Private Sub MDIForm_Load()
    'Configura Imagens da Barra de buttons
    mdiPrincipal.Show
    frmLogon.Show vbModal

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    dbBancoDados.Close

End Sub

Private Sub mnuSair_Click()
    Unload Me

End Sub


Private Sub tlbBotoes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case Is = "TRANSACAO"
            Call itmCadastro_Click(2)

    End Select

End Sub
