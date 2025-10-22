VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmLancamentoRubrica 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lançamento de Rubrica"
   ClientHeight    =   7005
   ClientLeft      =   210
   ClientTop       =   915
   ClientWidth     =   11565
   Icon            =   "frmLancamentoRubrica.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   11565
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshDigitacao 
      Height          =   555
      Left            =   750
      TabIndex        =   32
      Top             =   4620
      Visible         =   0   'False
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   979
      _Version        =   393216
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.ComboBox cmbFolha 
      Height          =   315
      Left            =   7140
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   600
      Width           =   1515
   End
   Begin VB.PictureBox pctRH 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   9600
      ScaleHeight     =   315
      ScaleWidth      =   1845
      TabIndex        =   38
      Top             =   105
      Width           =   1875
      Begin VB.Label lblRH 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RH - UniCEUMA"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   270
         Left            =   120
         TabIndex        =   39
         Top             =   30
         Width           =   1635
      End
   End
   Begin VB.CheckBox chkNaoCalcRub 
      Caption         =   "Salvar sem calcular rubrica"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   8850
      TabIndex        =   37
      Top             =   630
      Width           =   2655
   End
   Begin MSMask.MaskEdBox mskReferencia 
      Height          =   285
      Left            =   1050
      TabIndex        =   2
      Top             =   600
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   7
      Format          =   "mm/yyyy"
      Mask            =   "##/####"
      PromptChar      =   "_"
   End
   Begin VB.Frame fraRubrica 
      Caption         =   "Evento"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   120
      TabIndex        =   26
      Top             =   3240
      Width           =   11385
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshEvento 
         Height          =   3015
         Left            =   120
         TabIndex        =   31
         Top             =   630
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   5318
         _Version        =   393216
         FixedCols       =   0
         BackColorFixed  =   8438015
         ForeColorFixed  =   0
         BackColorSel    =   8421376
         BackColorBkg    =   12640511
         SelectionMode   =   1
         GridLineWidthFixed=   1
         FormatString    =   "Evento |Valor"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.TextBox txtValor 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   10110
         MaxLength       =   10
         TabIndex        =   30
         Top             =   270
         Width           =   945
      End
      Begin VB.ComboBox cmbEvento 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   810
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   210
         Width           =   8175
      End
      Begin VB.Label lblValor 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor :"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   9600
         TabIndex        =   29
         Top             =   300
         Width           =   450
      End
      Begin VB.Label lblEvento 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Evento :"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   165
         TabIndex        =   27
         Top             =   300
         Width           =   600
      End
   End
   Begin VB.Frame fraLotacao 
      Caption         =   "Lotação"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1365
      Left            =   120
      TabIndex        =   15
      Top             =   1800
      Width           =   11385
      Begin VB.TextBox txtStatusFolha 
         BackColor       =   &H00808000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   6660
         TabIndex        =   35
         Top             =   990
         Width           =   4635
      End
      Begin VB.TextBox txtDtSituacao 
         BackColor       =   &H00808000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1095
         TabIndex        =   34
         Top             =   990
         Width           =   1245
      End
      Begin VB.TextBox txtDtAdmissao 
         BackColor       =   &H00808000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3615
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   33
         Top             =   990
         Width           =   1245
      End
      Begin VB.ComboBox cmbSituacao 
         BackColor       =   &H00808000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         ItemData        =   "frmLancamentoRubrica.frx":0E42
         Left            =   6675
         List            =   "frmLancamentoRubrica.frx":0E4F
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   23
         ToolTipText     =   "Estado civil do funcionário"
         Top             =   630
         Width           =   3255
      End
      Begin VB.ComboBox CmbLotacao 
         Height          =   315
         Left            =   6675
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   270
         Width           =   4620
      End
      Begin VB.ComboBox CmbOcupacao 
         BackColor       =   &H00808000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1095
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   630
         Width           =   4545
      End
      Begin VB.ComboBox cmbEmpresa 
         Height          =   315
         Left            =   1095
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   270
         Width           =   4545
      End
      Begin VB.Label lblStatusFolha 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status da Folha :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   195
         Left            =   5160
         TabIndex        =   36
         Top             =   1035
         Width           =   1470
      End
      Begin VB.Label lblAdmissao 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Admissão :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   195
         Left            =   2640
         TabIndex        =   25
         Top             =   1035
         Width           =   930
      End
      Begin VB.Label lblSituacao 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Situação :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   195
         Left            =   5745
         TabIndex        =   22
         Top             =   690
         Width           =   885
      End
      Begin VB.Label lblData 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   195
         Left            =   510
         TabIndex        =   24
         Top             =   1035
         Width           =   540
      End
      Begin VB.Label lblOcupacao 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ocupação :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   195
         Left            =   60
         TabIndex        =   20
         Top             =   690
         Width           =   1005
      End
      Begin VB.Label lblLotacao 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lotação :"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   5955
         TabIndex        =   18
         Top             =   285
         Width           =   675
      End
      Begin VB.Label lblEmpresa 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empresa :"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   330
         TabIndex        =   16
         Top             =   315
         Width           =   705
      End
   End
   Begin VB.Frame fraFuncionario 
      Caption         =   "Identificação do Funcionário"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   11385
      Begin VB.Frame fraPosicao 
         Caption         =   "Posição"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   9000
         TabIndex        =   13
         Top             =   120
         Visible         =   0   'False
         Width           =   2055
         Begin VB.Label lblPosicao 
            Caption         =   "Posição 1 de 1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   240
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.TextBox txtMatricula 
         Height          =   285
         Left            =   945
         MaxLength       =   5
         TabIndex        =   10
         Top             =   270
         Width           =   795
      End
      Begin VB.TextBox txtNome 
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2415
         MaxLength       =   100
         TabIndex        =   12
         Top             =   270
         Width           =   6390
      End
      Begin VB.Label lblNome 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nome :"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1830
         TabIndex        =   11
         Top             =   300
         Width           =   555
      End
      Begin VB.Label lblMatricula 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Matrícula :"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   150
         TabIndex        =   9
         Top             =   300
         Width           =   765
      End
   End
   Begin MSComctlLib.Toolbar tlbBotoes 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11565
      _ExtentX        =   20399
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "NOVO"
            Object.ToolTipText     =   "Novo"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ABRIR"
            Object.ToolTipText     =   "Procurar (Pressione F2)"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SALVAR"
            Object.ToolTipText     =   "Salvar"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "EXCLUIR"
            Object.ToolTipText     =   "Excluir"
            Object.Width           =   5000
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   400
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "DIGITAR"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   4000
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "AJUDA"
            Object.ToolTipText     =   "Ajuda"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SAIR"
            Object.ToolTipText     =   "Fechar a tela"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSMask.MaskEdBox mskBaseCalculo 
      Height          =   285
      Left            =   5490
      TabIndex        =   5
      Top             =   600
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      MaxLength       =   7
      Format          =   "mm/yyyy"
      Mask            =   "##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label lblBaseCalculo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Base de Cálculo :"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   4185
      TabIndex        =   4
      Top             =   660
      Width           =   1245
   End
   Begin VB.Label lbl13Salario 
      AutoSize        =   -1  'True
      BackColor       =   &H00000080&
      Caption         =   "13° Salário (1ª Parcela)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   2040
      TabIndex        =   3
      Top             =   630
      Visible         =   0   'False
      Width           =   1965
   End
   Begin VB.Label lblFolha 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Folha :"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   6615
      TabIndex        =   6
      Top             =   645
      Width           =   480
   End
   Begin VB.Label lblReferencia 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Referência :"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   660
      Width           =   870
   End
End
Attribute VB_Name = "frmLancamentoRubrica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private sQuery          As String
Private sQueryLog       As String
Private sCampos         As String
Private sValores        As String
Private rstConsulta     As ADODB.Recordset
Private rstTemp         As ADODB.Recordset
Private bDigitacao      As Boolean
Private bFolhaAberta    As Boolean
Private iPosicao        As Integer
Private cRubrica        As New clsRubrica
Private sEvento         As String

Private lRegistrosAfetados  As Long

Private Sub chkNaoCalcRub_Click()
    Dim iRsp    As Integer

    If chkNaoCalcRub.Value = vbChecked Then
        iRsp = MsgBox("Você está optando por desativar " & vbCrLf & _
        "o cálculo automático da folha." & vbCrLf & _
        "Deseja Continuar?", vbYesNo + vbQuestion + vbDefaultButton2, "Aviso")
        If iRsp = vbNo Then
            chkNaoCalcRub.Value = vbUnchecked
        End If
    End If

End Sub

Private Sub cmbEmpresa_Click()
    If ActiveControl.Name = cmbEmpresa.Name Then
        If cmbEmpresa.Text <> "" And CmbLotacao.Text <> "" Then
            fraFuncionario.Enabled = False
            'Carrega o Grid Para Digitação
            Call PL_Carrega_mshDigitacao
        End If
    End If

End Sub

Private Sub cmbEvento_Click()
    If cmbEvento.Text <> "" Then
        txtValor.Text = FL_Retorna_ValorRubrica(FG_Codigo_Combo(cmbEvento.Text))
    End If

End Sub

Private Sub cmbEvento_GotFocus()
    sEvento = ""

End Sub

Private Sub cmbEvento_KeyPress(KeyAscii As Integer)
    'Guarda o Código do Evento
    If KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Then sEvento = sEvento & Chr(KeyAscii)

    If KeyAscii = vbKeyReturn And sEvento <> "" Then
        cmbEvento.ListIndex = FG_Retorna_Indice_Combo(cmbEvento, Format(sEvento, "0000"), 4, True)
        sEvento = ""
    End If

End Sub

Private Sub cmbEvento_Validate(Cancel As Boolean)
    If cmbEvento.Text <> "" Then
        If FG_Evento_Gerado(FG_Codigo_Combo(cmbEvento.Text)) Then
            cmbEvento.ListIndex = -1
            Cancel = True
        End If
    End If

End Sub

Private Sub cmbFolha_Validate(Cancel As Boolean)
    cmbFolha.Enabled = False

End Sub

Private Sub CmbLotacao_Click()
    If ActiveControl.Name = CmbLotacao.Name Then
        If cmbEmpresa.Text <> "" And CmbLotacao.Text <> "" Then
            fraFuncionario.Enabled = False
            'Carrega o Grid Para Digitação
            Call PL_Carrega_mshDigitacao
        End If
    End If

End Sub

Private Sub Form_Activate()
    'se os valores de autorizacao nao forem referentes a esta transacao
    'todas as linhas abaixo sao de seguranca
    'devera ser mudado o nome do modulo e da transacao
    Dim lModulo As String
    Dim lTran As String
    'verifica seguranca
    lModulo = "RH02M"   'seguranca
    lTran = "RH01A"     'seguranca
    If FG_Seguranca(G_User, G_Sis, lModulo, lTran, "", "") = False Then
       Call FG_MsgBoxPadrao(eMsgUsuarioNaoAutorizado)
       Unload Me
       Exit Sub
    End If

    Call PG_Monta_Combo(CmbOcupacao, eOcupacao)
    Call PG_Monta_Combo(cmbEmpresa, eEmpresa)
    Call PG_Monta_Combo(CmbLotacao, eLotacao)
    Call PG_Monta_Combo(cmbSituacao, eSituacaoFuncionario)
    Call PG_Monta_Combo(cmbEvento, eEvento)
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'Apaga o Nome da tela ativa para comecar a cronometar a saida da tela
    'quando entra no sistema
    If KeyAscii = vbKeyReturn Then
        SendKeys "{Tab}"
    Else
        Call PG_Verifica_Tecla(KeyAscii, eMascMaiusculas)
    End If

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    'Testa a Tecla Alfa Numérica
    Select Case KeyCode
        Case vbKeyEscape
            Call PL_Button_Novo
        Case vbKeyF2
            Call PL_Button_Abrir
        Case vbKeyF4
            'Inicia Digitação ou Vai para o Próximo Registro
            Call PL_Controla_Digitacao

    End Select

End Sub

Private Sub Form_Load()
    Call PG_Centraliza_form(Me)
    Call PG_Exibe_FiguraRH(Me)

    'Configura Imagens da Barra de buttons
    Screen.MousePointer = vbHourglass
    
    Call PG_Configura_Icone_Barra(tlbBotoes)
    tlbBotoes.Buttons(G_Excluir).Enabled = False
    tlbBotoes.Buttons(G_Salvar).Enabled = False
    tlbBotoes.Buttons(6).Enabled = False

    Call PG_Monta_Combo(cmbFolha, eFolha)
    Call PL_Configura_GridEvento
    
    mskReferencia.Text = Format(FG_Data_Sistema, "mm/yyyy")
'    mskReferencia.SetFocus

    bDigitacao = False
    Screen.MousePointer = vbDefault

End Sub

Private Sub mshEvento_Click()
    Dim iLinha      As Integer
    Dim iLinhaAtual As Integer
    Dim iColuna     As Integer

    With mshEvento
        iLinhaAtual = .Row
        'Restaura a Cor de Todas as Linhas com a Cor Padrao
        For iLinha = 1 To .Rows - 1
            .Row = iLinha
            For iLaco = 0 To .Cols - 1
                .Col = iLaco
                .CellBackColor = G_BRANCO
                .CellForeColor = G_PRETO
            Next
        Next

        'Muda a Cor da Linha Selecionada
        .Row = iLinhaAtual
        For iLaco = 0 To .Cols - 1
            .Col = iLaco
            .CellBackColor = &H808000
            .CellForeColor = G_BRANCO
        Next
    
        cmbEvento.ListIndex = FG_Retorna_Indice_Combo(cmbEvento, .TextMatrix(.Row, 0))
    End With

End Sub

Private Sub mshEvento_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo Trata_Erro
    Dim iRsp    As Integer
        
    'A Folha Não Está com Status 1 ou 6 e/ou A Situação do Funcionário Não Está Ativa
    If cmbEvento.Locked Then
        Exit Sub
    End If

    If KeyCode = vbKeyDelete And txtValor.Enabled Then
        If FG_MsgBoxPadrao(eMsgExclusaoRegistro) = vbNo Then Exit Sub
        If Not G_EXC Then
           Call FG_MsgBoxPadrao(eMsgUsuarioNaoAutorizado)
           Exit Sub
        End If

        Screen.MousePointer = vbHourglass
        With mshDigitacao
            sQuery = "Select  FK0504CODEVE, FK0510CODFUN, FK0510CODEMP, "
            sQuery = sQuery & "FK0510CODLOT, RH0510CODCBO, RH05ANORUB, RH05VALRUB "
            sQuery = sQuery & "From    RH05RUBRIT "
            sQuery = sQuery & "Where   "
            sQuery = sQuery & "FK0504CODEVE = '" & Right(mshEvento.TextMatrix(mshEvento.Row, 0), 4) & "' And "
            sQuery = sQuery & "FK0510CODFUN = " & txtMatricula.Text & " And "
            sQuery = sQuery & "FK0510CODEMP = '" & .TextMatrix(iPosicao, 0) & "' And "
            sQuery = sQuery & "FK0510CODLOT = '" & .TextMatrix(iPosicao, 1) & "' And "
            sQuery = sQuery & "RH0510CODCBO = '" & .TextMatrix(iPosicao, 2) & "' And "
            sQuery = sQuery & "RH05INDFER   = '" & Left(cmbFolha, 1) & "' And "
            sQuery = sQuery & "RH05ANORUB   = " & Right(mskReferencia.ClipText, 4) & Left(mskReferencia.ClipText, 2)
            Set rstTemp = G_Conexao_Folha.Execute(sQuery)

            If Not rstTemp.EOF Then
                G_Conexao_Folha.BeginTrans
                sQuery = "Delete RH05RUBRIT "
                sQuery = sQuery & "Where   "
                sQuery = sQuery & "FK0504CODEVE = '" & Right(mshEvento.TextMatrix(mshEvento.Row, 0), 4) & "' And "
                sQuery = sQuery & "FK0510CODFUN = " & txtMatricula.Text & " And "
                sQuery = sQuery & "FK0510CODEMP = '" & .TextMatrix(iPosicao, 0) & "' And "
                sQuery = sQuery & "FK0510CODLOT = '" & .TextMatrix(iPosicao, 1) & "' And "
                sQuery = sQuery & "RH0510CODCBO = '" & .TextMatrix(iPosicao, 2) & "' And "
                sQuery = sQuery & "RH05INDFER   = '" & Left(cmbFolha, 1) & "' And "
                sQuery = sQuery & "RH05ANORUB   = " & Right(mskReferencia.ClipText, 4) & Left(mskReferencia.ClipText, 2)
                G_Conexao_Folha.Execute sQuery, lRegistrosAfetados
                Call PG_Montar_Query_Log("E", rstTemp, sQuery, sQueryLog, "RH05RUBRIT")
                If lRegistrosAfetados > 0 Then
                   Call PG_Salva_Log("RH05RUBRIT", "E", sQueryLog, , txtMatricula.Text)
                End If
                G_Conexao_Folha.CommitTrans

                If chkNaoCalcRub.Value = vbUnchecked Then
                    '>>> Recalcula Rubrica <<<
                    With mshDigitacao
                        Set cRubrica = Nothing
                        If Val(Left(mskReferencia.Text, 2)) < 13 Then
                            cRubrica.Inicializar txtMatricula.Text, .TextMatrix(iPosicao, 0), _
                            .TextMatrix(iPosicao, 1), .TextMatrix(iPosicao, 2), .TextMatrix(iPosicao, 9), _
                            Right(mskReferencia.ClipText, 4) & Left(mskReferencia.ClipText, 2), _
                            "01/" & mskReferencia.Text, , FG_Codigo_Combo(cmbSituacao), Left(cmbFolha, 1)
                        Else
                            cRubrica.Inicializar txtMatricula.Text, .TextMatrix(iPosicao, 0), _
                            .TextMatrix(iPosicao, 1), .TextMatrix(iPosicao, 2), .TextMatrix(iPosicao, 9), _
                            Right(mskReferencia.ClipText, 4) & Left(mskReferencia.ClipText, 2), _
                            "01/" & mskBaseCalculo.Text, , FG_Codigo_Combo(cmbSituacao)
                        End If
                    End With
                    If cRubrica.Erro = False Then
                        cRubrica.Salva
                        Call FG_MsgBoxPadrao(eMsgDadosExcluidos)
                    Else
                        If cRubrica.ErroMsg = "" Then
                            MsgBox "Ocorreu um Erro não catalogado", vbInformation, "Erro"
                        Else
                            MsgBox cRubrica.ErroMsg, vbInformation, "Erro"
                        End If
                    End If
                Else
                    Call FG_MsgBoxPadrao(eMsgDadosExcluidos)
                End If
            Else
                MsgBox "Este Evento Já foi excluído por outro usuário.", vbExclamation, "Mensagem"
            End If
            cmbEvento.ListIndex = -1
            txtValor.Text = ""
            Call PL_Carrega_GridRubrica
        End With

        Screen.MousePointer = vbNormal
    End If

Exit Sub
Trata_Erro:
    Screen.MousePointer = vbNormal
    G_Conexao_Folha.RollbackTrans
End Sub

Private Sub mshEvento_SelChange()
    Call mshEvento_Click

End Sub

Private Sub mskBaseCalculo_GotFocus()
    mskBaseCalculo.SelStart = 0
    mskBaseCalculo.SelLength = Len(mskBaseCalculo.Text)

End Sub

Private Sub mskBaseCalculo_Validate(Cancel As Boolean)
    If Not FG_Retorna_MesAno(mskBaseCalculo) Then
        MsgBox "Referência inválida. Digite mm/aaaa", vbInformation, "Mensagem"
        mskBaseCalculo.SetFocus
        mskBaseCalculo.SelStart = 0
        mskBaseCalculo.SelLength = Len(mskBaseCalculo.Text)
        Cancel = True
    Else
        mskBaseCalculo.Enabled = False
    End If

End Sub

Private Sub mskReferencia_GotFocus()
    mskReferencia.SelStart = 0
    mskReferencia.SelLength = Len(mskReferencia.Text)

End Sub

Private Sub mskReferencia_Validate(Cancel As Boolean)
    If Not FG_Retorna_MesAno(mskReferencia, 15) Then
        MsgBox "Referência inválida. Digite mm/aaaa", vbInformation, "Mensagem"
        mskReferencia.SetFocus
        mskReferencia.SelStart = 0
        mskReferencia.SelLength = Len(mskReferencia.Text)
        Cancel = True
    Else
        mskReferencia.Enabled = False
        lbl13Salario.Visible = False
        cmbFolha.Enabled = True
        If Left(mskReferencia.Text, 2) = "13" Then
            lbl13Salario.Visible = True
            lbl13Salario.Caption = "13º Salário Integral"
        ElseIf Left(mskReferencia.Text, 2) = "14" Then
            lbl13Salario.Visible = True
            lbl13Salario.Caption = "13º Salário (1ª Parcela)"
        ElseIf Left(mskReferencia.Text, 2) = "15" Then
            lbl13Salario.Visible = True
            lbl13Salario.Caption = "13º Salário (2ª Parcela)"
        End If
        If lbl13Salario.Visible Then
            cmbFolha.Enabled = False
            mskBaseCalculo.Enabled = True
            mskBaseCalculo.SetFocus
        Else
            mskBaseCalculo.Text = "__/____"
        End If
    End If

End Sub

Private Sub tlbBotoes_ButtonClick(ByVal Button As MSComctlLib.Button)
  'modifica o ponteiro do mouse
  Screen.MousePointer = vbHourglass

  'inicio do select que acessa as procedures salvar, abrir, etc.
  Select Case Button.Key
    Case "NOVO"
        Call PL_Button_Novo

    Case "ABRIR"
        If bCarregaEmpresa Then
          Call PL_Button_Abrir
        Else
            FG_MsgBoxPadrao (eMsgUsuarioNaoAutorizado)
        End If

    Case "SALVAR"
'         Call PL_Button_Salvar

    Case "DIGITAR"
          Call PL_Controla_Digitacao

    Case "AJUDA"
    Case "SAIR"
         Unload Me
   End Select

   Screen.MousePointer = vbNormal

End Sub

Private Sub TxtMatricula_Change()
    'caso é o controle ativo executa a rotina e não estiver locado
    If txtMatricula.Locked Then Exit Sub

    If ActiveControl.Name = txtMatricula.Name Then
        'caso o campo contenha 4 caracteres executa a consulta
        'senão limpar os campos
         If Len(txtMatricula.Text) = txtMatricula.MaxLength Then
            'monta a sQuery para consultar os registro da tabela RH01FUNCIT
            sQuery = "SELECT * FROM RH01FUNCIT"
            sQuery = sQuery & " WHERE RH01CODFUN = " & txtMatricula
            Set rstConsulta = G_Conexao_Folha.Execute(sQuery)
            'caso existir o registro preencher o formulario
            'senão exibir mensagem e limpar os campos
            If Not rstConsulta.EOF Then
'                If FG_Checa_Funcionario(txtMatricula.Text, True) Then
                    txtNome.Text = Trim("" & rstConsulta!RH01NOMEMP)
                    fraLotacao.Enabled = False

                    'Carrega o Grid Para Digitação
                    Call PL_Carrega_mshDigitacao
'                Else
'                    MsgBox "Funcionário não cadastrado.", vbInformation, "Mensagem"
'                End If
            Else
                MsgBox "Funcionário não cadastrado.", vbInformation, "Mensagem"
            End If
         End If
    End If

End Sub

Private Sub PL_Configura_GridEvento()
    On Error Resume Next
    'PARA SABER O TAMANHO DA COLUNA BASTA MULTIPLICAR
    ' 110 * A QUANTIDADE DE CARACTERES DESEJADOS

    With mshEvento
        .Rows = 1
        .Rows = 2
        .FixedRows = 1
        .FormatString = "<Evento " & Space(150) & "|>Valor                 "

    End With
End Sub

Private Sub PL_Carrega_mshDigitacao()
    tlbBotoes.Buttons(6).Enabled = True
    tlbBotoes.Buttons("DIGITAR").Image = "DIGITAR"
    tlbBotoes.Buttons("DIGITAR").ToolTipText = "Iniciar Digitação (Pressione F4)"

    sQuery = "Select   FK1008CODEMP as Empresa,  FK1002CODLOT As Lotacao, FK1009CODCBO As Ocupacao, "
    sQuery = sQuery & "RH01CODFUN as CodFuncionario, RH01NOMEMP As NomFuncionario, RH10DATSIT as DtSituacao, "
    sQuery = sQuery & "RH10DATADM as DtAdmissao, RH10DATSIT as DtSituacao, FK1003CODSAL As CodSalario, "
'    sQuery = sQuery & "Rtrim(RH13ITEDES)+' - '+CONVERT(VARCHAR,FK1013SITFUN) as Situacao, RH10TIPFOL As TpFolha "
    sQuery = sQuery & "CONVERT(VARCHAR,FK1013SITFUN) as Situacao, RH10TIPFOL As TpFolha "
'    sQuery = sQuery & "From     RH10OCUFUT, RH13CONSTT, RH01FUNCIT "
    sQuery = sQuery & "From     RH10OCUFUT, RH01FUNCIT "
    sQuery = sQuery & "Where    "
'    sQuery = sQuery & "FK1013SITFUN *= RH13ITETAB     And "
    sQuery = sQuery & "FK1001CODFUN =  RH01CODFUN     And "
'    sQuery = sQuery & "RH13TABELA   =  'SITFUNCION'   And "
    sQuery = sQuery & "RH10TIPFOL Is Not Null And "

    If Not bCarregaEmpresa Then
        sQuery = sQuery & "FK1008CODEMP Not In(" & GC_CodEmpresas & ") And "
    End If

    If txtMatricula.Text <> "" Then
        sQuery = sQuery & "FK1001CODFUN  = " & txtMatricula.Text
        sQuery = sQuery & " Order By RH10DATSIT desc "
    ElseIf cmbEmpresa.Text <> "" Then
        sQuery = sQuery & "FK1008CODEMP  = '" & FG_Codigo_Combo(cmbEmpresa.Text) & "' AND "
        sQuery = sQuery & "FK1002CODLOT  = '" & FG_Codigo_Combo(CmbLotacao.Text) & "' "
        sQuery = sQuery & "Order by FK1001CODFUN, RH10DATSIT desc "
    End If
    Set rstTemp = G_Conexao_Folha.Execute(sQuery)
    mshDigitacao.Rows = 0
    mshDigitacao.Cols = 10

    Do While Not rstTemp.EOF
        With mshDigitacao
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = rstTemp!Empresa
            .TextMatrix(.Rows - 1, 1) = Trim(rstTemp!Lotacao)
            .TextMatrix(.Rows - 1, 2) = Trim(rstTemp!Ocupacao)
            .TextMatrix(.Rows - 1, 3) = Format(rstTemp!CodFuncionario, "00000")
            .TextMatrix(.Rows - 1, 4) = Trim(rstTemp!NomFuncionario)
            .TextMatrix(.Rows - 1, 5) = Format(rstTemp!DtAdmissao, "dd/mm/yyyy")
            .TextMatrix(.Rows - 1, 6) = Format(rstTemp!DtSituacao, "dd/mm/yyyy")
            .TextMatrix(.Rows - 1, 7) = rstTemp!Situacao
            .TextMatrix(.Rows - 1, 8) = "" & rstTemp!TpFolha
            .TextMatrix(.Rows - 1, 9) = "" & Trim(rstTemp!CodSalario)
        End With
        rstTemp.MoveNext
    Loop

End Sub

Private Sub PL_Controla_Digitacao()
    If mshDigitacao.Rows = 0 Then
        MsgBox "Não retornou dados.", vbInformation, "Mensagem"
        Screen.MousePointer = vbNormal
        Exit Sub
    End If
    If bDigitacao Then
        iPosicao = iPosicao + 1
    Else
        bDigitacao = True
        iPosicao = 0
    End If
    Call PL_Proximo_Registro

End Sub

Private Sub PL_Proximo_Registro()
    cmbEvento.ListIndex = -1

    If iPosicao >= mshDigitacao.Rows Then
        MsgBox "Não há mais registros.", vbInformation, "Mensagem"
        Call PL_Button_Novo
        Exit Sub
    End If

    tlbBotoes.Buttons(G_Abrir).Enabled = False
    tlbBotoes.Buttons("DIGITAR").Image = "PROXIMO"
    tlbBotoes.Buttons("DIGITAR").ToolTipText = "Próximo Registro (Pressione F4)"

    fraPosicao.Visible = True
    lblPosicao.Caption = "Posição " & Format(iPosicao + 1, "00") & " de " & Format(mshDigitacao.Rows, "00")

    fraRubrica.Enabled = True

    txtMatricula.Locked = True
    txtNome.Locked = True
    cmbEmpresa.Locked = True
    CmbLotacao.Locked = True

    With mshDigitacao
        'Carrega os Campos da Tela
        txtMatricula.Text = .TextMatrix(iPosicao, 3)
        txtNome.Text = .TextMatrix(iPosicao, 4)

        Call PG_Retorna_Campos(Me, "E", .TextMatrix(iPosicao, 0))
        Call PG_Retorna_Campos(Me, "L", .TextMatrix(iPosicao, 1))
        Call PG_Retorna_Campos(Me, "O", .TextMatrix(iPosicao, 2))
        Call PG_Retorna_Campos(Me, "S", .TextMatrix(iPosicao, 7))
        Call PG_Retorna_Campos(Me, "A", .TextMatrix(iPosicao, 5))
        Call PG_Retorna_Campos(Me, "D", .TextMatrix(iPosicao, 6))

        'Verifica o Status da Folha
        sQuery = "Select  Top 1 FK0713STAFOL CodStatus, Rtrim(RH13ITEDES) DesStatus, "
        sQuery = sQuery & "Convert(varchar,RH07DATFOL,120) DtStatus "
        sQuery = sQuery & "From RH07SITFOT, RH13CONSTT "
        sQuery = sQuery & "Where   "
        sQuery = sQuery & "FK0713STAFOL *= RH13ITETAB     And "
        sQuery = sQuery & "RH13TABELA   =  'SITFOLHA'     And "
        sQuery = sQuery & "FK0708CODEMP = '" & .TextMatrix(iPosicao, 0) & "' And "
        sQuery = sQuery & "RH07TIPFOL   =  '" & .TextMatrix(iPosicao, 8) & "'  And "
        sQuery = sQuery & "RH07INDFER   =  '" & Left(cmbFolha.Text, 1) & "'    And "
        sQuery = sQuery & "RH07MESANO   =  " & Right(mskReferencia.ClipText, 4) & Left(mskReferencia.ClipText, 2) & " "
        sQuery = sQuery & "Order By RH07DATFOL Desc"

        Set rstTemp = G_Conexao_Folha.Execute(sQuery)

        bFolhaAberta = False
        If Not rstTemp.EOF Then
            txtStatusFolha.Text = rstTemp!DesStatus & " (" & rstTemp!CodStatus & ") - Data: " & _
            Format(rstTemp!DtStatus, "dd/mm/yyyy") & Right(rstTemp!DtStatus, 9)
            If rstTemp!CodStatus = 1 Or rstTemp!CodStatus = 6 Then 'A Folha Está Aberta
                bFolhaAberta = True
            End If
        Else
            txtStatusFolha.Text = "REFERÊNCIA SEM STATUS"
        End If
        Set rstTemp = Nothing

        'Habilita ou Desabilita a digitação
        If InStr(1, GC_Situacoes_Ativas, FG_Codigo_Combo(cmbSituacao.Text)) = 0 Or Not bFolhaAberta Then
            'Habilita a digitação mesmo que as condições acima sejam verdadeiras
            'pois o usuário selecionou a digitação sem o uso da classe
            If chkNaoCalcRub = vbChecked Then
                txtValor.Enabled = True
            Else
                txtValor.Enabled = False
            End If
        Else
            txtValor.Enabled = True
        End If
    End With

    'Pega os Registros da Rubrica
    Call PL_Carrega_GridRubrica

End Sub

Private Sub PL_Carrega_GridRubrica()
    With mshDigitacao
        sQuery = "Select  Convert(varchar,RH04DESCRI)+'('+ "
        sQuery = sQuery & "Case RH04TIPEVE "
        sQuery = sQuery & "When  'P' Then 'PROVENTO' "
        sQuery = sQuery & "When  'D' Then 'DESCONTO' "
        sQuery = sQuery & "Else  ''  End"
        sQuery = sQuery & " +')'  As Evento, RH04CODEVE Codigo, RH05VALRUB As Valor "
        sQuery = sQuery & "From    RH05RUBRIT, RH04EVENTT "
        sQuery = sQuery & "Where   "
        sQuery = sQuery & "FK0504CODEVE = RH04CODEVE And "
        sQuery = sQuery & "FK0510CODFUN = " & txtMatricula.Text & " And "
        sQuery = sQuery & "FK0510CODEMP = '" & .TextMatrix(iPosicao, 0) & "' And "
        sQuery = sQuery & "FK0510CODLOT = '" & .TextMatrix(iPosicao, 1) & "' And "
        sQuery = sQuery & "RH0510CODCBO = '" & .TextMatrix(iPosicao, 2) & "' And "
        sQuery = sQuery & "RH05INDFER   = '" & Left(cmbFolha, 1) & "' And "
        sQuery = sQuery & "RH05ANORUB   = " & Right(mskReferencia.ClipText, 4) & Left(mskReferencia.ClipText, 2)
        sQuery = sQuery & " Order by RH04TIPEVE Desc, Codigo"
    End With
    Set rstTemp = G_Conexao_Folha.Execute(sQuery)

    'Carrega a Grid dos Eventos
    With mshEvento
        .Rows = 1
        If rstTemp.EOF Then
            .Rows = 2
        End If

        Do While Not rstTemp.EOF
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = rstTemp!Evento & " «» " & Format(rstTemp!Codigo, "0000")
            .TextMatrix(.Rows - 1, 1) = Format(rstTemp!Valor, "#,###,##0.00")
            rstTemp.MoveNext
        Loop
        .FixedRows = 1
        .FormatString = "<Evento " & Space(150) & "|>Valor                 "

    End With

End Sub

Private Sub PL_Button_Novo()
    Call PG_Limpa_Campos(Me, mshEvento.Name, mskReferencia.Name, cmbFolha.Name, mskBaseCalculo.Name)
    Call PL_Configura_GridEvento

    mskReferencia.Enabled = True
    mskBaseCalculo.Enabled = False
    cmbFolha.Enabled = True
'    mskReferencia.Text = Format(FG_Data_Sistema, "mm/yyyy")

    mshDigitacao.Rows = 0
    
    fraPosicao.Visible = False

    fraFuncionario.Enabled = True
    txtMatricula.Locked = False
    txtNome.Locked = False

    fraLotacao.Enabled = True
    cmbEmpresa.Locked = False
    CmbLotacao.Locked = False

    tlbBotoes.Buttons(G_Abrir).Enabled = True
    tlbBotoes.Buttons(6).Enabled = False
    tlbBotoes.Buttons("DIGITAR").Image = "DIGITAR"
    tlbBotoes.Buttons("DIGITAR").ToolTipText = "Iniciar Digitação  (Pressione F4)"
    bDigitacao = False
    mskReferencia.SetFocus

End Sub

Private Sub PL_Salvar()
    On Error GoTo Tratar_Erro

    Dim iRsp    As Integer
    
    iRsp = MsgBox("Salvar Rubrica?", vbQuestion + vbYesNo + vbDefaultButton1, "Mensagem")
    If iRsp = vbNo Then
        Exit Sub
    End If

    Screen.MousePointer = vbHourglass
    If cmbEvento.Text <> "" And txtValor <> "" Then
        G_Conexao_Folha.BeginTrans
        'Verifica se já Existe o Evento na Rubrica
        With mshDigitacao
            sQuery = "Select  FK0504CODEVE, FK0510CODFUN, FK0510CODEMP, "
            sQuery = sQuery & "FK0510CODLOT, RH0510CODCBO, RH05ANORUB, RH05VALRUB, RH05INDFER "
            sQuery = sQuery & "From    RH05RUBRIT "
            sQuery = sQuery & "Where   "
            sQuery = sQuery & "FK0504CODEVE = '" & FG_Codigo_Combo(cmbEvento.Text) & "' And "
            sQuery = sQuery & "FK0510CODFUN = " & txtMatricula.Text & " And "
            sQuery = sQuery & "FK0510CODEMP = '" & .TextMatrix(iPosicao, 0) & "' And "
            sQuery = sQuery & "FK0510CODLOT = '" & .TextMatrix(iPosicao, 1) & "' And "
            sQuery = sQuery & "RH0510CODCBO = '" & .TextMatrix(iPosicao, 2) & "' And "
            sQuery = sQuery & "RH05INDFER   = '" & Left(cmbFolha, 1) & "' And "
            sQuery = sQuery & "RH05ANORUB   = " & Right(mskReferencia.ClipText, 4) & Left(mskReferencia.ClipText, 2)
            Set rstTemp = G_Conexao_Folha.Execute(sQuery)

            If rstTemp.EOF Then
                If Not G_INS Then
                   Call FG_MsgBoxPadrao(eMsgUsuarioNaoAutorizado)
                   GoTo Fim
                End If

                'Se NÃO Existir, Cadastra o Evento na Rubrica
                sCampos = "FK0504CODEVE, FK0510CODFUN, FK0510CODEMP, "
                sCampos = sCampos & "FK0510CODLOT, RH0510CODCBO, RH05ANORUB, RH05VALRUB, RH05INDFER"

                sValores = "'" & FG_Codigo_Combo(cmbEvento.Text) & "', " & txtMatricula.Text & ", "
                sValores = sValores & "'" & .TextMatrix(iPosicao, 0) & "', '" & .TextMatrix(iPosicao, 1) & "', "
                sValores = sValores & "'" & .TextMatrix(iPosicao, 2) & "', "
                sValores = sValores & Right(mskReferencia.ClipText, 4) & Left(mskReferencia.ClipText, 2) & ", "
                sValores = sValores & FG_Conv_Ponto(CCur(txtValor.Text)) & ", '" & Left(cmbFolha, 1) & "'"

                sQuery = "Insert Into RH05RUBRIT (" & sCampos & ") Values (" & sValores & ")"

                Call PG_Montar_Query_Log("I", rstTemp, sQuery, sQueryLog, "RH05RUBRIT")
            Else
                If Not G_ALT Then
                   Call FG_MsgBoxPadrao(eMsgUsuarioNaoAutorizado)
                   GoTo Fim
                End If

                'Se Existir Atualiza o Valor do Evento na Rubrica
                sQuery = "Update RH05RUBRIT Set RH05VALRUB = " & FG_Conv_Ponto(CCur(txtValor.Text))
                sQuery = sQuery & " Where   "
                sQuery = sQuery & "FK0504CODEVE = '" & FG_Codigo_Combo(cmbEvento.Text) & "' And "
                sQuery = sQuery & "FK0510CODFUN = " & txtMatricula.Text & " And "
                sQuery = sQuery & "FK0510CODEMP = '" & .TextMatrix(iPosicao, 0) & "' And "
                sQuery = sQuery & "FK0510CODLOT = '" & .TextMatrix(iPosicao, 1) & "' And "
                sQuery = sQuery & "RH0510CODCBO = '" & .TextMatrix(iPosicao, 2) & "' And "
                sQuery = sQuery & "RH05INDFER   = '" & Left(cmbFolha, 1) & "' And "
                sQuery = sQuery & "RH05ANORUB   = " & Right(mskReferencia.ClipText, 4) & Left(mskReferencia.ClipText, 2)

                Call PG_Montar_Query_Log("A", rstTemp, sQuery, sQueryLog, "RH05RUBRIT")
            End If
        End With

        G_Conexao_Folha.Execute sQuery, lRegistrosAfetados
        If lRegistrosAfetados > 0 Then
            If rstTemp.EOF Then
                Call PG_Salva_Log("RH05RUBRIT", "I", sQueryLog, , txtMatricula.Text)
            Else
                Call PG_Salva_Log("RH05RUBRIT", "A", sQueryLog, , txtMatricula.Text)
            End If
        End If
        G_Conexao_Folha.CommitTrans

        If chkNaoCalcRub.Value = vbUnchecked Then
            'Recalcula Rubrica
            With mshDigitacao
                Set cRubrica = Nothing
                If Val(Left(mskReferencia.Text, 2)) < 13 Then
                    cRubrica.Inicializar txtMatricula.Text, .TextMatrix(iPosicao, 0), .TextMatrix(iPosicao, 1), _
                    .TextMatrix(iPosicao, 2), .TextMatrix(iPosicao, 9), FG_AnoMes(mskReferencia.ClipText), _
                    "01/" & mskReferencia.Text, , FG_Codigo_Combo(cmbSituacao), Left(cmbFolha, 1)
                Else
                    cRubrica.Inicializar txtMatricula.Text, .TextMatrix(iPosicao, 0), .TextMatrix(iPosicao, 1), _
                    .TextMatrix(iPosicao, 2), .TextMatrix(iPosicao, 9), FG_AnoMes(mskReferencia.ClipText), _
                    "01/" & mskBaseCalculo.Text, , FG_Codigo_Combo(cmbSituacao)
                End If
            End With
            If cRubrica.Erro = False Then
                cRubrica.Salva
            Else
                If cRubrica.ErroMsg = "" Then
                    MsgBox "Ocorreu um Erro não catalogado", vbInformation, "Erro"
                Else
                    MsgBox cRubrica.ErroMsg, vbInformation, "Erro"
                End If
            End If
        End If

        Call PL_Carrega_GridRubrica

        Call FG_MsgBoxPadrao(eMsgDadosSalvos)
    End If

Fim:
    cmbEvento.ListIndex = -1
    txtValor.Text = ""
    cmbEvento.SetFocus
    Screen.MousePointer = vbNormal

    Exit Sub

Tratar_Erro:
    cmbEvento.SetFocus
    Screen.MousePointer = vbNormal
    G_Conexao_Folha.RollbackTrans
    Call FG_MsgBoxPadrao(eMsgErroGravacao, sQuery)

End Sub

Private Sub PL_Button_Abrir()
    'prompt para montar a lista com código e serviço no formulário frmGlista
    If Trim(txtNome.Text) <> "" Then
        g_sql = "SELECT RH01CODFUN as Codigo,RH01NOMEMP as Nome  " & _
                " FROM RH01FUNCIT " & _
                " where RH01NOMEMP  like '" & Trim(txtNome.Text) & "%'  "
    Else
        Call FG_MsgBoxPadrao(eMsgInformeCampo, "nome com as iniciais do Funcionário")
        txtNome.SetFocus
        Exit Sub
    End If
    g_sql = g_sql & "order by RH01NOMEMP asc"

    'inicializa as variáveis
    G_Colunas_Lista = 2
    G_Lista(0).Mascara = "00000"
    G_Mensagem_Lista = "Não há Funcionários cadastrados"
    G_Titulo_Lista = "Lista de Funcionários"

    'passa o parametro para o sQuery do botão consulta do formulário frm Glista
    G_Tabela = "RH01FUNCIT"
    G_Atributo = "RH01CODFUN as Codigo,RH01NOMEMP as nome "
    G_Condicao = "RH01NOMEMP  like '"
    G_Ordem = "RH01NOMEMP "

    frmGLista.Show vbModal

    If G_Retorno_Lista(0) = Space$(0) Then
       Exit Sub
    End If

    'preenche os campos com o conteúdo do atributos\
    txtMatricula.SetFocus
    txtMatricula.Text = ""
    txtMatricula.Text = G_Retorno_Lista(0)

    G_Retorno_Lista(0) = Space$(0)
    G_Retorno_Lista(1) = Space$(0)

'    Call txtMatricula_LostFocus

End Sub

Private Function FL_Retorna_ValorRubrica(sEvento As String) As String
    With mshDigitacao
        sQuery = "Select  RH05VALRUB As Valor "
        sQuery = sQuery & "From    RH05RUBRIT "
        sQuery = sQuery & "Where   "
        sQuery = sQuery & "FK0504CODEVE = '" & sEvento & "' And "
        sQuery = sQuery & "FK0510CODFUN = " & txtMatricula.Text & " And "
        sQuery = sQuery & "FK0510CODEMP = '" & .TextMatrix(iPosicao, 0) & "' And "
        sQuery = sQuery & "FK0510CODLOT = '" & .TextMatrix(iPosicao, 1) & "' And "
        sQuery = sQuery & "RH0510CODCBO = '" & .TextMatrix(iPosicao, 2) & "' And "
        sQuery = sQuery & "RH05INDFER   = '" & Left(cmbFolha, 1) & "' And "
        sQuery = sQuery & "RH05ANORUB   = " & Right(mskReferencia.ClipText, 4) & Left(mskReferencia.ClipText, 2)
    End With
    Set rstTemp = G_Conexao_Folha.Execute(sQuery)
    If rstTemp.EOF Then
        FL_Retorna_ValorRubrica = ""
    Else
        FL_Retorna_ValorRubrica = Format(rstTemp!Valor, "#,###,##0.00")
    End If

End Function

Private Sub TxtMatricula_GotFocus()
    mskBaseCalculo.Enabled = False

End Sub

Private Sub txtMatricula_LostFocus()
    If Len(Trim(txtMatricula)) < 5 And Trim(txtMatricula) <> "" Then
        txtMatricula.SetFocus
        txtMatricula.Text = Format(txtMatricula.Text, "00000")
    End If

End Sub

Private Sub txtValor_GotFocus()
    txtValor.SelStart = 0
    txtValor.SelLength = Len(txtValor.Text)

End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)
    Call PG_Verifica_Tecla(KeyAscii, eMascMoeda, txtValor)

End Sub

Private Sub txtValor_LostFocus()
    If cmbEvento.Text <> "" Then
        Call PL_Salvar
    End If

End Sub

