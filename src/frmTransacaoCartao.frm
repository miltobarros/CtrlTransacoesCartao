VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form frmTransacaoCartao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Transações de Cartão de Crédito"
   ClientHeight    =   8745
   ClientLeft      =   345
   ClientTop       =   1050
   ClientWidth     =   9300
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTransacaoCartao.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8745
   ScaleWidth      =   9300
   Begin VB.Frame fraTransacoes 
      Caption         =   "Transações de Crédito"
      Height          =   7965
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   9045
      Begin VB.CommandButton cmdFiltro 
         Caption         =   "&Aplicar Filtro"
         Height          =   375
         Left            =   2430
         MaskColor       =   &H00FF0000&
         TabIndex        =   18
         Top             =   3045
         Width           =   1335
      End
      Begin VB.CommandButton cmProximaPagina 
         Caption         =   "Próxima Página"
         Height          =   495
         Left            =   6960
         TabIndex        =   17
         Top             =   7350
         Width           =   1095
      End
      Begin VB.CommandButton cmdPaginaAnterior 
         Caption         =   "Página Anterior"
         Height          =   495
         Left            =   360
         TabIndex        =   16
         Top             =   7350
         Width           =   1095
      End
      Begin MSDataGridLib.DataGrid DataGridTransacoes 
         Height          =   3735
         Left            =   360
         TabIndex        =   15
         Top             =   3510
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   6588
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txtValor 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3360
         MaxLength       =   18
         TabIndex        =   7
         Top             =   960
         Width           =   2505
      End
      Begin VB.ComboBox cmbStatus 
         Height          =   330
         ItemData        =   "frmTransacaoCartao.frx":030A
         Left            =   6240
         List            =   "frmTransacaoCartao.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1635
         Width           =   1845
      End
      Begin VB.TextBox txtDscTransacao 
         Height          =   1320
         Left            =   300
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   1635
         Width           =   5565
      End
      Begin VB.TextBox txtTransacao 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   1200
         MaxLength       =   6
         TabIndex        =   3
         Top             =   270
         Width           =   945
      End
      Begin VB.TextBox txtNoCartao 
         Height          =   315
         Left            =   300
         MaxLength       =   16
         TabIndex        =   5
         Top             =   960
         Width           =   2625
      End
      Begin MSMask.MaskEdBox MskDtTransacao 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
         Height          =   315
         Left            =   6225
         TabIndex        =   9
         Top             =   960
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   16
         Mask            =   "##/##/####\ ##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label lblTransacoes 
         AutoSize        =   -1  'True
         Caption         =   "Transações Cadastradas :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   240
         TabIndex        =   14
         Top             =   3120
         Width           =   2145
      End
      Begin VB.Label lblStatusTransacao 
         AutoSize        =   -1  'True
         Caption         =   "Status da Transação :"
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   6120
         TabIndex        =   12
         Top             =   1410
         Width           =   1605
      End
      Begin VB.Label lblValor 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor :"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   3210
         TabIndex        =   6
         Top             =   750
         Width           =   450
      End
      Begin VB.Label lblDscTransacao 
         AutoSize        =   -1  'True
         Caption         =   "Descrição :"
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   285
         TabIndex        =   10
         Top             =   1410
         Width           =   825
      End
      Begin VB.Label lblDtTransacao 
         AutoSize        =   -1  'True
         Caption         =   "Data da Transação :"
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   6105
         TabIndex        =   8
         Top             =   750
         Width           =   1470
      End
      Begin VB.Label lblTransacao 
         AutoSize        =   -1  'True
         Caption         =   "Transação :"
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   285
         TabIndex        =   2
         Top             =   330
         Width           =   870
      End
      Begin VB.Label lblNoCartao 
         AutoSize        =   -1  'True
         Caption         =   "Número do Cartão :"
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   285
         TabIndex        =   4
         Top             =   750
         Width           =   1395
      End
   End
   Begin MSComctlLib.Toolbar tlbBotoes 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9300
      _ExtentX        =   16404
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Novo"
            Object.ToolTipText     =   "Novo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Salvar"
            Object.ToolTipText     =   "Salvar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Excluir"
            Object.ToolTipText     =   "Excluir"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Procurar"
            Object.ToolTipText     =   "Procurar"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Imprimir"
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Fechar"
            Object.ToolTipText     =   "Fechar"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Ajuda"
            Object.ToolTipText     =   "Ajuda"
            ImageIndex      =   7
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "frmTransacaoCartao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rsCarregaGrid As ADODB.Recordset
Private Const iPageSize As Integer = 20
Private iCurrentPage As Integer
Private ioffsetValue As Integer
Private bFiltro As Boolean
Private sQryFiltro As String


Private Sub PL_LimpaTela()
    Call PG_Limpa_Campos(Me)
    txtTransacao.SetFocus
    txtTransacao.Text = ""
    sQryFiltro = ""

'    With AdodcTransacoes
'            .ConnectionString = dbBancoDados.ConnectionString
'            .MaxRecords = 200
'            .RecordSource = "Select * From dbo.CadastroTransacoes"
'            .Refresh
'    End With
'    DataGridTransacoes.Refresh
    iCurrentPage = 1
    Call PL_Carrega_DataGrid

End Sub

Private Sub PL_Recupera_Dados()
    On Error GoTo TrataErro
    
    G_sQry = "Select * From dbo.CadastroTransacoes Where Id_Transacao = " & txtTransacao.Text
    Set G_rsGlobal = dbBancoDados.Execute(G_sQry)
    With G_rsGlobal
        If Not .EOF Then
            txtNoCartao.Text = "" & !Numero_Cartao
            txtValor.Text = "" & Format(!Valor_Transacao, GC_Formato_Moeda)
            MskDtTransacao.Text = Format(!Data_Transacao, "dd/mm/yyyy hh:mm")
            txtDscTransacao.Text = "" & !Descricao
            cmbStatus.ListIndex = FG_Indice_Combo(cmbStatus, "" & !Status_Transacao)
        End If
        .Close
    End With
    Exit Sub

TrataErro:
    Screen.MousePointer = vbDefault
    'Grava o Erro no arquivo de Log
    Call GravarErroLog("Erro em procedimento do Cadastro de Transações: PL_Recupera_Dados " & Err.Number & " - " & Err.Description)
    MsgBox "Erro na tentativa de carregar os dados. " & vbCr & "Tente novamente mais tarde.", vbExclamation, "Aviso"
    txtTransacao.SetFocus


End Sub

Private Sub PL_Carrega_DataGrid()
    On Error GoTo TrataErro
    
    ioffsetValue = (iCurrentPage - 1) * iPageSize
    
    G_sQry = ""
    G_sQry = G_sQry & "WITH Paginado AS ("
    G_sQry = G_sQry & "SELECT   Id_Transacao, Numero_Cartao, Valor_Transacao, Data_Transacao, Descricao, Status_Transacao, "
    G_sQry = G_sQry & "ROW_NUMBER() OVER (ORDER BY Id_Transacao) AS RowNum "
    G_sQry = G_sQry & "FROM     dbo.CadastroTransacoes"
    
    If sQryFiltro <> "" Then
        G_sQry = G_sQry & " Where " & sQryFiltro
    End If
    
    G_sQry = G_sQry & ") "
    G_sQry = G_sQry & "SELECT * FROM Paginado WHERE RowNum BETWEEN "
    G_sQry = G_sQry & ((iCurrentPage - 1) * iPageSize + 1) & " AND " & (iCurrentPage * iPageSize)
    Set rsCarregaGrid = dbBancoDados.Execute(G_sQry)
    Set DataGridTransacoes.DataSource = rsCarregaGrid
    Exit Sub

TrataErro:
    Screen.MousePointer = vbDefault
    'Grava o Erro no arquivo de Log
    Call GravarErroLog("Erro em procedimento do Cadastro de Transações: PL_Carrega_DataGrid " & Err.Number & " - " & Err.Description)
    MsgBox "Erro na tentativa de carregar o DataGrid. " & vbCr & "Tente novamente mais tarde.", vbExclamation, "Aviso"

End Sub

Private Sub PL_Salvar_Registro()
    On Error GoTo TrataErro

    'Checa se os Campos Obrigatórios Estão Preenchidos
    If FG_Campo_Nao_Informado(txtNoCartao, txtValor, MskDtTransacao, txtDscTransacao, cmbStatus) Then Exit Sub

    Screen.MousePointer = vbHourglass
    G_sQry = "Select Id_Transacao, Numero_Cartao, Status_Transacao From dbo.CadastroTransacoes Where Id_Transacao = 0" & txtTransacao.Text
    Set G_rsGlobal = dbBancoDados.Execute(G_sQry)

    If G_rsGlobal.EOF Then
        'Insere Novo Registro
        G_sCampos = ""
        G_sCampos = G_sCampos & "Numero_Cartao, Valor_Transacao, Data_Transacao, Descricao, Status_Transacao"

        G_sValores = ""
        G_sValores = G_sValores & "'" & txtNoCartao.Text & "', " & FG_Conv_Ponto(CCur(txtValor.Text)) & ", "
        G_sValores = G_sValores & "'" & MskDtTransacao.Text & "', '" & txtDscTransacao.Text & "', "
        G_sValores = G_sValores & "'" & cmbStatus.Text & "'"

        'Query de Inserção de Registro
        G_sQry = " Insert Into dbo.CadastroTransacoes  (" & G_sCampos & ") Values (" & G_sValores & ")"
    Else
        If UCase$(G_rsGlobal!Status_Transacao) = "APROVADA" Then
            MsgBox "Transação com status de APROVADA." & vbCrLf & "Alteração não permitida.", vbInformation, "Aviso"
            G_rsGlobal.Close
            Exit Sub
        End If
        
        'Atualiza Registro
        G_sQry = ""
        G_sQry = G_sQry & "Update dbo.CadastroTransacoes  "
        G_sQry = G_sQry & "Set  "
        G_sQry = G_sQry & "Numero_Cartao = '" & txtNoCartao.Text & "', Valor_Transacao = " & FG_Conv_Ponto(CCur(txtValor.Text)) & ", "
        G_sQry = G_sQry & "Data_Transacao = '" & MskDtTransacao.Text & "', Descricao = '" & txtDscTransacao.Text & "', "
        G_sQry = G_sQry & "Status_Transacao = '" & cmbStatus.Text & "'"
    End If

    'Executa a Operação
    dbBancoDados.BeginTrans
    dbBancoDados.Execute G_sQry
    dbBancoDados.CommitTrans

    Screen.MousePointer = vbDefault
    Call FG_MsgBoxPadrao(eMsgDadosSalvos)
    Call PL_LimpaTela
    Exit Sub

TrataErro:
    dbBancoDados.RollbackTrans
    Screen.MousePointer = vbDefault
    'Grava o Erro no arquivo de Log
    Call GravarErroLog("Erro em procedimento do Cadastro de Transações: PL_Salvar_Registro " & Err.Number & " - " & Err.Description)
    Call FG_MsgBoxPadrao(eMsgErroGravacao)
    txtTransacao.SetFocus

End Sub

Private Sub PL_Excluir_Registro()
    On Error GoTo TrataErro
    
    Dim iRsp    As Integer

    'Checa se os Campos Obrigatórios Estão Preenchidos
    If FG_Campo_Nao_Informado(txtTransacao, txtNoCartao, txtValor, MskDtTransacao, txtDscTransacao, cmbStatus) Then Exit Sub
    
    G_sQry = "Select Id_Transacao, Numero_Cartao, Status_Transacao From dbo.CadastroTransacoes Where Id_Transacao = 0" & txtTransacao.Text
    Set G_rsGlobal = dbBancoDados.Execute(G_sQry)

    If G_rsGlobal.EOF Then
        MsgBox "Registro já excluído.", vbInformation, "Aviso"
        G_rsGlobal.Close
        Exit Sub
    Else
        If UCase$(G_rsGlobal!Status_Transacao) = "APROVADA" Then
            MsgBox "Transação com status de APROVADA." & vbCrLf & "Exclusão não permitida.", vbInformation, "Aviso"
            G_rsGlobal.Close
            Exit Sub
        End If
    End If
    G_rsGlobal.Close


    iRsp = MsgBox("Excluir Registro?", vbQuestion + vbYesNo + vbDefaultButton2, "Mensagem")
    If iRsp = vbNo Then Exit Sub

    Screen.MousePointer = vbHourglass
    'Apaga o Proprietário e seus Veículos
    G_sQry = "Delete From dbo.CadastroTransacoes Where Id_Transacao = 0" & txtTransacao.Text

    dbBancoDados.BeginTrans
    dbBancoDados.Execute G_sQry
    dbBancoDados.CommitTrans

    Screen.MousePointer = vbDefault
    Call FG_MsgBoxPadrao(eMsgDadosExcluidos)
    Call PL_LimpaTela
    txtTransacao.SetFocus
    Exit Sub

TrataErro:
    dbBancoDados.RollbackTrans
    Screen.MousePointer = vbDefault
    'Grava o Erro no arquivo de Log
    Call GravarErroLog("Erro em procedimento do Cadastro de Transações: PL_Excluir_Registro " & Err.Number & " - " & Err.Description)
    Call FG_MsgBoxPadrao(eMsgErroExclusao)
    txtTransacao.SetFocus

End Sub

Private Sub cmbStatus_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then cmbStatus.ListIndex = -1

End Sub

Private Sub cmbStatus_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyBack Then cmbStatus.ListIndex = -1

End Sub

Private Sub cmdFiltro_Click()
    On Error GoTo TrataErro

    bFiltro = False
    sQryFiltro = ""
    G_sQryBusca = " "
    If txtNoCartao.Text <> "" Then
        sQryFiltro = sQryFiltro & " Numero_Cartao like '" & txtNoCartao.Text & "%' "
        G_sQryBusca = " AND "
        bFiltro = True
    End If

    If txtValor.Text <> "" Then
        sQryFiltro = sQryFiltro & G_sQryBusca & " Valor_Transacao = " & FG_Conv_Ponto(CCur(txtValor.Text))
        G_sQryBusca = " AND "
        bFiltro = True
    End If

    If MskDtTransacao.ClipText <> "" Then
        'FILTRA APENAS A DATA IGNORANDO A HORA
        sQryFiltro = sQryFiltro & G_sQryBusca & " CONVERT(VARCHAR(10), Data_Transacao, 103)  = '" & Left(MskDtTransacao.Text, 10) & "'"
        G_sQryBusca = " AND "
        bFiltro = True
    End If

    If cmbStatus.Text <> "" Then
        sQryFiltro = sQryFiltro & G_sQryBusca & " Status_Transacao = '" & cmbStatus.Text & "'"
        bFiltro = True
    End If

    If Not bFiltro Then
        MsgBox "Preencha pelo menos um dos campos de filtro:" & vbCrLf & _
        "Número do Cartão, Valor, Data da Transação e/ou Status da Transação", vbInformation, "Aviso"
    Else
        iCurrentPage = 1
        Call PL_Carrega_DataGrid
    End If
    Exit Sub

TrataErro:
    Screen.MousePointer = vbDefault
    'Grava o Erro no arquivo de Log
    Call GravarErroLog("Erro em procedimento do Cadastro de Transações: cmdFiltro_Click " & Err.Number & " - " & Err.Description)
    MsgBox "Erro na tentativa de filtrar registros. " & vbCr & "Tente novamente mais tarde.", vbExclamation, "Aviso"
    txtTransacao.SetFocus

End Sub

Private Sub cmdPaginaAnterior_Click()
    If iCurrentPage > 1 Then
        iCurrentPage = iCurrentPage - 1
        Call PL_Carrega_DataGrid
    End If

End Sub

Private Sub cmProximaPagina_Click()
    iCurrentPage = iCurrentPage + 1
    Call PL_Carrega_DataGrid

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Call PG_Muda_Teclas(KeyAscii)
    Call PG_Verifica_Tecla(KeyAscii, eMascMaiusculas)

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case FG_Tecla_Atalho(KeyCode, Shift)
        Case "CTRL+N" 'Novo
            Call PL_LimpaTela

        Case "CTRL+S" 'Salvar
            Call PL_Salvar_Registro

        Case "CTRL+P" 'Procurar
'            Call PL_Procura_Registro

        Case "CTRL+E" 'Excluir
            Call PL_Excluir_Registro

        Case "CTRL+I" 'Imprimir
            Call PL_Exportar_Transacoes

        Case "ESC" ' Fechar
            Unload Me

        Case "F1" 'Ajuda

    End Select

End Sub

Private Sub Form_Load()
    'Centraliza Formulário
    Call PG_Centraliza_form(Me)

    'Configura Barra de Opções
    Call PG_Configura_Icone_Barra(tlbBotoes)
    
    'Adiciona os item da Combo do Status da Transação
    cmbStatus.AddItem "APROVADA"
    cmbStatus.AddItem "PENDENTE"
    cmbStatus.AddItem "CANCELADA"
    
    iCurrentPage = 1
    Call PL_Carrega_DataGrid


End Sub

Private Sub mskCep_KeyPress(KeyAscii As Integer)
    Call PG_Verifica_Tecla(KeyAscii, eMascNumero)

End Sub

Private Sub MskDtTransacao_KeyPress(KeyAscii As Integer)
    Call PG_Verifica_Tecla(KeyAscii, eMascNumero)

End Sub

Private Sub MskDtTransacao_Validate(Cancel As Boolean)
    If Trim(MskDtTransacao.ClipText) <> "" Then
        If Not IsDate(Left(MskDtTransacao.Text, 10)) Then
            MsgBox "Data inválida", vbExclamation, "Aviso"
            Cancel = True
        End If

        If Not FG_ValidaHora(Right(MskDtTransacao.Text, 5) & ":00") Then
            MsgBox "Hora inválida", vbExclamation, "Aviso"
            Cancel = True
        End If
    End If

End Sub

Private Sub tlbBotoes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case Is = "Novo"
            Call PL_LimpaTela

        Case Is = "Salvar"
            Call PL_Salvar_Registro

        Case Is = "Excluir"
            Call PL_Excluir_Registro

'        Case Is = "Procurar"
'            Call PL_Procura_Registro

        Case Is = "Imprimir"
            Call PL_Exportar_Transacoes

        Case Is = "Fechar"
            Unload Me

        Case Is = "Ajuda"
    End Select

End Sub

Private Sub txtNoCartao_KeyPress(KeyAscii As Integer)
    Call PG_Verifica_Tecla(KeyAscii, eMascNumero)

End Sub

Private Sub txtTransacao_GotFocus()
    txtTransacao.SelStart = 0
    txtTransacao.SelLength = Len(txtTransacao)

End Sub

Private Sub txtTransacao_KeyPress(KeyAscii As Integer)
    Call PG_Verifica_Tecla(KeyAscii, eMascNumero)

End Sub

Private Sub txtTransacao_LostFocus()
    If Val(txtTransacao.Text) > 0 Then
        Call PL_Recupera_Dados
    End If

End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)
    Call PG_Verifica_Tecla(KeyAscii, eMascMoeda)

End Sub

Public Sub PL_Exportar_Transacoes()
    On Error GoTo TrataErro

    Dim CaminhoArquivo As String
    Dim Arquivo As Integer
    Dim i As Integer
    Dim Linha As String
    Dim iRsp As Integer

    iRsp = MsgBox("Confirma Exportação das Transações do Mês " & Format(Date, "mm/yyyy") & "?", vbQuestion + vbYesNo + vbDefaultButton1, "Mensagem")
    If iRsp = vbNo Then Exit Sub
    
    ' Monta a query
    G_sQry = ""
    G_sQry = G_sQry & "SELECT   Id_Transacao AS 'Id Transação', Numero_Cartao AS 'No. Cartao', Valor_Transacao AS 'Valor da Transação', "
    G_sQry = G_sQry & "         Data_Transacao AS 'Data', Descricao AS 'Descrição', Status_Transacao AS 'Status' "
    G_sQry = G_sQry & "FROM     dbo.CadastroTransacoes WHERE MONTH(Data_Transacao) = MONTH(GETDATE())"
    G_sQry = G_sQry & "Order by Data_Transacao"

    ' Executa a query
    Set G_rsGlobal = dbBancoDados.Execute(G_sQry)
    
    ' Caminho do arquivo CSV
    CaminhoArquivo = App.Path & "\RelatorioMes.csv"
    
    ' Apaga o arquivo se já existir
    If Dir(CaminhoArquivo) <> "" Then
        Kill CaminhoArquivo
    End If

    ' Cria o arquivo CSV
    Arquivo = FreeFile
    Open CaminhoArquivo For Output As #Arquivo
    
    ' Escreve cabeçalhos (nomes das colunas)
    For i = 0 To G_rsGlobal.Fields.Count - 1
        Linha = Linha & G_rsGlobal.Fields(i).Name
        If i < G_rsGlobal.Fields.Count - 1 Then Linha = Linha & ";"
    Next i
    Print #Arquivo, Linha
    
    ' Escreve os dados linha por linha
    Do While Not G_rsGlobal.EOF
        Linha = ""
        For i = 0 To G_rsGlobal.Fields.Count - 1
            Linha = Linha & """" & Replace(G_rsGlobal.Fields(i).Value & "", """", "'") & """"
            If i < G_rsGlobal.Fields.Count - 1 Then Linha = Linha & ";"
        Next i
        Print #Arquivo, Linha
        G_rsGlobal.MoveNext
    Loop
    
    Close #Arquivo
    
    MsgBox "Relatório exportado com sucesso para: " & CaminhoArquivo & "\RelatorioMes.csv", vbInformation, "Exportação CSV"
    
    ' Fecha conexões
    G_rsGlobal.Close
    Set G_rsGlobal = Nothing
    Exit Sub

TrataErro:
    Call GravarErroLog("Erro em procedimento do Cadastro de Transações: PL_Exportar_Transacoes " & Err.Number & " - " & Err.Description)
    MsgBox "Erro ao gerar relatório: " & Err.Description, vbCritical, "Erro"

End Sub

