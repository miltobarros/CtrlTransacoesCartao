Attribute VB_Name = "Modulo_Geral"
Option Explicit

'Manipulação do Arquivo ".INI"
Declare Function GravaEntradaPrivIni Lib "kernel32" _
      Alias "WritePrivateProfileStringA" _
     (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
     ByVal lpString As Any, _
     ByVal lpFileName As String) As Long
Declare Function apiLeEntradaPrivIni Lib "kernel32" _
     Alias "GetPrivateProfileStringA" _
     (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
     ByVal lpDefault As String, _
     ByVal lpReturnedString As String, ByVal nSize As Long, _
     ByVal lpFileName As String) As Long
Declare Function GetVolumeInformation Lib "kernel32" _
     Alias "GetVolumeInformationA" _
     (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer _
     As String, ByVal _
     nVolumeNameSize As Long, lpVolumeSerialNumber As Long, _
     lpMaximumComponentLength As Long, lpFileSystemFlags As Long, _
     ByVal lpFileSystemNameBuffer As String, _
     ByVal nFileSystemNameSize As Long) As Long

'API DO USUÁRIO DO LOGADO NO WINDOWS
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
(ByVal lpBuffer As String, nSize As Long) As Long

'NOME DOS MESES
Global Const GC_Meses = "janeiro,fevereiro,março,abril,maio,junho,julho,agosto,setembro,outubro,novembro,dezembro"

'Constantes das Cores
Global Const GC_Vermelho = &HFF&
Global Const GC_Azul = &HFF0000
Global Const GC_AzulEscuro = &H8000000D
Global Const GC_Preto = &H80000008
Global Const GC_Branco = &H80000005
Global Const GC_Cinza = &HC0C0C0
Global Const GC_Verde = &H8000&

'Constantes das Máscaras
Global Const GC_Formato_Cep = "#####\-###"
Global Const GC_Formato_Data = "dd/mm/yyyy"
Global Const GC_Formato_Hora = "hh:mm:ss"
Global Const GC_Formato_CGC = "@@.@@@.@@@/@@@@-@@"
Global Const GC_Formato_CPF = "@@@.@@@.@@@-@@"
Global Const GC_Formato_Fone = "(@@)@@@@-@@@@"
Global Const GC_Formato_Moeda = "###,###,##0.00"
Global Const GC_Chave = "322224"

'Variáveis de Conexão com o Banco de Dados
Public dbGranSecurity   As New ADODB.Connection
Public dbBancoDados     As New ADODB.Connection


'Variaveis Globais
Public G_rsGlobal   As ADODB.Recordset
Public G_sCampos    As String
Public G_sValores   As String
Public G_sQry       As String
Public G_User       As String
Public G_UserWin    As String
Public G_TpUser     As String

Public G_NomeEmpresa    As String
Public G_EndEmpresa     As String
Public G_BaiEmpresa     As String
Public G_CidEmpresa     As String
Public G_EstEmpresa     As String
Public G_CEPEmpresa     As String
Public G_PaisEmpresa    As String
Public G_FoneEmpresa    As String
Public G_FaxEmpresa     As String
Public G_EmailEmpresa   As String

Public G_RegAfetados As Integer
Public G_ColConsulta(5) As String 'Conteúdo das Colunas do FlexGrid da Tela de Consulta

'Variavéis do Arquivo ".INI"
Public G_Servidor   As String
Public G_NomeBanco  As String

'Variavéis do Arquivo "pwdgrtion.INI"
Public G_SenhaUser  As String


'Variáveis para o Formulário de Consulta ( FrmConsulta )
Public G_sQryBusca As String
Public G_rsBusca  As ADODB.Recordset

'Constantes de verificação de teclas
Public Enum eMascara
    eMascTudo = 0
    eMascLetra = 1
    eMascMoeda = 2
    eMascNumero = 4
    eMascNota = 8
    eMascMinusculas = 16
    eMascMaiusculas = 32
    eMascNumPonto = 64
    
End Enum

'constantes para Mensagens padrão
Public Enum eMsg
    eMsgAcentos = 0
    eMsgAlteracaoRegistro = 1
    eMsgArquivoNaoEncontrado = 2
    eMsgInexistente = 3
    eMsgDadosExcluidos = 6
    eMsgDadosSalvos = 8
    eMsgDataInvalida = 10
    eMsgDataDeveSerMenorIgual = 11
    eMsgErroExclusao = 12
    eMsgErroGravacao = 13
    eMsgErroInstalacao = 14
    eMsgExclusaoRegistro = 16
    eMsgImpressoraNaoInstalada = 19
    eMsgInclusaoRegistro = 20
    eMsgInconsistenciaRelatorio = 21
    eMsgInformeCampo = 22
    eMsgInformeCamposObrigatorios = 23
    eMsgInformePeloMenosUmCampo = 25
    eMsgMediaInvalida = 26
    eMsgNaoExistemInformacoesExcluir = 27
    eMsgRelatorioSemDados = 28
    eMsgRelatorioIndisponivel = 30
    eMsgSelecioneRelatorio = 31
    eMsgSemRegistros = 32
    eMsgTamanhoCampo = 34
    eMsgTelaEmConstrucao = 35
    eMsgTelaIndisponivel = 37
    eMsgUsuarioLogado = 38
    eMsgUsuarioNaoAutorizado = 39
    eMsgCampoNumerico = 50
    eMsgValorMoedaInvalido = 51
    eMsgRegistroJaExiste = 52
    eMsgExclusaoIntegridade = 53
    eMsgRegistroJaExcluido = 54
End Enum

Public Const GC_NomeSistema = "Transacoes"

Public Enum eOpcao
    eUF = 1
    eConstantes = 2
    eItemConstante = 3
End Enum

Public Sub PG_Monta_Combo(ByRef cmbCombo As ComboBox, ByVal eOp As eOpcao, _
    Optional ByVal sVar1 As String, Optional ByVal sVar2 As String)

    On Error GoTo Fim

    'Procedimento que lê registros e popula a combo passada por cmbCampus,
    'de acordo com eOp
    Dim rstConsulta As ADODB.Recordset
    Dim sQuery As String
    Dim sMascara1   As String
    Dim sMascara2   As String
    Dim sConteudo   As String
    Dim sValor      As String, iLaco        As Integer
    Dim iPosicao1   As Integer, iPosicao2   As Integer

    sConteudo = cmbCombo.Text
    Select Case eOp
    
'    Case eAnoSem
'        sMascara1 = "00000"
'        sQuery = "SELECT DISTINCT(pfb0anoref) as Campo1,'' as Campo2 "
'        sQuery = sQuery & "From "
'        sQuery = sQuery & "PFB0PERLET "
'        sQuery = sQuery & "ORDER BY PFB0ANOREF "

    Case eUF
        cmbCombo.AddItem "AC"
        cmbCombo.AddItem "AL"
        cmbCombo.AddItem "AM"
        cmbCombo.AddItem "AP"
        cmbCombo.AddItem "BA"
        cmbCombo.AddItem "CE"
        cmbCombo.AddItem "DF"
        cmbCombo.AddItem "ES"
        cmbCombo.AddItem "EX" '- Exterior
        cmbCombo.AddItem "GO"
        cmbCombo.AddItem "MA"
        cmbCombo.AddItem "MG"
        cmbCombo.AddItem "MS"
        cmbCombo.AddItem "MT"
        cmbCombo.AddItem "PA"
        cmbCombo.AddItem "PB"
        cmbCombo.AddItem "PE"
        cmbCombo.AddItem "PI"
        cmbCombo.AddItem "PR"
        cmbCombo.AddItem "RJ"
        cmbCombo.AddItem "RN"
        cmbCombo.AddItem "RO"
        cmbCombo.AddItem "RR"
        cmbCombo.AddItem "RS"
        cmbCombo.AddItem "SC"
        cmbCombo.AddItem "SE"
        cmbCombo.AddItem "SP"
        cmbCombo.AddItem "TO"
        cmbCombo.ListIndex = FG_Indice_Combo(cmbCombo, "MA")
        Exit Sub
    
    Case eConstantes
        'Filtra apenas as constantes que podem ser cadastradas pelo usuário
        'a primeira constante deve ser cadastrada pelo NTI
        sQuery = "Select VM02Camp03 As Campo1, VM02Chave As Campo2 "
        sQuery = sQuery & "From  VM02Constante "
        sQuery = sQuery & "Where VM02Camp02 = 'V' "
        sQuery = sQuery & "Group by VM02Camp03, VM02Chave "
        sQuery = sQuery & "Order by VM02Camp03"

    Case eItemConstante
        sMascara2 = "000"
        'Filtra apenas as constantes que podem ser cadastradas pelo usuário
        'a primeira constante deve ser cadastrada pelo NTI
        'VM02Camp01 = Descrição
        'VM02IteCha = Código
        sQuery = "Select VM02Camp01 As Campo1, VM02IteCha As Campo2 "
        sQuery = sQuery & "From  VM02Constante "
        sQuery = sQuery & "Where VM02Camp02 = 'V' And "
        sQuery = sQuery & "      VM02Chave = '" & sVar1 & "' "
        sQuery = sQuery & "Order by VM02Camp01"

    End Select

    Set rstConsulta = dbGranSecurity.Execute(sQuery)

    cmbCombo.Clear
    Do Until rstConsulta.EOF
       If Trim(rstConsulta!Campo2) <> "" Then
          cmbCombo.AddItem IIf(sMascara1 <> "", Format$(Trim$(rstConsulta!Campo1), sMascara1), _
          Trim$(rstConsulta!Campo1)) & " «» " & _
          IIf(sMascara2 <> "", Format$(Trim$(rstConsulta!Campo2), sMascara2), Trim$(rstConsulta!Campo2))
       Else
          If sMascara1 <> "" Then
             cmbCombo.AddItem Format$(Trim$(rstConsulta!Campo1), sMascara1)
          Else
             cmbCombo.AddItem Trim$(rstConsulta!Campo1)
          End If
       End If
       rstConsulta.MoveNext
    Loop

    If sConteudo <> "" Then
        cmbCombo.ListIndex = FG_Indice_Combo(cmbCombo, sConteudo)
    End If

Fim:
        
End Sub

Public Function FG_Carrega_Empresa(ByVal sUsuario As String, ByVal sSenha As String) As Boolean
    'Carrega os Dados da Empresa Cliente do GranSecurity

    Set dbGranSecurity = CreateObject("ADODB.Connection")
    dbGranSecurity.ConnectionString = ("DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & G_Servidor & ";PORT=3306;DATABASE=GranSecurity;USER=" & sUsuario & ";PASSWORD=" & sSenha)

    dbGranSecurity.ConnectionTimeout = 60
    dbGranSecurity.CommandTimeout = 400
    dbGranSecurity.CursorLocation = adUseClient
    dbGranSecurity.Open

    'Dados da Empresa
    G_sQry = ""
    G_sQry = G_sQry & "Select * From GS05TABEMP Where GS05EMPCOD = 1"
    Set G_rsGlobal = dbGranSecurity.Execute(G_sQry)
    If G_rsGlobal.EOF Then
        MsgBox "Empresa Não Cadastrada.", vbCritical, "A T E N Ç Ã O"
        FG_Carrega_Empresa = False
    Else
        G_NomeEmpresa = G_rsGlobal!GS05EMPNOM
        G_EndEmpresa = G_rsGlobal!GS05EMPEND
        G_BaiEmpresa = G_rsGlobal!GS05EMPBAI
        G_CidEmpresa = G_rsGlobal!GS05EMPCID
        G_EstEmpresa = G_rsGlobal!GS05EMPEST
        G_CEPEmpresa = Format(G_rsGlobal!GS05EMPCEP, GC_Formato_Cep)
        G_PaisEmpresa = G_rsGlobal!GS05EMPPAI
        G_FoneEmpresa = G_rsGlobal!GS05EMPTEL
        G_FaxEmpresa = G_rsGlobal!GS05EMPFAX
        G_EmailEmpresa = G_rsGlobal!GS05EMPEML
        FG_Carrega_Empresa = True
    End If
    G_rsGlobal.Close

End Function

Public Function FG_Carrega_Usuario(ByVal sUsuario As String) As Boolean
    'Carrega os Dados do Usuário no GranSecurity

    'Usuário Admin do Banco de Dados
    If sUsuario = "root" Then
        FG_Carrega_Usuario = True
        G_User = "ADMIN"
        G_TpUser = "M" '(C)omum ou (M)aster
        Exit Function
    End If

    'Dados do Usuário
    G_sQry = ""
    G_sQry = G_sQry & "Select * From GS01TABUSU Where GS01USUCOD = '" & sUsuario & "'"
    If GC_NomeSistema <> "Gransecurity" Then
        Set G_rsGlobal = dbGranSecurity.Execute(G_sQry)
    Else
        Set G_rsGlobal = dbBancoDados.Execute(G_sQry)
    End If
    If G_rsGlobal.EOF Then
        FG_Carrega_Usuario = False
    Else
        G_User = G_rsGlobal!GS01USUCOD
        G_TpUser = G_rsGlobal!GS01USUTIP '(C)omum ou (M)aster
        
        'Senha Expirada
        If FG_Conv_Data(G_rsGlobal!GS01USUEXP) < FG_Conv_Data(FG_Data_Servidor) Then
            MsgBox "Acesso Expirado! Contate o Administrador.", vbInformation, "ATENÇÃO"
            FG_Carrega_Usuario = False
            Exit Function
        'Senha Irá Expirar Então Avisa
        ElseIf DateDiff("d", FG_Data_Servidor, G_rsGlobal!GS01USUEXP) <= 10 Then
            MsgBox "Seu acesso expira em " & DateDiff("d", FG_Data_Servidor, G_rsGlobal!GS01USUEXP) & " dias! Contate o Administrador.", vbInformation, "ATENÇÃO"
        End If

        FG_Carrega_Usuario = True
    End If
    G_rsGlobal.Close

End Function


Function FG_Tira_Acento(Palavra As String) As String
    Dim i As Integer
    Dim Retorno As String
    Retorno = ""
    For i = 1 To Len(Trim(Palavra))
        Select Case Mid(Palavra, i, 1)
            Case "Á", "Ã", "À", "Â"
                Retorno = Retorno & "A"
            Case "É", "Ê"
                Retorno = Retorno & "E"
            Case "Í", "Î", "Ì"
                Retorno = Retorno & "I"
            Case "Ó", "Ô", "Õ", "Ò"
                Retorno = Retorno & "O"
            Case "Ú", "Û", "Ù"
                Retorno = Retorno & "U"
            Case "Ç"
                Retorno = Retorno & "C"
            Case Else
                Retorno = Retorno & Mid(Palavra, i, 1)
        End Select
    Next
    FG_Tira_Acento = Retorno

End Function

Public Function FG_Extenso(cValor As Currency) As String

    Dim sExt As String
    Dim vPal_1 As Variant
    Dim vPal_2 As Variant
    Dim vPal_3 As Variant
    Dim sLiteral As String
    
    
    Dim iTrilhao As Integer
    Dim iBilhao  As Integer
    Dim iMilhao  As Integer
    Dim iMil  As Integer
    Dim iMoeda  As Integer
    Dim iCentavo  As Integer
    
    Dim iVar_1 As Integer
    Dim iVar_2_3  As Integer
    Dim iVar_2  As Integer
    Dim iVar_3  As Integer
    Dim iVar  As Integer

    vPal_1 = Array("  ", "UM ", "DOIS ", "TRÊS ", "QUATRO ", "CINCO ", "SEIS ", "SETE ", "OITO ", "NOVE ", _
                  "DEZ ", "ONZE ", "DOZE ", "TREZE ", "QUATORZE ", "QUINZE ", "DEZESSEIS ", "DEZESSETE ", "DEZOITO ", "DEZENOVE ")
    vPal_2 = Array("    ", "    ", "VINTE ", "TRINTA ", "QUARENTA ", "CINQUENTA ", "SESSENTA ", "SETENTA ", "OITENTA ", "NOVENTA ")
    vPal_3 = Array("     ", "     ", "DUZENTOS ", "TREZENTOS ", "QUATROCENTOS ", "QUINHENTOS ", "SEISCENTOS ", "SETECENTOS ", "OITOCENTOS ", "NOVECENTOS ")
    
    sExt = ""

    sLiteral = CStr(Format(cValor, "000000000000000.00"))
    'Let sLiteral = CStr(Format(cValor, "##,##0.00"))

    iTrilhao = Val(Mid(sLiteral, 1, 3))
    iBilhao = Val(Mid(sLiteral, 4, 3))
    iMilhao = Val(Mid(sLiteral, 7, 3))
    iMil = Val(Mid(sLiteral, 10, 3))
    iMoeda = Val(Mid(sLiteral, 13, 3))
    iCentavo = Val(Mid(sLiteral, 17, 2))

P100_trilhoes:
    If iTrilhao = 0 Then GoTo P110_bilhoes
    sLiteral = CStr(Format(iTrilhao, "000"))
    GoSub separa 'Simula o redefines do cobol
    GoSub escreve_extenso
    
    If iVar = 1 Then
       sExt = sExt & "TRILHÃO "
    Else
      sExt = sExt & "TRILHÕES "
    End If

P110_bilhoes:
    If iBilhao = 0 Then GoTo P120_milhoes
    sLiteral = CStr(Format(iBilhao, "000"))
    GoSub separa
    GoSub escreve_extenso
    If iVar = 1 Then
       sExt = sExt & "BILHÃO "
    Else
       sExt = sExt & "BILHÕES "
    End If

P120_milhoes:
    If iMilhao = 0 Then GoTo P130_milhares
    sLiteral = CStr(Format(iMilhao, "000"))
    GoSub separa
    GoSub escreve_extenso
    If iVar = 1 Then
       sExt = sExt & "MILHÃO "
    Else
       sExt = sExt & "MILHÕES "
    End If

P130_milhares:
    If iMil = 0 Then GoTo P140_iMoedas
    sLiteral = CStr(Format(iMil, "000"))
    GoSub separa
    GoSub escreve_extenso
    sExt = sExt & "iMil "

P140_iMoedas:
    If cValor < CCur(1#) Then GoTo P150_iCentavos
    sLiteral = CStr(Format(iMoeda, "000"))
    GoSub separa
    If iVar > 0 Then GoSub escreve_extenso
    If cValor > CCur(999999.99) And iMil = 0 And iMoeda = 0 Then
      sExt = sExt & "DE "
    End If
    If iVar = 1 And cValor < CCur(2#) Then
      sExt = sExt & "REAL "
    Else
      sExt = sExt & "REAIS "
    End If

P150_iCentavos:
    If iCentavo = 0 Then GoTo Fim
    If cValor > CCur(0.99) Then
      sExt = sExt & "E "
    End If
    sLiteral = CStr(Format(iCentavo, "000"))
    GoSub separa
    GoSub escreve_extenso
    If iVar = 1 Then
      sExt = sExt & "iCentavo "
    Else
      sExt = sExt & "iCentavoS "
    End If

Fim:
    FG_Extenso = sExt
    Exit Function

separa:
    Let iVar_1 = Val(Mid(sLiteral, 1, 1))
    Let iVar_2_3 = Val(Mid(sLiteral, 2, 2))
    Let iVar_2 = Val(Mid(sLiteral, 2, 1))
    Let iVar_3 = Val(Mid(sLiteral, 3, 1))
    Let iVar = (iVar_1 * 100) + iVar_2_3
    Return

escreve_extenso:
    If iVar_1 > 0 Then GoSub tabela_tres
    If iVar_2_3 = 0 Then Return
    If iVar_1 > 0 Then
      sExt = sExt & "E "
    End If
    If iVar_2_3 < 20 Then
       sExt = sExt & vPal_1(iVar_2_3)
       Return
    Else
       sExt = sExt & vPal_2(iVar_2)
    End If
    If iVar_3 > 0 Then
      sExt = sExt & "E "
      sExt = sExt & vPal_1(iVar_3)
    End If

    Return

tabela_tres:
    If iVar_1 <> 1 Then
       sExt = sExt & vPal_3(iVar_1)
    End If

    If iVar = 100 And iCentavo = 0 And sExt <> "" Then
       sExt = sExt & "E "
    End If
    
    If iVar = 100 Then
       sExt = sExt & "CEM "
    Else
       If iVar_1 = 1 Then
          sExt = sExt & "CENTO "
       End If
    End If
    
    Return

End Function


Public Sub PG_CentralizaForm(NomeForm As Form)
    NomeForm.Left = (Screen.Width - NomeForm.Width) / 2
    NomeForm.Top = (Screen.Height - NomeForm.Height) / 8

End Sub

Function FG_Tecla_Atalho(Cod As Integer, WShift As Integer) As String
    
    Dim ShiftDown, AltDown, CtrlDown, Txt
    ShiftDown = (WShift And vbShiftMask) > 0
    AltDown = (WShift And vbAltMask) > 0
    CtrlDown = (WShift And vbCtrlMask) > 0

    Dim Tecla As String
    
    If ShiftDown And CtrlDown And AltDown Then
        Txt = "SHIFT+CTRL+ALT"
    ElseIf ShiftDown And AltDown Then
        Txt = "SHIFT+ALT"
    ElseIf ShiftDown And CtrlDown Then
        Txt = "SHIFT+CTRL"
    ElseIf CtrlDown And AltDown Then
        Txt = "CTRL+ALT"
    ElseIf ShiftDown Then
        Txt = "SHIFT"
    ElseIf CtrlDown Then
    Txt = "CTRL"
    ElseIf AltDown Then
        Txt = "ALT"
    ElseIf WShift = 0 Then
        Txt = ""
    End If
    
    Tecla = ""
    Select Case Cod
        Case vbKeyF1
            Tecla = "F1"
        Case vbKeyF2
            Tecla = "F2"
        Case vbKeyF3
            Tecla = "F3"
        Case vbKeyF4
            Tecla = "F4"
        Case vbKeyF5
            Tecla = "F5"
        Case vbKeyF6
            Tecla = "F6"
        Case vbKeyF7
            Tecla = "F7"
        Case vbKeyF8
            Tecla = "F8"
        Case vbKeyF9
            Tecla = "F9"
        Case vbKeyF10
            Tecla = "F10"
        Case vbKeyF11
            Tecla = "F11"
        Case vbKeyF12
            Tecla = "F12"
        Case 27
            Tecla = "ESC"
        Case Else
            Tecla = Chr(Cod)
    End Select
    
    If Txt <> "" And Tecla <> "" Then
           FG_Tecla_Atalho = Txt & IIf(Tecla <> "", "+" & Tecla, "")
    ElseIf Txt <> "" And Tecla = "" Then
        FG_Tecla_Atalho = Txt
    ElseIf Txt = "" And Tecla <> "" Then
        FG_Tecla_Atalho = Tecla
    End If
        
End Function

Function FG_Alinha_Esquerda(sPalavra As String, iTamanho As Integer) As String
    Dim sEspaco As String
    sEspaco = String(iTamanho, " ")
    'Faz alinhamento a esquerda
    LSet sEspaco = sPalavra
    FG_Alinha_Esquerda = sEspaco

End Function

Function FG_Alinha_Direita(sPalavra As String, iTamanho As Integer) As String
    Dim sEspaco As String
    sEspaco = String(iTamanho, " ")
    'Faz alinhamento a esquerda
    RSet sEspaco = sPalavra
    FG_Alinha_Direita = sEspaco

End Function

Function FG_Troca_Sequencia(ByVal Indice As Long, Sequencia As String) As String
         Dim i As Integer
         Dim PosVirgulaAtual As Integer
         Dim PosVirgulaAnterior As Integer
         Dim NumVirgulasLidas As Integer
         
         'Adiciona mais uma virgula para facilitar o calculo
         Sequencia = Trim(Sequencia) & ","
         
         'Se Foi passado o indice 0
         If Indice = 0 Then
            FG_Troca_Sequencia = ""
            Exit Function
         End If
         
         NumVirgulasLidas = 1
         PosVirgulaAnterior = 0
         
         For i = 1 To Len(Trim(Sequencia))
             
             'Intercepta a posicao da virgula
             If Mid(Sequencia, i, 1) = "," Then
                
                PosVirgulaAtual = i
                
                
                'testa se a virgula corresponde ao indice da sequencia
                If NumVirgulasLidas = Indice Then
                   FG_Troca_Sequencia = Mid(Sequencia, PosVirgulaAnterior + 1, (PosVirgulaAtual - PosVirgulaAnterior) - 1)
                   Exit Function
                Else
                   PosVirgulaAnterior = PosVirgulaAtual
                   NumVirgulasLidas = NumVirgulasLidas + 1
                End If
                
             ElseIf Mid(Sequencia, i, 1) = " " Then
                   FG_Troca_Sequencia = " "
                   Exit Function
             End If
         Next

FG_Troca_Sequencia = " "
End Function

Public Sub PG_Limpa_Campos(Controles As Object, Optional bEsconderControle As Boolean = False, Optional Nome_Controle = "", Optional Nome_Controle2 = "", Optional Nome_Controle3 = "", Optional Nome_Controle4 = "", Optional Nome_Controle5 = "", Optional Nome_Controle6 = "", Optional Nome_Controle7 = "", Optional Nome_Controle8 = "", Optional Nome_Controle9 = "", Optional Nome_Controle10 = "", Optional Nome_Controle11 = "", Optional Nome_Controle12 = "", Optional Nome_Controle13 = "", Optional Nome_Controle14 = "", Optional Nome_Controle15 = "", Optional Nome_Controle16 = "")

    On Error Resume Next
    Dim Mascara As String
   'O parametro Nome_Controle serve para indicar algum controle
   'que nao deve ser limpo.
   
   'O parametro bEsconderControle Torna o Controle Invisível e Limpa o ToolTipText do mesmo

    Dim curControl As Control
    'Limpa TextBox / Labels e MaskEdits
    
    For Each curControl In Controles
        If UCase(curControl.Name) <> UCase(Nome_Controle) And _
          UCase(curControl.Name) <> UCase(Nome_Controle2) And _
          UCase(curControl.Name) <> UCase(Nome_Controle3) And _
          UCase(curControl.Name) <> UCase(Nome_Controle4) And _
          UCase(curControl.Name) <> UCase(Nome_Controle5) And _
          UCase(curControl.Name) <> UCase(Nome_Controle6) And _
          UCase(curControl.Name) <> UCase(Nome_Controle7) And _
          UCase(curControl.Name) <> UCase(Nome_Controle8) And _
          UCase(curControl.Name) <> UCase(Nome_Controle9) And _
          UCase(curControl.Name) <> UCase(Nome_Controle10) And _
          UCase(curControl.Name) <> UCase(Nome_Controle11) And _
          UCase(curControl.Name) <> UCase(Nome_Controle12) And _
          UCase(curControl.Name) <> UCase(Nome_Controle13) And _
          UCase(curControl.Name) <> UCase(Nome_Controle14) And _
          UCase(curControl.Name) <> UCase(Nome_Controle15) And _
          UCase(curControl.Name) <> UCase(Nome_Controle16) And _
          TypeOf curControl Is Frame = False And _
          TypeOf curControl Is Toolbar = False Then

            If bEsconderControle Then
                curControl.Visible = False
                curControl.ToolTipText = ""
            End If

            If TypeOf curControl Is TextBox Then
               curControl.Text = Space$(0)
    '            curControl.BackColor = vbWhite
            ElseIf TypeOf curControl Is MaskEdBox Then
                   Mascara = curControl.Mask
                   curControl.Mask = ""
                   curControl.Text = ""
                   curControl.Mask = Mascara
            ElseIf TypeOf curControl Is OptionButton Then
                   curControl.Value = False
            ElseIf TypeOf curControl Is Image Then
                   curControl.Picture = Nothing
            ElseIf TypeOf curControl Is CheckBox Then
                   curControl.Value = False
            ElseIf TypeOf curControl Is ComboBox Then
                   curControl.ListIndex = -1
            ElseIf TypeOf curControl Is Label Then
                If UCase(Mid(curControl.Name, 1, 3)) = "LBL" Then 'Para Limpar somente Labels Variáveis
                   'curControl.Caption = Space$(0)
                End If
            End If
       End If
     Next

End Sub

Function FG_Form_Carregado(fFormName As Form) As Boolean
    Dim i As Integer
   
    For i = 0 To Forms.Count - 1
       If Forms(i) Is fFormName Then
          FG_Form_Carregado = True
          Exit Function
       End If
    Next
     FG_Form_Carregado = False

End Function

Public Function FG_Valida_CPF(sCpf As String, Optional bMensagem As Boolean = True) As Boolean
    ' Valida o CPF
     'Obs. Os parametros devem ser passados sem nenhuma pontuação!
    'Dim WdigitoDoCPF
    Dim wSomaDosProdutos
    Dim wResto
    Dim wDigitChk1
    Dim wDigitChk2
    Dim wStatus
    Dim wI
    
    wSomaDosProdutos = 0
    
    For wI = 1 To 9
        wSomaDosProdutos = wSomaDosProdutos + Val(Mid(sCpf, wI, 1)) * (11 - wI)
    Next wI
    wResto = wSomaDosProdutos - Int(wSomaDosProdutos / 11) * 11
    wDigitChk1 = IIf(wResto = 0 Or wResto = 1, 0, 11 - wResto)
    
    wSomaDosProdutos = 0
    For wI = 1 To 9
        wSomaDosProdutos = wSomaDosProdutos + (Val(Mid(sCpf, wI, 1)) * (12 - wI))
    Next wI
    wSomaDosProdutos = wSomaDosProdutos + (2 * wDigitChk1)
    wResto = wSomaDosProdutos - Int(wSomaDosProdutos / 11) * 11
    wDigitChk2 = IIf(wResto = 0 Or wResto = 1, 0, 11 - wResto)
    
    If Mid(sCpf, 10, 1) = Mid(Trim(Str(wDigitChk1)), 1, 1) And Mid(sCpf, 11, 1) = Mid(Trim(Str(wDigitChk2)), 1, 1) Then
        FG_Valida_CPF = True
    Else
        FG_Valida_CPF = False
        If bMensagem Then
            MsgBox "CPF inválido", vbExclamation, "Aviso"
        End If
    End If

End Function

Public Function FG_Valida_CNPJ(ByVal CNPJ As String, Optional bMensagem As Boolean = True) As Integer

    Dim Soma, Digito, CNPJ1, CNPJ2, MULT, Controle, j, i

    Soma = 0
    Digito = 0

    CNPJ1 = Left(CNPJ, 12)
    CNPJ2 = Right(CNPJ, 2)

    MULT = "543298765432"

    Controle = vbNullString

    For j = 1 To 2
        Soma = 0

        For i = 1 To 12
            Soma = Soma + Val(Mid(CNPJ1, i, 1)) * Val(Mid(MULT, i, 1))
        Next i

        If j = 2 Then Soma = Soma + (2 * Digito)

        Digito = ((Soma * 10) Mod 11)

        If Digito = 10 Then Digito = 0

        Controle = Controle + Trim$(Str$(Digito))

        MULT = "654329876543"
    Next j

    If Controle <> CNPJ2 Then
       FG_Valida_CNPJ = False
       If bMensagem Then
          MsgBox "CNPJ inválido", vbExclamation, "Aviso"
       End If
    Else
       FG_Valida_CNPJ = True
    End If
            
End Function

Function FG_ValidaHora(textoHora As String) As Boolean
    ' 1. Usa IsDate para garantir que a string é uma data/hora válida.
    If IsDate(textoHora) Then
        ' 2. Acessa a hora, minuto e segundo.
        Dim hora As Integer
        Dim minuto As Integer
        Dim segundo As Integer
        
        ' Extrai a hora, minuto e segundo
        hora = Hour(CDate(textoHora))
        minuto = Minute(CDate(textoHora))
        segundo = Second(CDate(textoHora))

        ' 3. Verifica se os valores estão dentro dos intervalos corretos
        If hora >= 0 And hora < 24 And _
           minuto >= 0 And minuto < 60 And _
           segundo >= 0 And segundo < 60 Then
            FG_ValidaHora = True
        Else
            FG_ValidaHora = False
        End If
    Else
        FG_ValidaHora = False
    End If

'    If Not FG_ValidaHora Then
'          MsgBox "Hora inválida", vbExclamation, "Aviso"
'    End If

End Function

Public Function FG_Modulo11(Numero As String) As String
  Dim i As Integer
  Dim Produto As Integer
  Dim Multiplicador As Integer
  Dim Digito As Integer
  ' Válida Argumento
  If Not IsNumeric(Numero) Then
    FG_Modulo11 = ""
    Exit Function
  End If
  ' Cálcula Dígito no Módulo 11
  Multiplicador = 2
  For i = Len(Numero) To 1 Step -1
    Produto = Produto + Val(Mid(Numero, i, 1)) * Multiplicador
    Multiplicador = IIf(Multiplicador = 9, 2, Multiplicador + 1)
  Next
  ' Exceção
  Digito = 11 - Int(Produto Mod 11)
  Digito = IIf(Digito = 10 Or Digito = 11, 0, Digito)
  ' Retorna
  FG_Modulo11 = Trim(Str(Digito))
End Function

Public Function ValidaCGC(Cgc As String) As Boolean
    'Obs. Os parametros devem ser passados sem nenhuma pontuação!
     ' Válida argumento
     If Len(Cgc) <> 14 Then
       ValidaCGC = False
       Exit Function
     End If
     ' Válida Primeiro Dígito
     If FG_Modulo11(Left(Cgc, 12)) <> Mid(Cgc, 13, 1) Then
       ValidaCGC = False
       Exit Function
     End If
     ' Válida Segundo Dígito
     If FG_Modulo11(Left(Cgc, 13)) <> Mid(Cgc, 14, 1) Then
       ValidaCGC = False
       Exit Function
     End If
     ValidaCGC = True
End Function

Sub PG_Configura_Icone_Barra(tlbBarra As Toolbar)
On Error GoTo Trata_Erro

  Set tlbBarra.ImageList = mdiPrincipal.imgCadastros
  
  With tlbBarra
     .Buttons("Novo").Image = "Novo"
     .Buttons("Salvar").Image = "Salvar"
     .Buttons("Excluir").Image = "Excluir"
     .Buttons("Procurar").Image = "Procurar"
     .Buttons("Imprimir").Image = "Imprimir"
     .Buttons("Fechar").Image = "Fechar"
     .Buttons("Ajuda").Image = "Ajuda"
     .Buttons("Selecionar").Image = "Selecionar"

  End With
  Exit Sub
  
Trata_Erro:
If Err.Number = ccElemNotFound Then
    Err.Clear
    Resume Next
End If

End Sub

Public Sub PG_Muda_Teclas(iTecla As Integer)
    If iTecla = vbKeyReturn Then
        SendKeys "{TAB}", True
        SendKeys "{NUMLOCK}", True
    End If

    'Modifica a Digitação de Aspas Simples
    If iTecla = 39 Then
      iTecla = 96
    End If

End Sub

Public Function FG_Campo_Nao_Informado(ByRef oObjeto1 As Object, Optional ByRef oObjeto2 As Object, Optional ByRef oObjeto3 As Object, Optional ByRef oObjeto4 As Object, Optional ByRef oObjeto5 As Object, Optional ByRef oObjeto6 As Object, Optional ByRef oObjeto7 As Object, Optional ByRef oObjeto8 As Object, Optional ByRef oObjeto9 As Object, Optional ByRef oObjeto10 As Object, Optional ByRef oObjeto11 As Object, Optional ByRef oObjeto12 As Object, Optional ByRef oObjeto13 As Object, Optional ByRef oObjeto14 As Object, Optional ByRef oObjeto15 As Object) As Boolean

    On Error GoTo TrataErro

    'Recebe os objetos da tela e testa se os que ele passou foram digitados,
    'caso contrario faz critica e cai o focu nele

    Dim oCorrente As Object, yLaco As Byte, bErro As Boolean

    FG_Campo_Nao_Informado = False
    yLaco = 1

    Do While yLaco <= 15
        bErro = False
        Select Case yLaco

        Case 1
            Set oCorrente = oObjeto1

        Case 2
            Set oCorrente = oObjeto2

        Case 3
            Set oCorrente = oObjeto3

        Case 4
            Set oCorrente = oObjeto4

        Case 5
            Set oCorrente = oObjeto5

        Case 6
            Set oCorrente = oObjeto6

        Case 7
            Set oCorrente = oObjeto7

        Case 8
            Set oCorrente = oObjeto8

        Case 9
            Set oCorrente = oObjeto9

        Case 10
            Set oCorrente = oObjeto10

        Case 11
            Set oCorrente = oObjeto11

        Case 12
            Set oCorrente = oObjeto12

        Case 13
            Set oCorrente = oObjeto13

        Case 14
            Set oCorrente = oObjeto14

        Case 15
            Set oCorrente = oObjeto15

        End Select

        If oCorrente Is Nothing Then Exit Function

        
        If Trim(oCorrente.Text) = "" Then
            If Not bErro Then
                MsgBox "Informe Campos Obrigatórios.", vbExclamation, "Mensagem"
                FG_Campo_Nao_Informado = True
                oCorrente.SetFocus
                Exit Function
            End If
            Err.Clear
        End If

        If oCorrente.ClipText = "" Then
            If Not bErro Then
                MsgBox "Informe Campos Obrigatórios.", vbExclamation, "Mensagem"
                FG_Campo_Nao_Informado = True
                oCorrente.SetFocus
                Exit Function
            End If
            Err.Clear
        End If

        If oCorrente.Value = "" Then
            If Not bErro Then
                MsgBox "Informe Campos Obrigatórios.", vbExclamation, "Mensagem"
                FG_Campo_Nao_Informado = True
                oCorrente.SetFocus
                Exit Function
            End If
            Err.Clear
        End If

        yLaco = yLaco + 1
    Loop

TrataErro:
    If Err.Number = 438 Then
        bErro = True
        Resume Next
    End If

End Function

Public Sub PG_Centraliza_form(lForm As Form)
  lForm.Top = 0
  lForm.Left = mdiPrincipal.Width / 2 - lForm.Width / 2

End Sub

Public Sub PG_Verifica_Tecla(ByRef iCodTecla As Integer, Optional iMasc As eMascara, Optional ByRef oObjeto As Object, Optional ByVal yDecimais As Byte = 2)
    'Procedimento que mascara a digitação, de acordo com a máscara informada

    Dim iPosicao    As Integer

    'NÃO ACEITA ACENTOS .
    If iCodTecla > 122 Then
       iCodTecla = 0
    End If

    If iCodTecla = Asc("'") Or iCodTecla = 34 Then
        'Se for aspas simples ou duplas, despreza digitação
        'devido a não ser aceito no SQL
        iCodTecla = 0
    ElseIf iCodTecla = vbKeyBack Then
        Exit Sub
    
    ElseIf iMasc = eMascNumero Then
        'Se for máscara NÚMERO, só aceita números de 0 a 9 e carac. controle
        If iCodTecla < vbKey0 Or iCodTecla > vbKey9 Then
            If iCodTecla >= vbKeySpace Then
                iCodTecla = 0
            End If
        End If
    
    ElseIf iMasc = eMascNumPonto Then
        'Se for máscara NÚMERO e Ponto, só aceita números de 0 a 9, (.) ponto, (-) traço e carac. controle
        If (iCodTecla < vbKey0 Or iCodTecla > vbKey9) And iCodTecla <> Asc(".") And iCodTecla <> Asc("-") Then
            If iCodTecla >= vbKeySpace Then
                iCodTecla = 0
            End If
        End If
    
    ElseIf iMasc = eMascMaiusculas Then
        'Se for máscara MAIÚSCULAS, converte tudo para maiúsculas
        iCodTecla = Asc(UCase(Chr(iCodTecla)))
    ElseIf iMasc = eMascMinusculas Then
        'Se for máscara MINÚSCULAS, converte tudo para minúsculas
        iCodTecla = Asc(LCase(Chr(iCodTecla)))
    ElseIf iMasc = eMascLetra Then
        'Se for máscara LETRA, só aceita letras, espaços e carac. controle
        If iCodTecla < vbKeyA Then
            If iCodTecla > vbKeySpace Then
               iCodTecla = 0
            End If
        End If
    ElseIf iMasc = eMascMoeda Then
        'Se for máscara MOEDA, só aceita números, vírgulas
        'Caso não seja informado o número de casas decimais após
        'a vírgula, considera 2
        If iCodTecla < vbKey0 Or iCodTecla > vbKey9 Then
            If iCodTecla = Asc(",") Then
               If oObjeto Is Nothing Then
               Else
                  If InStr(1, oObjeto.Text, ",") > 0 Then
                     iCodTecla = 0
                  End If
               End If
            Else
               iCodTecla = 0
            End If
        Else
            'Se for números, verifica se possui 2 números após a vírgula
            'se possuir, despreza a digitação
            If oObjeto Is Nothing Then
            Else
               iPosicao = InStr(1, oObjeto.Text, ",")
               If iPosicao > 0 Then
                  If Len(oObjeto.Text) - iPosicao >= yDecimais Then
                     If oObjeto.SelStart >= Len(oObjeto.Text) - yDecimais Then
                        iCodTecla = 0
                     End If
                  End If
               End If
            End If
        End If
    ElseIf iMasc = eMascNota Then
        'Se for máscara NOTA, só aceita S/N, números, vírgula, ponto e carac. controle
        'Nesta máscara é obrigatório a informação de oObjeto, que é o objeto
        'de digitação da nota (geralmente um TextBox)

'        If icodtecla = vbKeyReturn Then
'           icodtecla = FG_Valida_Formato(txtNotaAv.Text, "##.##", icodtecla)
'
        'Troca vírgula por ponto
        If iCodTecla = Asc(",") Then
            iCodTecla = Asc(".")
        'Se já digitou um S, despreza a digitação de outro S
        ElseIf iCodTecla = vbKeyS Then
            iCodTecla = 0
            oObjeto = "S/N"
            oObjeto.SelStart = 0
            oObjeto.SelLength = Len(oObjeto)
        'Se já digitou um /, despreza a digitação de outro /
        ElseIf iCodTecla = Asc("/") Then
            iCodTecla = 0
            oObjeto = "S/N"
            oObjeto.SelStart = 0
            oObjeto.SelLength = Len(oObjeto)
        'Se já digitou um N, despreza a digitação de outro N
        ElseIf iCodTecla = vbKeyN Then
            iCodTecla = 0
            oObjeto = "S/N"
            oObjeto.SelStart = 0
            oObjeto.SelLength = Len(oObjeto)
        ElseIf (iCodTecla <> vbKeyS) And _
               (iCodTecla <> vbKeyBack) And _
               (iCodTecla <> vbKeyReturn) And _
               (iCodTecla <> Asc(".")) And _
               (iCodTecla <> Asc("/")) And _
               (iCodTecla <> vbKeyN) And _
               (iCodTecla < vbKey0) Or _
               (iCodTecla > vbKey9) Then
               iCodTecla = 0
        End If
        'Se já existir ponto, despreza a nova digitação de .
        If iCodTecla = Asc(".") And _
           InStr(1, oObjeto, ".") > 0 Then
           iCodTecla = 0
        End If
    End If
    
End Sub

Public Function FG_Retirar_Atalho(Nome As String) As String
   'Função que retira atalhos "&" de Caption, converte para maiúsculas
   'e retira os acentos
   Dim Posicao As Integer
   
   Nome = FG_RetirarAcentos(Nome)
   Posicao = InStr(1, Nome, "&")
   If Posicao > 0 Then
      Nome = Mid(Nome, 1, Posicao - 1) & Mid(Nome, Posicao + 1, Len(Nome) - Posicao)
   End If
   FG_Retirar_Atalho = Nome

End Function

Public Function FG_RetirarAcentos(Nome As String) As String
   'Função que retira acentos e retorna em maiúsculas
   Dim iiLaco As Integer
   
   Nome = UCase(Nome)
   For iiLaco = 1 To Len(Nome)
      If Mid(Nome, iiLaco, 1) = "À" Or _
         Mid(Nome, iiLaco, 1) = "Á" Or _
         Mid(Nome, iiLaco, 1) = "Ã" Or _
         Mid(Nome, iiLaco, 1) = "Â" Or _
         Mid(Nome, iiLaco, 1) = "Ä" Then
         Nome = Left(Nome, iiLaco - 1) & "A" & Right(Nome, Len(Nome) - iiLaco)
      ElseIf Mid(Nome, iiLaco, 1) = "È" Or _
         Mid(Nome, iiLaco, 1) = "É" Or _
         Mid(Nome, iiLaco, 1) = "Ê" Or _
         Mid(Nome, iiLaco, 1) = "Ë" Then
         Nome = Left(Nome, iiLaco - 1) & "E" & Right(Nome, Len(Nome) - iiLaco)
      ElseIf Mid(Nome, iiLaco, 1) = "Ì" Or _
         Mid(Nome, iiLaco, 1) = "Í" Or _
         Mid(Nome, iiLaco, 1) = "Î" Or _
         Mid(Nome, iiLaco, 1) = "Ï" Then
         Nome = Left(Nome, iiLaco - 1) & "I" & Right(Nome, Len(Nome) - iiLaco)
      ElseIf Mid(Nome, iiLaco, 1) = "Ò" Or _
         Mid(Nome, iiLaco, 1) = "Ó" Or _
         Mid(Nome, iiLaco, 1) = "Õ" Or _
         Mid(Nome, iiLaco, 1) = "Ô" Or _
         Mid(Nome, iiLaco, 1) = "Ö" Then
         Nome = Left(Nome, iiLaco - 1) & "O" & Right(Nome, Len(Nome) - iiLaco)
      ElseIf Mid(Nome, iiLaco, 1) = "Ù" Or _
         Mid(Nome, iiLaco, 1) = "Ú" Or _
         Mid(Nome, iiLaco, 1) = "Û" Or _
         Mid(Nome, iiLaco, 1) = "Ü" Then
         Nome = Left(Nome, iiLaco - 1) & "U" & Right(Nome, Len(Nome) - iiLaco)
      ElseIf Mid(Nome, iiLaco, 1) = "Ç" Then
         Nome = Left(Nome, iiLaco - 1) & "C" & Right(Nome, Len(Nome) - iiLaco)
      End If
   Next
   FG_RetirarAcentos = UCase(Nome)

End Function

Public Function FG_MsgBoxPadrao(iCodMsg As eMsg, Optional vCampo1 As Variant, Optional vCampo2 As Variant) As Long

    'Função que exibe mensagens. Dependendo do tipo de mensagem poder utilizar
    'as variáveis vCampo1, vCampo2
    'Em alguns casos, retorna valores de resposta.

    Select Case iCodMsg

    Case eMsgAcentos
        MsgBox "Por favor, não informe acentos" & vbCrLf & "ou caracteres Ç, " & Chr(34) & " ou '", vbExclamation, "Aviso"

    Case eMsgAlteracaoRegistro
        FG_MsgBoxPadrao = MsgBox("Confirma alteração dos dados?", vbQuestion + vbYesNo + vbDefaultButton1, "Alteração")

    Case eMsgArquivoNaoEncontrado
        MsgBox "O arquivo " & vCampo1 & " não foi encontrado." & vbCrLf & "Por favor, verifique.", vbExclamation, "Arquivo não encontrado"

    Case eMsgCampoNumerico
        MsgBox "Esse campo possui formato numérico." & vbCrLf & "Digite apenas números.", vbExclamation

    Case eMsgDataInvalida
        MsgBox "Data Inválida", vbExclamation, "Aviso"

    Case eMsgDataDeveSerMenorIgual
        If IsMissing(vCampo1) Then
           MsgBox "Data deve ser menor ou igual à data de hoje.", vbExclamation, "Aviso"
        Else
           MsgBox "Data " & vCampo1 & " deve ser menor ou igual à data de hoje.", vbExclamation, "Aviso"
        End If

    Case eMsgDadosExcluidos
        MsgBox "Dados excluídos com sucesso", vbInformation, "Informação"

    Case eMsgDadosSalvos
        MsgBox "Dados salvos com sucesso", vbInformation, "Informação"

    Case eMsgErroGravacao, -2147217873
        If IsMissing(vCampo1) Then
            MsgBox "Erro na tentativa de gravar. " & vbCr & "Tente novamente mais tarde.", vbExclamation, "Aviso"
        Else
            MsgBox "Erro na tentativa de gravar. (Erro " & vCampo1 & ")" & vbCr & "Tente novamente mais tarde.", vbExclamation, "Aviso"
        End If

    Case eMsgErroExclusao
        If IsMissing(vCampo1) Then
            MsgBox "Erro na tentativa de excluir. " & vbCr & "Tente novamente mais tarde.", vbExclamation, "Aviso"
        Else
            MsgBox "Erro na tentativa de excluir. (Erro " & vCampo1 & ")" & vbCr & "Tente novamente mais tarde.", vbExclamation, "Aviso"
        End If

    Case eMsgErroInstalacao
        If IsMissing(vCampo1) Then
            MsgBox "Falta(m) algum(ns) arquivo(s) no seu computador" & vbCr & "para utilizar esta função." & vbCr & "Por favor, entre em contato com o Suporte Técnico.", vbExclamation, "Aviso"
        Else
            MsgBox "Falta(m) algum(ns) arquivo(s) no seu computador" & vbCr & "para utilizar esta função." & vbCr & "Por favor, entre em contato com o Suporte Técnico." & vbCrLf & "Erro: " & vCampo1, vbExclamation, "Aviso"
        End If

    Case eMsgExclusaoRegistro
        FG_MsgBoxPadrao = MsgBox("Confirma exclusão dos dados ?", vbQuestion + vbYesNo + vbDefaultButton2, "Exclusão")

    Case eMsgImpressoraNaoInstalada
        MsgBox "Não existe nenhuma impressora instalada no Windows.", vbInformation, "Aviso"

    Case eMsgInclusaoRegistro
        FG_MsgBoxPadrao = MsgBox("Confirma inclusão ?", vbQuestion + vbYesNo + vbDefaultButton1, "Inclusão")

    Case eMsgInconsistenciaRelatorio
        FG_MsgBoxPadrao = MsgBox("O relatório não pode ser exibido." & vbCrLf & "Favor entrar em contato com o CPD. Erro 171", vbExclamation, "Relatório")
    
    Case eMsgInexistente
        MsgBox vCampo1 & " não existe no sistema.", vbExclamation, "Aviso"

    Case eMsgInformeCampo
        MsgBox "Informe " & vCampo1, vbExclamation, "Aviso"

    Case eMsgInformeCamposObrigatorios
        MsgBox "Informe campos obrigatórios em vermelho", vbExclamation, "Campo não informado"

    Case eMsgInformePeloMenosUmCampo
        MsgBox "Informe pelo menos um dos campos em azul", vbExclamation, "Campo não informado"

    Case eMsgNaoExistemInformacoesExcluir
        MsgBox "Não existem informações" & vbCrLf & "para ser excluídas", vbInformation, "Aviso"
    
    Case eMsgRelatorioSemDados
        Screen.MousePointer = vbNormal
        MsgBox "Este relatório não retornou dados.", vbInformation, "Nenhum dado"

    Case eMsgRelatorioIndisponivel
        MsgBox "Este relatório ainda está em desenvolvimento." & vbCrLf & _
               "Por favor, aguarde a liberação.", vbExclamation, "Relatório em construção"

    Case eMsgSelecioneRelatorio
        MsgBox "Selecione relatório para impressão", vbExclamation, "Aviso"

    Case eMsgSemRegistros
        MsgBox "Não existe nenhum registro" & vbCrLf & "que atenda às condições especificadas.", vbInformation, "Nenhum dado"

    Case eMsgTamanhoCampo
        MsgBox vCampo1 & " deve ter exatamente " & vCampo2 & " caracteres.", vbExclamation, "Tamanho incorreto"

    Case eMsgTelaEmConstrucao
        MsgBox "Esta tela ainda está em desenvolvimento." & vbCrLf & _
               "Por favor, aguarde a liberação.", vbInformation, "Tela em construção"

    Case eMsgTelaIndisponivel
        MsgBox "Esta tela/relatório está indisponível temporariamente" & vbCrLf & _
               "devido à manutenção do sistema." & vbCrLf & _
               "Por favor, tente mais tarde", vbExclamation, "Tela indisponível"

    Case eMsgUsuarioLogado
        MsgBox "Usuário já está logado no sistema." & vbCrLf & _
        "O empréstimo de senhas é passível de punição." & vbCrLf & _
        "Favor contactar administrador do sistema.", vbCritical, "Segurança"

    Case eMsgUsuarioNaoAutorizado
        MsgBox "Usuário não autorizado." & vbCrLf & _
        "Favor contactar administrador do sistema.", vbExclamation, "Segurança"

    Case eMsgValorMoedaInvalido
        MsgBox "Valor de moeda inválido.", vbExclamation, "Aviso"
        
    Case eMsgRegistroJaExiste
        If IsMissing(vCampo1) Then
            MsgBox "O registro já existe.", vbInformation, "Aviso"
        Else
            MsgBox "O registro já existe. Para - " & vCampo1, vbInformation, "Aviso"
        End If

    Case eMsgExclusaoIntegridade
        MsgBox "Existe registro associado." & vbCrLf & _
        "Exclusão não autorizada.", vbExclamation, "Aviso"

    Case eMsgRegistroJaExcluido
        MsgBox "O registro já foi excluído.", vbInformation, "Aviso"

    End Select

End Function

Public Function FG_Codigo_Combo(ByVal sConteudo As String, Optional ByVal bAntesdoSinal As Boolean = False, _
    Optional ByVal bRetornaNull As Boolean = True) As String
    '* se bRetornaNull = True e sConteudo = "" então FG_Codigo_Combo = "Null"
    '* sConteudo = refere-se ao conteúdo da Combo
    '* bAntesdoSinal = verifica se o código deve ser pego a partir da direita ou da esquerda do Sinal
    'Ex.: 000-Descrição, então deve ser pego a partir da Esquerda até antes do sinal -
    '     Descrição-000, então deve ser pego a partir da Direita depois do sinal -

    If Trim(sConteudo) = "" Then
        If bRetornaNull Then
            FG_Codigo_Combo = "Null"
        Else
            FG_Codigo_Combo = ""
        End If
        Exit Function
    End If

    If bAntesdoSinal Then
        FG_Codigo_Combo = Trim(Mid(sConteudo, 1, InStr(1, sConteudo, "«»") - 2))
    Else
       FG_Codigo_Combo = Trim(Mid(sConteudo, InStr(1, sConteudo, "«»") + 2))
    End If

End Function

Public Sub FG_Carrega_Tag_Combo(ByRef cmbCombo As ComboBox)
    'Carrega o Código contido na ComboBox para o Tag da mesma
    cmbCombo.Tag = FG_Codigo_Combo(cmbCombo.Text)
    If cmbCombo.Tag = "" Then cmbCombo.Tag = "Null"

End Sub

Public Function FG_Indice_Combo(ByVal cmbCombo As ComboBox, ByVal sItem As String, _
    Optional ByVal yNumCarac As Byte, Optional ByVal bDireita As Boolean = False, _
    Optional ByVal sFormato As String = "") As Integer

    'Função que retorna o índice do item na cmbCombo dado por sItem
    'de acordo com a quantidade de caracteres dada por yNumCarac.
    'Caso yNumCarac não seja informado, compara o sItem com cada item inteiro Combo
    'Se não encontrar, retorna -1

    Dim iLaco   As Integer
    If sItem = "" Then
        FG_Indice_Combo = -1
        Exit Function
    End If
    
    If sFormato <> "" Then
        sItem = Format(sItem, sFormato)
    End If

    For iLaco = 0 To cmbCombo.ListCount - 1
        If bDireita Then
            If yNumCarac > 0 Then
'                If Left(cmbCombo.List(iLaco), 4) = "PRO-" Then MsgBox "OU"
                If UCase(sItem) = Right(UCase(cmbCombo.List(iLaco)), Len(sItem)) Then
                    FG_Indice_Combo = iLaco
                    Exit Function
                End If
            Else
                If UCase(sItem) = UCase(cmbCombo.List(iLaco)) Then
                    FG_Indice_Combo = iLaco
                    Exit Function
                End If
            End If
        Else
            If yNumCarac > 0 Then
                If UCase(sItem) = Left(UCase(cmbCombo.List(iLaco)), Len(sItem)) Then
                    FG_Indice_Combo = iLaco
                    Exit Function
                End If
            Else
                If UCase(sItem) = UCase(cmbCombo.List(iLaco)) Then
                    FG_Indice_Combo = iLaco
                    Exit Function
                End If
            End If
        End If
    Next

    FG_Indice_Combo = -1

End Function

Public Sub PG_Grava_CacheIni(NomeDoINI As String)
    Dim Ret As Integer
    Ret = GravaEntradaPrivIni(0&, 0&, 0&, NomeDoINI)
End Sub

Function FG_LeEntradaIni(ByVal sSecao As String, ByVal sEntrada As String, _
    ByVal sValorDefault As String, ByVal sNomeDoINI As String) As String
    Dim sTemp As String * 255
    Dim lRet As Long
    lRet = apiLeEntradaPrivIni(sSecao, sEntrada, sValorDefault, _
         sTemp, Len(sTemp), sNomeDoINI)
    If lRet Then
       FG_LeEntradaIni = Left$(sTemp, lRet)
    End If

End Function

Public Function FG_Data_Valida(lData As String, Optional bMsg As Boolean = False) As Boolean

  '************************************************'
  'Funcao para validar data no formato DD/MM/AAAA  '
  '************************************************'

  Dim sAnoData, sMesData As Integer

  If FG_Conv_Data(lData) = "" Then
     FG_Data_Valida = False
     GoTo Sair
  End If

  sMesData = Val(Mid(lData, 4, 2))
  sAnoData = Val(Mid(lData, 7, 4))

  If (sAnoData > 1000 Or sAnoData < 100) And _
     (sMesData > 0 And sMesData < 13) Then
      FG_Data_Valida = IsDate(lData)
  Else
      FG_Data_Valida = False
  End If
  
Sair:
  If FG_Data_Valida = False Then
      If bMsg Then FG_MsgBoxPadrao (eMsgDataInvalida)
  End If

End Function

Public Function FG_Conv_Data(ByVal sData As String, Optional ByVal sFormato As String = "yyyy/mm/dd") As String
    'Esta Função é para ser uzada de preferência em Querys

    If sData = "" Or sData = "  /  /    " Or sData = "/  /" Then
        FG_Conv_Data = "Null"
    Else
        FG_Conv_Data = "'" & Format(sData, sFormato) & "'"
    End If

End Function

Public Sub PG_Limpa_Combo(ByVal iCodTecla As Integer, ByRef cmbCombo As ComboBox)
    If iCodTecla = vbKeyDelete Or iCodTecla = vbKeyBack Then
        cmbCombo.ListIndex = -1
    End If

End Sub

Public Function FG_Semaforo(ByVal sConexao As String, ByVal sTabela As String, ByVal sCampo As String, _
Optional ByVal sCampo1 As String, Optional ByVal sValCampo1 As String) As Long
    Dim rsSemaforo As ADODB.Recordset
    Dim sQry        As String
    Dim dbBanco  As ADODB.Connection

    Set dbBanco = CreateObject("ADODB.Connection")
    dbBanco.ConnectionString = sConexao

    dbBanco.ConnectionTimeout = 60
    dbBanco.CommandTimeout = 400
    dbBanco.CursorLocation = adUseClient
    dbBanco.Open

    ' Testa se Existe a Tabela no Semáforo
    sQry = "Select Convert(" & sCampo & ", SIGNED) As Codigo From " & sTabela & " "
    If sCampo1 <> "" And sValCampo1 <> "" Then
        sQry = sQry & "Where " & sCampo1 & " = " & sValCampo1 & " "
    End If
    sQry = sQry & "Order by Codigo Desc"

    Set rsSemaforo = dbBanco.Execute(sQry)
    If rsSemaforo.EOF Then
        FG_Semaforo = 1
    Else
        FG_Semaforo = CCur(rsSemaforo!Codigo) + 1
    End If
    dbBanco.Close

End Function

Public Sub PG_Configura_Formulario(oForm As Object)
    oForm.Top = 0
    oForm.Left = 0
    oForm.Height = mdiPrincipal.ScaleHeight
    oForm.Width = mdiPrincipal.ScaleWidth

End Sub

Public Function FG_Data_Servidor(Optional ByVal sFormato As String = "dd/mm/yyyy") As String
    Dim sQry    As String
    Dim rsData  As ADODB.Recordset

    sQry = "Select Convert(Now(), Date) As DataServidor"
    Set rsData = dbBancoDados.Execute(sQry)
    If Not rsData.EOF Then
        FG_Data_Servidor = Format(rsData!DataServidor, sFormato)
    Else
        MsgBox "Erro na Data do Servidor.", vbInformation, "Aviso"
        FG_Data_Servidor = Format(Date, sFormato)
    End If
    rsData.Close

End Function

Public Function FG_Hora_Servidor(Optional ByVal sFormato As String = "hh:mm:ss") As String
    Dim sQry    As String
    Dim rsHora  As ADODB.Recordset

    sQry = "Select Convert(Now(), Time) As HoraServidor"
    Set rsHora = dbBancoDados.Execute(sQry)
    If Not rsHora.EOF Then
        FG_Hora_Servidor = Format(rsHora!HoraServidor, sFormato)
    Else
        MsgBox "Erro na Data do Servidor.", vbInformation, "Aviso"
        FG_Hora_Servidor = Format(Time, sFormato)
    End If
    rsHora.Close

End Function

Sub Main()
    Screen.MousePointer = vbHourglass

    'verifica se a aplicação já está aberta
    If App.PrevInstance Then
        MsgBox "A Aplicação Já Se Encontra Aberta!", vbInformation, App.EXEName
        End
    End If

    'LER ARQUIVO INI
    G_Servidor = FG_LeEntradaIni("CONEXAO", "NOME_SERVIDOR", "", App.Path & "\cnnDesafio.ini")
    G_NomeBanco = FG_LeEntradaIni("CONEXAO", "NOME_BANCO", "", App.Path & "\cnnDesafio.ini")

    Screen.MousePointer = vbDefault
    mdiPrincipal.Show

End Sub

Public Function FG_Criptografa(Palavra As String, Operacao As String) As String
    
    'esta funcao serve tanto para criptografar quanto para descriptografar
    
    Dim A As String
    Dim B As String
    Dim Pos As Integer
    Dim i As Integer
    Dim Retorno As String
    Dim chave As String
    
    chave = GC_Chave
    'Testa se foi passado uma chave para criptografia, troca de caracteres
'    If Format(chave) = "" Then
        'O esquema de chave é simples, basta passar numeros no intervalo de 1 a 5
        'qualquer numero maior que 5 sera convertido automaticamente para 5
        'Quando a operacao é C criptografar, sera montada uma sequencia com esta chave
        ' Ex.: A palavra "INFORMATICA", será montada a seguinte string com o mesmo tamanho
        '      string =  "32132132132"
        '
        '     O objetivo disto e subtrair o codigo Ascii pela sequencia de cada numero da
        '     string
        'Ex.: O Ascii da Letra I sera subtraido por 3, da N por 2, da F por 1 e assim por diante
        
        'Quando a operacao for D, o processo e o mesmo so que somando o ascci pelo numero
        'da string montada
        'Quando a operacao é D descriptografia
'        chave = "312123"
'    End If
    
    If Operacao = "C" Then
        'Inverte a palavra com a chave a palavra passada primeiro se for criptografia
        Palavra = FG_Troca_Caracteres(Palavra, chave, Operacao)
    End If
    
    'Uma variável é o inverso da outra, que tem todos os caracters visíveis de ascii 32 as 255
    A = " !" & Chr(34) & "#$%&'()*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\]^_`abcdefghijklmnopqrstuvwxyz{|}~€‚ƒ„…†‡ˆ‰Š‹ŒŽ‘’“”•–—˜™š›œžŸ ¡¢£¤¥¦§¨©ª«¬­®¯°±²³´µ¶·¸¹º»¼½¾¿ÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖ×ØÙÚÛÜÝÞßàáâãäåæçèéêëìíîïðñòóôõö÷øùúûüýþÿ"
    B = "ÿþýüûúùø÷öõôóòñðïîíìëêéèçæåäãâáàßÞÝÜÛÚÙØ×ÖÕÔÓÒÑÐÏÎÍÌËÊÉÈÇÆÅÄÃÂÁÀ¿¾½¼»º¹¸·¶µ´³²±°¯®­¬«ª©¨§¦¥¤£¢¡ Ÿžœ›š™˜—–•”“’‘ŽŒ‹Š‰ˆ‡†…„ƒ‚€~}|{zyxwvutsrqponmlkjihgfedcba`_^]\[ZYXWVUTSRQPONMLKJIHGFEDCBA@?>=<;:9876543210/.-,+*)('&%$#" & Chr(34) & "! "
    For i = 1 To Len(Trim(Palavra))
        'Acha A posicao na string A
        Pos = InStr(A, Mid(Palavra, i, 1))
        If Pos > 0 Then
           'Com a mesma posicao em A troca-se pela equivalente em B
           Retorno = Retorno & Mid(B, Pos, 1)
        End If
    Next
    
    If Operacao = "D" Then
        'Inverte a palavra depois , se for descriptografia
        Retorno = FG_Troca_Caracteres(Retorno, chave, Operacao)
    End If
    
    FG_Criptografa = Retorno
    
End Function

Public Function FG_Troca_Caracteres(Palavra As String, chave As String, Operacao As String) As String

    Dim i As Integer
    Dim Cont As Integer
    'Variavel para controle de virada da chave
    Dim Limite As Integer
    'variavel para o indice
    Dim Indice As Integer
    'Variavel para guardar a letra individualamente
    Dim Letra As String
    Dim LetraTrocada As String
    'Variavel com a palavra trocada
    Dim PalavraRetorno As String
    
    'Conta quantas letra tem na chave
    Limite = Len(chave)
    
    'Iicializa o contador com 1
    Cont = 1
    
    'Inicializa a palavra de retorno
    PalavraRetorno = ""
    
    
    'Varre a palavra para substituir pelo ascii correspondente pela chave
    For i = 1 To Len(Palavra)
        
        'Controle de que posicao da chave esta sendo usado, o primeiro segundo, terceiro ...
        'para cada palavra da string
        'Se a chave for menor que a apalavra, retornar ao comeco Dela
        If Cont > Limite Then Cont = 1
        
        'extrai a letra
        Letra = Mid(Palavra, i, 1)
        'extrai o indice de troca de acordo com a posicao da chave
        Indice = Mid(chave, Cont, 1)
        
        'troca a letra
        If Operacao = "C" Then
            'Se é criptografar incremente o ascii com o valor do indice da chave
            LetraTrocada = Chr(Asc(Letra) + Indice)
        Else
            'Faz o contrario se é para Descriptografar
            LetraTrocada = Chr(Asc(Letra) - Indice)
        End If
        
        'monta astring de retorno com a troca a letra
        PalavraRetorno = PalavraRetorno & LetraTrocada
        
        'Incrementa a chave de troca de caracteres
        Cont = Cont + 1
    Next
    FG_Troca_Caracteres = PalavraRetorno

End Function

Public Function FG_Troca_Caracter(sFrase As String, sCarac1 As String, Optional sCarac2 As String) As String
    'Função que troca caractere dado por sCarac1 por sCarac2
    'Caso sCarac2 seja vazio, o caractere sCarac1 será na realidade excluído
    
    Dim iPosicao As Integer
    
    sFrase = Trim(sFrase)
    iPosicao = 1
    Do While iPosicao > 0
        iPosicao = InStr(iPosicao, sFrase, sCarac1)
        If iPosicao > 0 Then
            sFrase = Left(sFrase, iPosicao - 1) & sCarac2 & Right(sFrase, Len(sFrase) - iPosicao)
            iPosicao = iPosicao - 1
        End If
    Loop
    
    FG_Troca_Caracter = sFrase

End Function

Public Function FG_Iniciais_Maiusculas(sNome As String) As String
    Dim i, j    As Integer
    Dim vPre    As Variant
    Dim bAchou  As Boolean
    Dim sPalavra As String

    vPre = Array("de ", "do ", "da ", "dos ", "das ", "e ", "em ")

    sNome = StrConv(Trim(sNome), vbProperCase)
    For i = 1 To Len(Trim(sNome))
        sPalavra = sPalavra & Mid(sNome, i, 1)
        If Mid(sNome, i, 1) = " " Then
            For j = 0 To 6
                If vPre(j) = LCase(sPalavra) Then
                    sPalavra = vPre(j)
                    FG_Iniciais_Maiusculas = FG_Iniciais_Maiusculas & sPalavra
                    bAchou = True
                    Exit For
                End If
            Next
            If Not bAchou Then
                FG_Iniciais_Maiusculas = FG_Iniciais_Maiusculas & sPalavra
            End If
            sPalavra = ""
            bAchou = False
        End If
    Next
    FG_Iniciais_Maiusculas = FG_Iniciais_Maiusculas & sPalavra

End Function

Public Function FG_Conv_Ponto(ByVal Valor As String, Optional NumeroDecimais As Integer = 2, Optional ByVal bTrunca As Boolean = True, Optional ByVal bArredonda As Boolean = False) As String

        'Esta funcao transforma um numero formato windows com virgula em numero formato SqlServer
        'com ponto decimal, e colocar casas decimais com a direita preenchida com zero

        ' Valor  = parametro usado para receber o numero que tem que ser string
        'Decimal = parametro usado para a precisao de decimais usadas, se o
        '          numero contiver um numero de decimais maior que o parametro, elas serao
        '          truncadas, caso sejam menores seram completados com zeros a direita
        'Se bArredonda=true, então haverá arredondamento matemático. Caso contrário
        'haverá truncamento simples

       Dim ValorRetorno As String
       Dim ValorTemp As String
       Dim i As Integer
       Dim Decimais As String
       Dim PosicaoVirgula As Integer
       Dim Temp As String

       ValorTemp = CCur(Valor)
       ValorRetorno = ""

       If Trim(Valor) = "" Then
          ValorRetorno = "0"
       End If

       i = InStr(1, ValorTemp, ".")
       If i > 0 Then
          ValorTemp = Left(ValorTemp, i - 1) & "," & Right(ValorTemp, Len(ValorTemp) - i)
       Else
          i = InStr(1, ValorTemp, ",")
          If i = 0 Then
             ValorTemp = ValorTemp & ","
          End If
       End If

       If bArredonda Then
          ValorTemp = FG_ArredondarValor(ValorTemp, NumeroDecimais)
       ElseIf bTrunca Then
          ValorTemp = FG_TruncarValor(ValorTemp, NumeroDecimais)
       End If

       'Trunca as decimais

       'Troca virgulas por ponto
       For i = 1 To Len(ValorTemp)
            'Testa se ha algum espaco em branco, caso sim , nao deixar passar
            If Mid(ValorTemp, i, 1) <> " " Then
                'Troca a virgula por ponto
                If Mid(ValorTemp, i, 1) = "," Then
                    ValorRetorno = ValorRetorno & "."
                Else
                    ValorRetorno = ValorRetorno & Mid(ValorTemp, i, 1)
                End If
            End If
       Next

       FG_Conv_Ponto = ValorRetorno

End Function

Public Function FG_ArredondarValor(Valor As String, Optional Decimais As Integer) As String
        
    Dim strInteiros As String, strDecimais As String, strAux As String
    Dim intI As Integer
    Dim strDecimaisOrig As String

    If InStr(1, Trim(Valor), ",", 1) > 0 Then
        strInteiros = Left(Valor, InStr(1, Valor, ",") - 1)
        If Len(Decimais) = 0 Then
            strDecimais = Mid(Valor, Len(strInteiros) + 2, 2)
            Valor = Format(strInteiros & "," & strDecimais, "########0.00")
        Else
            strDecimais = Right(Valor, Len(Valor) - InStr(1, Valor, ","))
            strAux = "########0."
            For intI = 1 To Val(Decimais)
                If intI > 8 Then Exit For
                strAux = strAux & "0"
            Next intI
            strDecimaisOrig = strDecimais
            strDecimais = Round(strDecimais / 10 ^ Len(strDecimaisOrig), Decimais)
            If strDecimais = 1 Then
               strDecimais = "0"
               strInteiros = Val(strInteiros) + 1
            Else
               strDecimais = Right(strDecimais, Len(strDecimais) - InStr(1, strDecimais, ","))
            End If
            'strDecimais = Int(strDecimais / (10 ^ (Decimais - 1)))
            Valor = Format(strInteiros & "," & strDecimais, strAux)
        End If
    Else
        Valor = Format(Valor, "########0.00")
    End If
    FG_ArredondarValor = Valor

End Function

Public Function FG_TruncarValor(Valor As String, Optional Decimais As Integer) As String
    
    Dim strInteiros As String, strDecimais As String, strAux As String
    Dim intI As Integer

    If InStr(1, Trim(Valor), ",", 1) > 0 Then
        strInteiros = Left(Valor, InStr(1, Valor, ",") - 1)
        If Len(Decimais) = 0 Then
            strDecimais = Mid(Valor, Len(strInteiros) + 2, 2)
            Valor = Format(strInteiros & "," & strDecimais, "########0.00")
        Else
            strDecimais = Mid(Valor, Len(strInteiros) + 2, Val(Decimais))
            strAux = "########0."
            For intI = 1 To Val(Decimais)
                If intI > 8 Then Exit For
                strAux = strAux & "0"
            Next intI
            Valor = Format(strInteiros & "," & strDecimais, strAux)
        End If
    Else
        Valor = Format(Valor, "########0.00")
    End If
    FG_TruncarValor = Valor

End Function

Public Function FG_Formata_CNPJCPF(ByVal sCNPJCPF As String) As String
    If sCNPJCPF = "" Then Exit Function
    If Len(sCNPJCPF) = 11 Then
        FG_Formata_CNPJCPF = Format(sCNPJCPF, GC_Formato_CPF)
    ElseIf Len(sCNPJCPF) = 11 Then
        FG_Formata_CNPJCPF = Format(sCNPJCPF, GC_Formato_CGC)
    End If

End Function

Public Function FG_Retorna_Indice_Butao(ByVal tlbBotao As Toolbar, ByVal sKeyBotao As String) As Integer
    Dim iBotao As Integer
    
    FG_Retorna_Indice_Butao = 1
    For iBotao = 1 To tlbBotao.Buttons.Count
        If tlbBotao.Buttons(iBotao).Key = sKeyBotao Then
            FG_Retorna_Indice_Butao = iBotao
            Exit For
        End If
    Next

End Function

Public Function FG_AnoMes(ByVal sAnoMes As String) As String
    FG_AnoMes = Format(sAnoMes, "yyyymm")

End Function

Public Function FG_AnoMes_Anterior(ByVal sMesAno As String) As String
    sMesAno = "01/" & sMesAno
    sMesAno = DateAdd("m", -1, sMesAno)
    FG_AnoMes_Anterior = Format(sMesAno, "yyyymm")

End Function

Public Function FG_Usuario_Windows() As String
    'Retorna o nome do usuário logado no Windows NT, 2000, XP...
    Dim sNome       As String
    Dim lTamanho    As Long
    
    sNome = Space(200)
    lTamanho = Len(sNome)
    Call GetUserName(sNome, lTamanho)
    sNome = Left$(sNome, lTamanho)
    FG_Usuario_Windows = UCase$(sNome)

End Function

Public Function FG_Existe_Arquivo(Arquivo As String) As Boolean
    'testa a existencia de um arquivo
    On Error Resume Next

    Open Arquivo For Input As #1

    FG_Existe_Arquivo = IIf(Err = 0, True, False)

    Close #1

    Err = 0
    
    On Error GoTo 0
    Err.Clear

End Function

Public Sub PG_Seleciona_Campo(ByVal oCampo As Object)
    oCampo.SelStart = 0
    oCampo.SelLength = Len(oCampo.Text)

End Sub

Public Sub GravarErroLog(ByVal MensagemErro As String)
    On Error Resume Next
    
    Dim CaminhoLog As String
    Dim NomeArquivo As String
    Dim Arquivo As Integer
    Dim DataAtual As String
    Dim HoraAtual As String
    Dim LinhaLog As String

    ' Monta nome do arquivo no formato AAAA-MM-DD.log
    DataAtual = Format(Date, "yyyy-mm-dd")
    NomeArquivo = DataAtual & ".log"
    
    ' Define o caminho da pasta onde o log será salvo
    CaminhoLog = App.Path & "\Logs\"
    
    ' Cria a pasta se não existir
    If Dir(CaminhoLog, vbDirectory) = "" Then
        MkDir CaminhoLog
    End If
    
    ' Caminho completo do arquivo de log
    CaminhoLog = CaminhoLog & NomeArquivo
    
    ' Monta linha de log com data e hora
    HoraAtual = Format(Time, "hh:nn:ss")
    LinhaLog = "[" & DataAtual & " " & HoraAtual & "] " & MensagemErro
    
    ' Abre o arquivo para adicionar a linha no final
    Arquivo = FreeFile
    Open CaminhoLog For Append As #Arquivo
        Print #Arquivo, LinhaLog
    Close #Arquivo
End Sub

