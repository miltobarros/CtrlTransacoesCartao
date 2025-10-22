Attribute VB_Name = "Modulo_Particular"
Option Explicit

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

