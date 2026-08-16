VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOcorrencias 
   Caption         =   "OCORRENCIAS / SUGESTOES"
   ClientHeight    =   8415.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13620
   OleObjectBlob   =   "frmOcorrencias.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmOcorrencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnReport_Click()

    frmReport.Show vbModeless
    
End Sub

Private Sub UserForm_Initialize()

'On Error GoTo TrataErro

    With cmbStatus
        .Clear
        .AddItem "TODOS"
        .AddItem "PENDENTE"
        .AddItem "CONCLUIDO"
        .ListIndex = 0
    End With

   With lvOcorrencias
    
    .View = lvwReport
    .AllowColumnReorder = True
    .FullRowSelect = True
    .HideSelection = False
    .Gridlines = True
    
    .ColumnHeaders.Clear

    .ColumnHeaders.Add , , "ID", 40
    .ColumnHeaders.Add , , "DATA", 60
    .ColumnHeaders.Add , , "OCORRÊNCIA", 140
    .ColumnHeaders.Add , , "TELA", 90
    .ColumnHeaders.Add , , "DESCRIÇÃO", 200
    .ColumnHeaders.Add , , "STATUS", 80
    .ColumnHeaders.Add , , "USÚARIO", 70
    .ColumnHeaders.Add , , "VERSÃO", 60

End With

    CarregarOcorrencias
    CarregarDatas
    
Exit Sub

TrataErro:

    MsgBox "Erro " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical, "Erro"

End Sub

Private Sub CarregarOcorrencias()

    On Error GoTo TrataErro

    Dim rs As ADODB.Recordset
    Dim SQL As String
    Dim Etapa As String
    Dim ITEM As ListItem

Etapa = "Montando SQL"

    SQL = "SELECT " & _
          "O.IDREPORT, " & _
          "O.DATA, " & _
          "O.OCORRENCIA, " & _
          "O.TELA, " & _
          "O.DESCRICAO, " & _
          "O.STATUS, " & _
          "U.USUARIO, " & _
          "O.VERSAO_SISTEMA " & _
          "FROM TAB_OCORRENCIAS O " & _
          "LEFT JOIN TAB_USUARIOS U ON U.IDUSUARIO = O.IDUSUARIO " & _
          "WHERE 1=1 "

    If cmbDataInicial.Value <> "" Then

        If cmbDataInicial.ListIndex > 0 Then

            SQL = SQL & _
                " AND O.DATA >= '" & _
                Format$(NzDate(cmbDataInicial.Value), "yyyy-mm-dd") & "'"

        End If

    End If


    If cmbDataFinal.Value <> "" Then

        If cmbDataFinal.ListIndex > 0 Then

            SQL = SQL & _
                " AND O.DATA <= '" & _
                Format$(NzDate(cmbDataFinal.Value), "yyyy-mm-dd") & "'"

        End If

    End If

    If cmbStatus.Value <> "" And _
       cmbStatus.Value <> "TODOS" Then

        SQL = SQL & _
            " AND O.STATUS = " & _
            SqlTexto(cmbStatus.Value)

    End If

    SQL = SQL & _
          " ORDER BY O.IDREPORT DESC"
          
Etapa = "Executando SQL"
Set rs = Conn.Execute(SQL)

Etapa = "Limpando ListView"
lvOcorrencias.ListItems.Clear

        lvOcorrencias.ListItems.Clear

Etapa = "Carregando ListView"

        Do While Not rs.EOF
        
            Set ITEM = lvOcorrencias.ListItems.Add(, , Nz(rs!IDREPORT))
        
            ITEM.ListSubItems.Add , , Format(rs!Data, "dd/mm/yyyy")
            ITEM.ListSubItems.Add , , Nz(rs!OCORRENCIA)
            ITEM.ListSubItems.Add , , Nz(rs!tela)
            ITEM.ListSubItems.Add , , Nz(rs!DESCRICAO)
            ITEM.ListSubItems.Add , , Nz(rs!STATUS)
            ITEM.ListSubItems.Add , , Nz(rs!Usuario)
            ITEM.ListSubItems.Add , , Nz(rs!VERSAO_SISTEMA)
        
            rs.MoveNext
        
        Loop


    '========================================
    ' TOTAL
    '========================================

    'txtTotal.Value = lvOcorrencias.ListItems.Count


    '========================================
    ' FECHA
    '========================================

    rs.Close
    Set rs = Nothing

    Exit Sub


TrataErro:

    MsgBox "Etapa: " & Etapa & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical, "frmOcorrencias"

End Sub

Private Sub cmbDataInicial_Change()

'On Error GoTo TrataErro

    FormatarDatacmb cmbDataInicial

Exit Sub

TrataErro:

    MsgBox "Erro " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical, "Erro"

End Sub


Private Sub cmbDataFinal_Change()

'On Error GoTo TrataErro

    FormatarDatacmb cmbDataFinal

Exit Sub

TrataErro:

    MsgBox "Erro " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical, "Erro"

End Sub


Private Sub cmbStatus_Change()

'On Error GoTo TrataErro

    CarregarOcorrencias

Exit Sub

TrataErro:

    MsgBox "Erro " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical, "Erro"

End Sub

Private Sub lvOcorrencias_DblClick()

    Dim ID As Long

    'If lvOcorrencias.SelectedItems = "" Then Exit Sub

    ID = CLng(lvOcorrencias.SelectedItem)

    frmReport.IDOcorrencia = ID

    frmReport.Show vbModeless

End Sub

Private Sub CarregarDatas()

    On Error GoTo TrataErro

    Dim rs As ADODB.Recordset
    Dim SQL As String
    Dim DataInicial As Date
    Dim DataFinal As Date
    Dim DataAtual As Date

    SQL = "SELECT MIN(DATA) AS DATAINICIAL, " & _
          "MAX(DATA) AS DATAFINAL " & _
          "FROM TAB_OCORRENCIAS " & _
          "WHERE DATA IS NOT NULL"

    Set rs = Conn.Execute(SQL)

    cmbDataInicial.Clear
    cmbDataFinal.Clear

    If Not IsNull(rs!DataInicial) Then

        DataInicial = CDate(rs!DataInicial)
        DataFinal = CDate(rs!DataFinal)

        DataAtual = DataInicial

        Do While DataAtual <= DataFinal

            cmbDataInicial.AddItem Format(DataAtual, "dd/mm/yyyy")
            cmbDataFinal.AddItem Format(DataAtual, "dd/mm/yyyy")

            DataAtual = DateAdd("d", 1, DataAtual)

        Loop

    End If

    rs.Close
    Set rs = Nothing

    Exit Sub

TrataErro:

    MsgBox "Erro " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical, "Erro ao carregar datas"

End Sub



