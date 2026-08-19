VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFluxoCaixa 
   Caption         =   "FLUXO DE CAIXA"
   ClientHeight    =   9690.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18765
   OleObjectBlob   =   "frmFluxoCaixa.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmFluxoCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub lvFluxoCaixa_DblClick()

On Error GoTo TratarErro

Dim dados() As String
Dim ano As Long
Dim mes As Long
Dim ITEM As ListItem

Set ITEM = lvFluxoCaixa.SelectedItem

If ITEM Is Nothing Then Exit Sub

dados = Split(ITEM.Tag, "|")

ano = CLng(dados(0))
mes = CLng(dados(1))

    frmVendaDiaria.ano = ano
    frmVendaDiaria.mes = mes
    frmVendaDiaria.Show vbModal

    Exit Sub

TratarErro:

    modSistema.tela = "FluxoCaixa - LVFluxo DBLCLICK"
    modSistema.DescErro = Err.Description
    modSistema.nErro = Err.Number

    Call modSistema.ReportarErro
    
    MsgBox "Erro: " & Err.Number & vbCrLf & _
                        Err.Description, vbInformation, "SISTEMA"

End Sub

Private Sub lvPagar_DblClick()

On Error GoTo TratarErro

    If lvPagar.SelectedItem Is Nothing Then Exit Sub

    With frmContas
        .ModoNovo = False
        .TipoConta = "PAGAR"
        .IdConta = CLng(lvPagar.SelectedItem.Tag)
        .CarregarConta
        .Show
    End With

    Exit Sub

TratarErro:

    modSistema.tela = "FluxoCaixa - lvpagar dblclick"
    modSistema.DescErro = Err.Description
    modSistema.nErro = Err.Number

    Call modSistema.ReportarErro
    
    MsgBox "Erro: " & Err.Number & vbCrLf & _
                        Err.Description, vbInformation, "SISTEMA"

End Sub

Private Sub lvReceber_DblClick()

On Error GoTo TratarErro

    If lvReceber.SelectedItem Is Nothing Then Exit Sub

    With frmContas
        .ModoNovo = False
        .TipoConta = "RECEBER"
        .IdConta = CLng(lvReceber.SelectedItem.Tag)
        .CarregarConta
        .Show
    End With

    Exit Sub

TratarErro:

    modSistema.tela = "FluxoCaixa - lvreceber dblclick"
    modSistema.DescErro = Err.Description
    modSistema.nErro = Err.Number

    Call modSistema.ReportarErro
    
    MsgBox "Erro: " & Err.Number & vbCrLf & _
                        Err.Description, vbInformation, "SISTEMA"

End Sub

Private Sub btnNovaConta_Click()

    With frmContas
        .ModoNovo = True
        .TipoConta = "PAGAR"
        .IdConta = 0
        .PrepararNovaConta
        .Show
    End With

End Sub

Private Sub userform_activate()

    CarregarFluxoCaixa

End Sub

Private Sub UserForm_Initialize()

On Error GoTo TratarErro

    With lvFluxoCaixa

        .View = lvwReport
        .Gridlines = True
        .AllowColumnReorder = True
        .HideSelection = False
        .FullRowSelect = True
        
        .ColumnHeaders.Clear
        .ListItems.Clear

        .ColumnHeaders.Add , , "ANO", 50
        .ColumnHeaders.Add , , "MÊS", 70
        .ColumnHeaders.Add , , "ENTRADAS", 80
        .ColumnHeaders.Add , , "SAÍDAS", 80
        .ColumnHeaders.Add , , "SALDO", 80
        .ColumnHeaders.Add , , "N° PEDIDOS", 80
        .ColumnHeaders.Add , , "TICKET MÉDIO", 100
        .ColumnHeaders.Add , , "ACUMULADO", 100

    End With

    CarregarFluxoCaixa
     PreencherLvPagar
    PreencherLvReceber

    Exit Sub

TratarErro:

    modSistema.tela = "FluxoCaixa - initialize"
    modSistema.DescErro = Err.Description
    modSistema.nErro = Err.Number

    Call modSistema.ReportarErro
    
    MsgBox "Erro: " & Err.Number & vbCrLf & _
                        Err.Description, vbInformation, "SISTEMA"

End Sub
Private Sub CarregarFluxoCaixa()

On Error GoTo TratarErro

    Dim rs As New ADODB.Recordset
    Dim ITEM As ListItem

    lvFluxoCaixa.ListItems.Clear

    rs.Open SQL_FLUXO_CAIXA, Conn, adOpenForwardOnly, adLockReadOnly

    Do Until rs.EOF

        Set ITEM = lvFluxoCaixa.ListItems.Add(, , rs!ano)

        ITEM.ListSubItems.Add , , UCase(Format(DateSerial(rs!ano, rs!mes, 1), "mmmm"))
        ITEM.ListSubItems.Add , , FormatCurrency(rs!ENTRADA)
        ITEM.ListSubItems.Add , , FormatCurrency(rs!SAIDA)
        ITEM.ListSubItems.Add , , FormatCurrency(rs!SALDO)
        ITEM.ListSubItems.Add , , rs!PEDIDOS
        ITEM.ListSubItems.Add , , FormatCurrency(rs!TM)
        ITEM.ListSubItems.Add , , FormatCurrency(rs!ACUMULADO)

        txtEntrada.Value = FormatCurrency(rs!ENTRADA)
        txtSaida.Value = FormatCurrency(rs!SAIDA)
        txtSaldo.Value = FormatCurrency(rs!SALDO)
        txtAcumulado.Value = FormatCurrency(rs!ACUMULADO)
        
        ITEM.Tag = rs!ano & "|" & rs!mes
        
        rs.MoveNext

    Loop

    rs.Close
    Set rs = Nothing

    Exit Sub

TratarErro:

    modSistema.tela = "FluxoCaixa - carregar lv fluxo"
    modSistema.DescErro = Err.Description
    modSistema.nErro = Err.Number

    Call modSistema.ReportarErro
    
    MsgBox "Erro: " & Err.Number & vbCrLf & _
                        Err.Description, vbInformation, "SISTEMA"

End Sub

Private Function SQL_FLUXO_CAIXA() As String

    Dim sql As String

    sql = "SELECT M.ANO,M.MES,"
    sql = sql & "COALESCE(C.SAIDA,0) SAIDA,"
    sql = sql & "COALESCE(V.ENTRADA,0) ENTRADA,"
    sql = sql & "COALESCE(V.ENTRADA,0)-COALESCE(C.SAIDA,0) SALDO,"
    sql = sql & "COALESCE(P.PEDIDOS,0) PEDIDOS,"
    sql = sql & "ROUND(COALESCE(V.ENTRADA,0)/NULLIF(P.PEDIDOS,0),2) TM,"
    sql = sql & "SUM(COALESCE(V.ENTRADA,0)-COALESCE(C.SAIDA,0)) OVER(ORDER BY M.ANO,M.MES) ACUMULADO "

    sql = sql & "FROM("
    sql = sql & "SELECT YEAR(DATAVENDA) ANO,MONTH(DATAVENDA) MES FROM TAB_VENDAS GROUP BY 1,2 "
    sql = sql & "UNION "
    sql = sql & "SELECT YEAR(DATACOMPRA),MONTH(DATACOMPRA) FROM TAB_COMPRAS GROUP BY 1,2"
    sql = sql & ")M "

    sql = sql & "LEFT JOIN("
    sql = sql & "SELECT YEAR(DATAPAGAMENTO) ANO,MONTH(DATAPAGAMENTO) MES,SUM(C.VALOR) SAIDA "
    sql = sql & "FROM TAB_CONTASAPAGAR C GROUP BY 1,2"
    sql = sql & ")C ON C.ANO=M.ANO AND C.MES=M.MES "

    sql = sql & "LEFT JOIN("
    sql = sql & "SELECT YEAR(DATAVENDA) ANO,MONTH(DATAVENDA) MES,SUM(VALORTOTAL) ENTRADA "
    sql = sql & "FROM TAB_VENDAS GROUP BY 1,2"
    sql = sql & ")V ON V.ANO=M.ANO AND V.MES=M.MES "

    sql = sql & "LEFT JOIN("
    sql = sql & "SELECT YEAR(DATAVENDA) ANO,MONTH(DATAVENDA) MES,COUNT(*) PEDIDOS "
    sql = sql & "FROM TAB_VENDAS GROUP BY 1,2"
    sql = sql & ")P ON P.ANO=M.ANO AND P.MES=M.MES "

    sql = sql & "ORDER BY 1,2"
    
    SQL_FLUXO_CAIXA = sql
    
End Function

Sub PreencherLvPagar()

On Error GoTo TratarErro

    Dim rs As ADODB.Recordset
    Dim sql As String
    Dim ITEM As ListItem

    sql = "SELECT IDAPAGAR, VENCIMENTO, VALOR, STATUS " & _
          "FROM tab_contasapagar " & _
          "WHERE IFNULL(STATUS,'') <> 'CANCELADO' " & _
          "ORDER BY VENCIMENTO DESC"

    Set rs = New ADODB.Recordset
    rs.Open sql, Conn, adOpenForwardOnly, adLockReadOnly

    With lvPagar
        .ListItems.Clear
        .ColumnHeaders.Clear
        .View = lvwReport
        .Gridlines = True
        .FullRowSelect = True

        .ColumnHeaders.Add , , "VENCIMENTO", 90
        .ColumnHeaders.Add , , "VALOR", 80
        .ColumnHeaders.Add , , "STATUS", 90
    End With

    Do While Not rs.EOF

        Set ITEM = lvPagar.ListItems.Add(, , Format(rs!VENCIMENTO, "dd/mm/yyyy"))
        ITEM.Tag = rs!IDAPAGAR

        ITEM.ListSubItems.Add , , Format(NzDbl(rs!valor), "#,##0.00")
        ITEM.ListSubItems.Add , , Nz(rs!STATUS)

        rs.MoveNext

    Loop

    rs.Close
    Set rs = Nothing

    Exit Sub

TratarErro:

    modSistema.tela = "FluxoCaixa - carregar lv pagar"
    modSistema.DescErro = Err.Description
    modSistema.nErro = Err.Number

    Call modSistema.ReportarErro
    
    MsgBox "Erro: " & Err.Number & vbCrLf & _
                        Err.Description, vbInformation, "SISTEMA"

End Sub

Sub PreencherLvReceber()

On Error GoTo TratarErro

    Dim rs As ADODB.Recordset
    Dim sql As String
    Dim ITEM As ListItem

    sql = "SELECT IDARECEBER, VENCIMENTO, VALOR, STATUS " & _
          "FROM tab_contasareceber " & _
          "WHERE IFNULL(STATUS,'') <> 'CANCELADO' " & _
          "ORDER BY VENCIMENTO DESC"

    Set rs = New ADODB.Recordset
    rs.Open sql, Conn, adOpenForwardOnly, adLockReadOnly

    With lvReceber
        .ListItems.Clear
        .ColumnHeaders.Clear
        .View = lvwReport
        .Gridlines = True
        .FullRowSelect = True

        .ColumnHeaders.Add , , "VENCIMENTO", 90
        .ColumnHeaders.Add , , "VALOR", 80
        .ColumnHeaders.Add , , "STATUS", 90
    End With

    Do While Not rs.EOF

        Set ITEM = lvReceber.ListItems.Add(, , Format(rs!VENCIMENTO, "dd/mm/yyyy"))
        ITEM.Tag = rs!IDARECEBER

        ITEM.ListSubItems.Add , , Format(NzDbl(rs!valor), "#,##0.00")
        ITEM.ListSubItems.Add , , Nz(rs!STATUS)

        rs.MoveNext

    Loop

    rs.Close
    Set rs = Nothing

    Exit Sub

TratarErro:

    modSistema.tela = "FluxoCaixa - Carregar lv receber"
    modSistema.DescErro = Err.Description
    modSistema.nErro = Err.Number

    Call modSistema.ReportarErro
    
    MsgBox "Erro: " & Err.Number & vbCrLf & _
                        Err.Description, vbInformation, "SISTEMA"

End Sub
