VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMovimentoEstoque 
   Caption         =   "MOVIMENTO ESTOQUE"
   ClientHeight    =   9420.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19755
   OleObjectBlob   =   "frmMovimentoEstoque.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMovimentoEstoque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Carregando As Boolean

Private Sub UserForm_Initialize()

    Carregando = True

    With lvMovimento

        .View = lvwReport
        .Gridlines = True
        .AllowColumnReorder = True
        .FullRowSelect = True
        .HideSelection = False

        .ColumnHeaders.Add , , "ID", 50
        .ColumnHeaders.Add , , "PRODUTO", 180
        .ColumnHeaders.Add , , "TIPO", 90
        .ColumnHeaders.Add , , "QUANTIDADE", 80
        .ColumnHeaders.Add , , "CUSTO", 80
        .ColumnHeaders.Add , , "EST. ANT.", 80
        .ColumnHeaders.Add , , "EST. POST.", 80
        .ColumnHeaders.Add , , "DATA", 100
        .ColumnHeaders.Add , , "USUARIO", 100
        .ColumnHeaders.Add , , "OBSERVAÇÃO", 200

    End With


    CarregarFiltros
CarregarProdutosCombo cmbProduto

cmbProduto.ListIndex = 0
cmbTipoMovimento.ListIndex = 0
cmbDataInicial.ListIndex = 0
cmbDataFinal.ListIndex = 0
cmbUsuario.ListIndex = 0


    Carregando = False

cmbProduto.AddItem "TODOS", 0
cmbProduto.ListIndex = 0
    CarregarMovimento
    AtualizarTotal

End Sub

Private Sub CarregarFiltros()

    Dim rs As ADODB.Recordset
    Dim sql As String

    '==================================
    ' TIPO MOVIMENTO
    '==================================

    sql = "SELECT DISTINCT TIPOMOVIMENTO " & _
          "FROM TAB_MOVIMENTOESTOQUE " & _
          "ORDER BY 1"

    Set rs = Conn.Execute(sql)

    With cmbTipoMovimento

        .Clear

        .AddItem "TODOS"

        Do While Not rs.EOF

            .AddItem Nz(rs!TIPOMOVIMENTO)

            rs.MoveNext

        Loop

    End With

    rs.Close
    Set rs = Nothing



    '==================================
    ' DATA MOVIMENTO
    '==================================

    sql = "SELECT DISTINCT DATE(DATAMOVIMENTO) AS DATA " & _
      "FROM TAB_MOVIMENTOESTOQUE " & _
      "ORDER BY 1"

Set rs = Conn.Execute(sql)

With cmbDataInicial

    .Clear
    .AddItem "TODOS"

    Do While Not rs.EOF

        .AddItem NzDate(rs!Data)
        rs.MoveNext

    Loop

End With

sql = "SELECT DISTINCT DATE(DATAMOVIMENTO) AS DATA " & _
      "FROM TAB_MOVIMENTOESTOQUE " & _
      "ORDER BY 1"

Set rs = Conn.Execute(sql)

With cmbDataFinal

    .Clear
    .AddItem "TODOS"

    Do While Not rs.EOF

        .AddItem NzDate(rs!Data)
        rs.MoveNext

    Loop

End With

    rs.Close
    Set rs = Nothing



    '==================================
    ' USUARIO
    '==================================

    sql = "SELECT DISTINCT(M.IDUSUARIO), U.USUARIO " & _
          "FROM TAB_MOVIMENTOESTOQUE M " & _
          "JOIN TAB_USUARIOS U ON U.IDUSUARIO=M.IDUSUARIO " & _
          "ORDER BY U.USUARIO"


    Set rs = Conn.Execute(sql)


    With cmbUsuario

        .Clear

        .ColumnCount = 2
        .ColumnWidths = "0 pt;120 pt"


        .AddItem 0
        .List(0, 1) = "TODOS"


        Do While Not rs.EOF

            .AddItem rs!idUsuario
            .List(.ListCount - 1, 1) = rs!Usuario

            rs.MoveNext

        Loop

    End With


    rs.Close
    Set rs = Nothing


End Sub

Private Sub ValidarPeriodo()

    If Nz(cmbDataInicial.Value, "") <> "" And _
       Nz(cmbDataFinal.Value, "") <> "" Then


        If cmbDataInicial.ListIndex > 0 And _
           cmbDataFinal.ListIndex > 0 Then


            If NzDate(cmbDataInicial.Value) > NzDate(cmbDataFinal.Value) Then

                MsgBox "A data inicial não pode ser maior que a data final.", vbExclamation


                cmbDataFinal.ListIndex = 0


            End If


        End If


    End If

End Sub

Public Sub AtualizarTela()

    CarregarMovimento
    AtualizarTotal

End Sub


Private Sub cmbProduto_change()

    If Carregando Then Exit Sub

    AtualizarTela

End Sub


Private Sub cmbTipoMovimento_Change()

    If Carregando Then Exit Sub

    AtualizarTela

End Sub

Private Sub cmbDataInicial_change()

    If Carregando Then Exit Sub
    FormatarData cmbDataInicial
    ValidarPeriodo
    AtualizarTela

End Sub


Private Sub cmbDataFinal_change()

    If Carregando Then Exit Sub
FormatarData cmbDataFinal
    ValidarPeriodo
    AtualizarTela

End Sub


Private Sub cmbUsuario_change()

    If Carregando Then Exit Sub

    AtualizarTela

End Sub

Private Function FiltroMovimento() As String

    Dim Where As String

    Where = ""


    If cmbProduto.ListIndex > 0 Then

        Where = Where & _
        " AND M.IDPRODUTO=" & _
        CLng(cmbProduto.List(cmbProduto.ListIndex, 0))

    End If


    If cmbTipoMovimento.ListIndex > 0 Then

        Where = Where & _
        " AND M.TIPOMOVIMENTO='" & _
        Replace(cmbTipoMovimento.Value, "'", "''") & "'"

    End If


    If cmbDataInicial.ListIndex > 0 Then

    Where = Where & _
    " AND DATE(M.DATAMOVIMENTO) >= DATE('" & _
    Format$(NzDate(cmbDataInicial.Value), "yyyy-mm-dd") & "')"

End If


If cmbDataFinal.ListIndex > 0 Then

    Where = Where & _
    " AND DATE(M.DATAMOVIMENTO) <= DATE('" & _
    Format$(NzDate(cmbDataFinal.Value), "yyyy-mm-dd") & "')"

End If


    If cmbUsuario.ListIndex > 0 Then

        Where = Where & _
        " AND M.IDUSUARIO=" & _
        CLng(cmbUsuario.List(cmbUsuario.ListIndex, 0))

    End If


    If Where <> "" Then

        FiltroMovimento = " WHERE " & Mid$(Where, 6)

    Else

        FiltroMovimento = ""

    End If


End Function

Public Sub CarregarMovimento()

On Error GoTo TratarErro

    Dim sql As String
    Dim rs As ADODB.Recordset


    sql = "SELECT " & _
          "M.IDMOVIMENTO AS ID, " & _
          "P.NOME AS PRODUTO, " & _
          "M.TIPOMOVIMENTO AS TIPO, " & _
          "M.QUANTIDADE, " & _
          "M.CUSTOUNITARIO, " & _
          "M.ESTOQUEANTERIOR, " & _
          "M.ESTOQUEPOSTERIOR, " & _
          "M.DATAMOVIMENTO, " & _
          "U.USUARIO, " & _
          "M.OBSERVACAO " & _
          "FROM TAB_MOVIMENTOESTOQUE M " & _
          "JOIN TAB_PRODUTOS P ON P.IDPRODUTO=M.IDPRODUTO " & _
          "LEFT JOIN TAB_USUARIOS U ON U.IDUSUARIO=M.IDUSUARIO "


    sql = sql & FiltroMovimento()


    sql = sql & " ORDER BY M.IDMOVIMENTO DESC"


    Set rs = Conn.Execute(sql)


    lvMovimento.ListItems.Clear


    Do While Not rs.EOF


        With lvMovimento.ListItems.Add(, , Nz(rs!ID))

            .SubItems(1) = Nz(rs!PRODUTO)
            .SubItems(2) = Nz(rs!tipo)
            .SubItems(3) = NzDbl(rs!QUANTIDADE)
            .SubItems(4) = NzDbl(rs!CUSTOUNITARIO)
            .SubItems(5) = NzDbl(rs!ESTOQUEANTERIOR)
            .SubItems(6) = NzDbl(rs!ESTOQUEPOSTERIOR)
            .SubItems(7) = NzDate(rs!DATAMOVIMENTO)
            .SubItems(8) = Nz(rs!Usuario)
            .SubItems(9) = Nz(rs!OBSERVACAO)

        End With


        rs.MoveNext


    Loop


    rs.Close
    Set rs = Nothing


Exit Sub


TratarErro:

    If Not rs Is Nothing Then

        If rs.State = adStateOpen Then rs.Close

        Set rs = Nothing

    End If


    MsgBox "Erro " & Err.Number & vbCrLf & Err.Description, vbExclamation


End Sub

Public Sub AtualizarTotal()

On Error GoTo TratarErro

    Dim sql As String
    Dim rs As ADODB.Recordset


    sql = "SELECT COUNT(*) AS TOTAL " & _
          "FROM TAB_MOVIMENTOESTOQUE M " & _
          "JOIN TAB_PRODUTOS P ON P.IDPRODUTO=M.IDPRODUTO " & _
          "LEFT JOIN TAB_USUARIOS U ON U.IDUSUARIO=M.IDUSUARIO "


    sql = sql & FiltroMovimento()


    Set rs = Conn.Execute(sql)


    txtTotal.Value = NzDbl(rs!Total)


    rs.Close
    Set rs = Nothing


Exit Sub


TratarErro:

    If Not rs Is Nothing Then

        If rs.State = adStateOpen Then rs.Close

        Set rs = Nothing

    End If


    MsgBox "Erro " & Err.Number & vbCrLf & Err.Description, vbExclamation


End Sub
