VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLogEventos 
   Caption         =   "LOG EVENTOS"
   ClientHeight    =   8415.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13620
   ' OleObjectBlob removido na versao publica
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmLogEventos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Carregando As Boolean

Private Sub UserForm_initialize()

Carregando = True
    With lvLog
    
    .View = lvwReport
    .Gridlines = True
    .AllowColumnReorder = True
    .FullRowSelect = True
    .HideSelection = False
    
    .ColumnHeaders.Add , , "ID", 50
    .ColumnHeaders.Add , , "TIPO", 90
    .ColumnHeaders.Add , , "DESCRIÇÃO", 220
    .ColumnHeaders.Add , , "ID RELACIONADO", 90
    .ColumnHeaders.Add , , "DATA", 100
    .ColumnHeaders.Add , , "USUARIO", 90
    
    
    End With
    
CarregarFiltros

cmbTipoEvento.AddItem "TODOS", 0
cmbIdRel.AddItem "TODOS", 0

cmbTipoEvento.ListIndex = 0
cmbDataInicial.ListIndex = 0
cmbDataFinal.ListIndex = 0
cmbUsuario.ListIndex = 0
cmbIdRel.ListIndex = 0

Carregando = False
CarregarLog
    AtualizarTotal
End Sub

Public Sub AtualizarTela()

    CarregarLog
    AtualizarTotal

End Sub

Private Sub cmbIdRel_Change()
If Carregando Then Exit Sub
    AtualizarTela

End Sub

Private Sub cmbTipoEvento_Change()
If Carregando Then Exit Sub
    AtualizarTela
End Sub

Private Sub cmbDataInicial_Change()

    If Carregando Then Exit Sub
FormatarDatacmb cmbDataInicial
    ValidarPeriodo
    AtualizarTela

End Sub


Private Sub cmbDataFinal_Change()

    If Carregando Then Exit Sub
FormatarDatacmb cmbDataFinal
    ValidarPeriodo
    AtualizarTela

End Sub

Private Sub cmbUsuario_Change()
If Carregando Then Exit Sub
    AtualizarTela
End Sub

Public Sub CarregarLog()

On Error GoTo TratarErro

    Dim sql As String
    Dim Where As String
    Dim rs As ADODB.Recordset

    sql = "SELECT IDEVENTO AS ID, TIPOEVENTO AS TIPO, DESCRICAO, " & _
          "IDRELACIONADO, DATAEVENTO AS DATA, U.USUARIO AS USUARIO " & _
          "FROM TAB_LOGEVENTOS L " & _
          "JOIN TAB_USUARIOS U ON U.IDUSUARIO=L.IDUSUARIO "

    Where = ""

    If cmbTipoEvento.ListIndex > 0 Then
        Where = Where & " AND L.TIPOEVENTO='" & Replace(cmbTipoEvento.Value, "'", "''") & "'"
    End If

    If cmbDataInicial.ListIndex > 0 Then

    Where = Where & _
    " AND DATE(L.DATAEVENTO) >= DATE('" & _
    Format$(NzDate(cmbDataInicial.Value), "yyyy-mm-dd") & "')"

End If


If cmbDataFinal.ListIndex > 0 Then

    Where = Where & _
    " AND DATE(L.DATAEVENTO) <= DATE('" & _
    Format$(NzDate(cmbDataFinal.Value), "yyyy-mm-dd") & "')"

End If

    If cmbUsuario.ListIndex > 0 Then
        Where = Where & " AND L.IDUSUARIO=" & CLng(cmbUsuario.List(cmbUsuario.ListIndex, 0))
    End If

    If cmbIdRel.ListIndex > 0 Then
    Where = Where & " AND L.IDRELACIONADO = " & CLng(cmbIdRel.Value)
    End If

    If Where <> "" Then
        sql = sql & " WHERE " & Mid$(Where, 6)
    End If

    sql = sql & " ORDER BY L.IDEVENTO DESC"

    Set rs = Conn.Execute(sql)

    lvLog.ListItems.Clear

    Do While Not rs.EOF

        With lvLog.ListItems.Add(, , Nz(rs!id))
            .SubItems(1) = Nz(rs!tipo)
            .SubItems(2) = Nz(rs!DESCRICAO)
            .SubItems(3) = NzDbl(rs!IDRELACIONADO)
            .SubItems(4) = NzDate(rs!Data)
            .SubItems(5) = Nz(rs!Usuario)
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

Private Sub CarregarFiltros()

Dim rs As ADODB.Recordset
Dim sql As String

sql = "SELECT DISTINCT(TIPOEVENTO) AS EVENTO FROM tab_logeventos ORDER BY 1 asc"

Set rs = Conn.Execute(sql)

With cmbTipoEvento
.Clear

Do While Not rs.EOF
    .AddItem Nz(rs!evento)
        rs.MoveNext
    Loop
End With

rs.Close
Set rs = Nothing

sql = "SELECT DISTINCT DATE(DATAEVENTO) AS DATA " & _
      "FROM TAB_LOGEVENTOS " & _
      "ORDER BY 1 DESC"

Set rs = Conn.Execute(sql)

With cmbDataInicial

    .Clear
    .AddItem "TODOS"

    Do While Not rs.EOF

        .AddItem NzDate(rs!Data)
        rs.MoveNext

    Loop

End With

rs.Close
Set rs = Nothing


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

sql = "SELECT DISTINCT(L.IDUSUARIO) AS ID, U.USUARIO AS USUARIO FROM tab_logeventos L " & _
      "JOIN tab_usuarios U ON U.IDUSUARIO = L.IDUSUARIO "
      
Set rs = Conn.Execute(sql)

With cmbUsuario

    .Clear

    .ColumnCount = 2
    .ColumnWidths = "0 pt;120 pt"

    .AddItem 0
    .List(0, 1) = "TODOS"

    Do While Not rs.EOF

        .AddItem rs!id
        .List(.ListCount - 1, 1) = rs!Usuario

        rs.MoveNext

    Loop

End With

rs.Close
Set rs = Nothing

sql = "SELECT DISTINCT(IDRELACIONADO) AS IDREL FROM tab_logeventos ORDER BY 1 desc"

Set rs = Conn.Execute(sql)

With cmbIdRel
.Clear

    Do While Not rs.EOF
        .AddItem Nz(rs!IDREL)
            rs.MoveNext
        Loop
    End With
rs.Close
Set rs = Nothing

End Sub

Public Sub AtualizarTotal()

    Dim sql As String
    Dim Where As String
    Dim rs As ADODB.Recordset

    sql = "SELECT COUNT(*) AS TOTAL " & _
          "FROM TAB_LOGEVENTOS L " & _
          "JOIN TAB_USUARIOS U ON U.IDUSUARIO = L.IDUSUARIO "

    Where = ""

    If cmbTipoEvento.ListIndex > 0 Then
        Where = Where & _
            " AND L.TIPOEVENTO = '" & _
            Replace(cmbTipoEvento.Value, "'", "''") & "'"
    End If

    If cmbDataInicial.ListIndex > 0 Then

    Where = Where & _
    " AND DATE(L.DATAEVENTO) >= DATE('" & _
    Format$(NzDate(cmbDataInicial.Value), "yyyy-mm-dd") & "')"

End If


If cmbDataFinal.ListIndex > 0 Then

    Where = Where & _
    " AND DATE(L.DATAEVENTO) <= DATE('" & _
    Format$(NzDate(cmbDataFinal.Value), "yyyy-mm-dd") & "')"

End If

    If cmbUsuario.ListIndex > 0 Then
        Where = Where & _
            " AND L.IDUSUARIO = " & _
            CLng(cmbUsuario.List(cmbUsuario.ListIndex, 0))
    End If

If cmbIdRel.ListIndex > 0 Then
    Where = Where & " AND L.IDRELACIONADO = " & CLng(cmbIdRel.Value)
End If

    If Where <> "" Then
        sql = sql & " WHERE " & Mid$(Where, 6)
    End If

    Set rs = Conn.Execute(sql)

    txtTotal.Value = NzDbl(rs!Total)

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
