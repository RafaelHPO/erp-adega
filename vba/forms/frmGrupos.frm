VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmGrupos 
   Caption         =   "CATEGORIAS"
   ClientHeight    =   6810
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7710
   OleObjectBlob   =   "frmGrupos.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmGrupos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AplicarPaleta()

    AplicarTema Me
    AplicarTemaLv lvGrupos
    
End Sub

Private Sub UserForm_initialize()

    With lvGrupos
        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = True
        .HideSelection = False
        .ColumnHeaders.Clear
    End With

    With lvGrupos.ColumnHeaders
        .Add , , "ID", 60
        .Add , , "GRUPO", 200
    End With

    CarregarGrupos

End Sub

Public Sub CarregarGrupos()

    Dim rs As ADODB.Recordset
    Dim ITEM As ListItem

    Set rs = Conn.Execute("SELECT IDGRUPO,DESCRICAO FROM TAB_GRUPOS ORDER BY DESCRICAO")

    lvGrupos.ListItems.Clear

    Do While Not rs.EOF

        Set ITEM = lvGrupos.ListItems.Add(, , rs!IdGrupo)

        ITEM.SubItems(1) = Nz(rs!DESCRICAO)

        rs.MoveNext

    Loop

    rs.Close
    Set rs = Nothing

End Sub

Private Sub btnNovoGrupo_Click()

    frmCadGrupo.Tag = ""

    frmCadGrupo.Show vbModal

    CarregarGrupos

End Sub

Private Sub btnEditarGrupo_Click()

    If lvGrupos.SelectedItem Is Nothing Then

        MsgBox "Selecione um grupo.", vbExclamation
        Exit Sub

    End If


    frmCadGrupo.Tag = lvGrupos.SelectedItem.Text

    frmCadGrupo.Show vbModal

    CarregarGrupos

End Sub
Private Sub btnExcluirGrupo_Click()

Dim rs As ADODB.Recordset
Dim sql As String
Dim resp As VbMsgBoxResult


If lvGrupos.SelectedItem Is Nothing Then
    MsgBox "Selecione um grupo.", vbExclamation
    Exit Sub
End If


resp = MsgBox("Excluir grupo selecionado?", vbYesNo + vbQuestion)

If resp = vbNo Then Exit Sub


sql = "CALL PROC_EXCLUIRGRUPO(" & _
      lvGrupos.SelectedItem.Text & ")"


Set rs = Conn.Execute(sql)


If Not rs.EOF Then
    MsgBox rs!Retorno, vbInformation
End If


rs.Close
Set rs = Nothing


CarregarGrupos

End Sub

