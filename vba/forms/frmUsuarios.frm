VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmUsuarios 
   Caption         =   "USUÁRIOS"
   ClientHeight    =   6750
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7905
   OleObjectBlob   =   "frmUsuarios.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AplicarPaleta()

    AplicarTema Me
    AplicarBotaoPrincipal btnNovoCad
    AplicarTemaLv lvUsuarios

End Sub

Private Sub UserForm_initialize()

  With lvUsuarios
        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = True
        .HideSelection = False
        .ColumnHeaders.Clear
    End With

    'Cabeçalhos
    With lvUsuarios.ColumnHeaders
        .Add , , "ID", 90
        .Add , , "Nome", 160
        .Add , , "Usuário", 120
    End With

    CarregarUsuarios
    AplicarPaleta
    
End Sub

Private Sub btnNovoCad_Click()

frmCadUsuario.Show vbModeless

End Sub


Public Sub CarregarUsuarios()

    Dim rs As ADODB.Recordset
    Dim sql As String
    Dim ITEM As ListItem

    sql = "SELECT IDUSUARIO, NOME, USUARIO FROM TAB_USUARIOS"

    Set rs = New ADODB.Recordset
    rs.Open sql, Conn, adOpenStatic, adLockReadOnly

    lvUsuarios.ListItems.Clear

    Do While Not rs.EOF

        Set ITEM = lvUsuarios.ListItems.Add(, , rs!idUsuario)
        ITEM.SubItems(1) = rs!NOME
        ITEM.SubItems(2) = rs!Usuario

        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing

End Sub

Private Sub btnEditarUsuario_Click()

    If lvUsuarios.SelectedItem Is Nothing Then
        MsgBox "Selecione um usuário", vbExclamation
        Exit Sub
    End If

    Dim idUsuario As Long
    idUsuario = lvUsuarios.SelectedItem.Text

    frmCadUsuario.Tag = idUsuario
    frmCadUsuario.Show vbModeless

End Sub

Private Sub btnExcluirUsuario_Click()

    Dim sql As String
    Dim resp As VbMsgBoxResult
    Dim idUsuario As Long

    If lvUsuarios.SelectedItem Is Nothing Then
        MsgBox "Selecione um usuario", vbExclamation
        Exit Sub
    End If

    idUsuario = CLng(lvUsuarios.SelectedItem.Text)

    resp = MsgBox("Tem certeza que deseja excluir este usuario?", _
                  vbYesNo + vbQuestion, _
                  "Confirmação")

    If resp = vbNo Then Exit Sub

    sql = "DELETE FROM TAB_USUARIOS WHERE IDUSUARIO = " & idUsuario

    Conn.Execute sql

    MsgBox "Usuario excluído com sucesso!", vbInformation

    Call CarregarUsuarios   ' recarrega o ListView

End Sub

