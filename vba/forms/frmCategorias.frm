VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCategorias 
   Caption         =   "CATEGORIAS"
   ClientHeight    =   6810
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7710
   OleObjectBlob   =   "frmCategorias.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCategorias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub AplicarPaleta()

    AplicarTema Me
    AplicarTemaLv lvCategorias
    AplicarBotaoPrincipal btnNovaCategoria
    AplicarBotaoPrincipal btnExcluirCategoria
    
End Sub

Private Sub UserForm_initialize()

Label1.Visible = False
Label2.Visible = False
Label3.Visible = False
txtIdCat.Visible = False
txtDescricao.Visible = False
txtMarkup.Visible = False
btnSalvar.Visible = False

    With lvCategorias

        .View = lvwReport
        .Gridlines = False
        .FullRowSelect = True
        .HideSelection = False

        .ColumnHeaders.Clear

        .ColumnHeaders.Add , , "ID", 50
        .ColumnHeaders.Add , , "CATEGORIA", 180
        .ColumnHeaders.Add , , "MARKUP (%)", 80

    End With

    CarregarCategorias
    AplicarPaleta

End Sub

Private Sub btnNovaCategoria_Click()

    Label1.Visible = True
    Label2.Visible = True
    Label3.Visible = True

    txtIdCat.Visible = True
    txtDescricao.Visible = True
    txtMarkup.Visible = True

    btnSalvar.Visible = True

    txtIdCat.Value = ""
    txtDescricao.Value = ""
    txtMarkup.Value = ""

    txtDescricao.SetFocus

End Sub

Public Sub CarregarCategorias()

    Dim rs As ADODB.Recordset
    Dim ITEM As ListItem

    Set rs = Conn.Execute( _
        "SELECT IDCATEGORIA, DESCRICAO, MARKUP " & _
        "FROM TAB_CATEGORIAS " & _
        "ORDER BY DESCRICAO")

    lvCategorias.ListItems.Clear

    Do While Not rs.EOF

        Set ITEM = lvCategorias.ListItems.Add(, , rs!IdCategoria)

        ITEM.SubItems(1) = Nz(rs!DESCRICAO)
        ITEM.SubItems(2) = Format(Nz(rs!Markup), "0") & " %"

        rs.MoveNext

    Loop

    rs.Close
    Set rs = Nothing

End Sub
Private Sub btnEditarCategoria_Click()

    Dim rs As ADODB.Recordset
    Dim sql As String
    Dim IdCategoria As Long

    If lvCategorias.SelectedItem Is Nothing Then

        MsgBox "Selecione uma categoria.", vbExclamation
        Exit Sub

    End If

    IdCategoria = CLng(lvCategorias.SelectedItem.Text)

    sql = "SELECT IDCATEGORIA, DESCRICAO, MARKUP " & _
          "FROM TAB_CATEGORIAS " & _
          "WHERE IDCATEGORIA = " & IdCategoria

    Set rs = Conn.Execute(sql)

    If Not rs.EOF Then

        Label1.Visible = True
        Label2.Visible = True
        Label3.Visible = True

        txtIdCat.Visible = True
        txtDescricao.Visible = True
        txtMarkup.Visible = True

        btnSalvar.Visible = True

        txtIdCat.Value = rs!IdCategoria
        txtDescricao.Value = Nz(rs!DESCRICAO)
        txtMarkup.Value = Nz(rs!Markup)

    End If

    rs.Close
    Set rs = Nothing

End Sub

Private Sub btnSalvar_Click()

    Dim sql As String

    If Trim(txtDescricao.Value) = "" Then

        MsgBox "Informe a descrição.", vbExclamation
        Exit Sub

    End If

    If Trim(txtMarkup.Value) = "" Then

        MsgBox "Informe o markup.", vbExclamation
        Exit Sub

    End If

    If Trim(txtIdCat.Value) = "" Then

        sql = "INSERT INTO TAB_CATEGORIAS " & _
              "(DESCRICAO, MARKUP) VALUES (" & _
              SqlTexto(txtDescricao.Value) & ", " & _
              SqlNumero(txtMarkup.Value) & ")"

    Else

        sql = "UPDATE TAB_CATEGORIAS SET " & _
              "DESCRICAO = " & SqlTexto(txtDescricao.Value) & ", " & _
              "MARKUP = " & SqlNumero(txtMarkup.Value) & " " & _
              "WHERE IDCATEGORIA = " & CLng(txtIdCat.Value)

    End If

    Conn.Execute sql

    MsgBox "Categoria salva com sucesso.", vbInformation

    CarregarCategorias

    txtIdCat.Value = ""
    txtDescricao.Value = ""
    txtMarkup.Value = ""

End Sub

