VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmConsulta 
   Caption         =   "CONSULTA DE PRODUTO"
   ClientHeight    =   4305
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7020
   ' OleObjectBlob removido na versao publica
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Userform_activate()

    CarregarProdutosCombo cmbProduto

    txtEAN5.SetFocus

End Sub

Private Sub txtEAN5_AfterUpdate()

    SelecionarProdutoPorEAN cmbProduto, txtEAN5.Value

End Sub

Private Sub cmbProduto_change()

    Dim rs As ADODB.Recordset
    Dim sql As String
    Dim IDProduto As Long

    If cmbProduto.ListIndex = -1 Then Exit Sub

    IDProduto = CLng(cmbProduto.List(cmbProduto.ListIndex, 0))

    sql = "SELECT PRECOVENDA, ESTOQUEATUAL, TIPO " & _
          "FROM TAB_PRODUTOS " & _
          "WHERE IDPRODUTO = " & IDProduto

    Set rs = Conn.Execute(sql)

    If Not rs.EOF Then

        txtPreco.Value = Format(IIf(IsNull(rs!PRECOVENDA), 0, rs!PRECOVENDA), "0.00")
        txtEstoque.Value = IIf(IsNull(rs!ESTOQUEATUAL), 0, rs!ESTOQUEATUAL)

    End If

    rs.Close
    Set rs = Nothing

End Sub

