VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmVendaDiaria 
   Caption         =   "VENDAS POR DIA"
   ClientHeight    =   8895.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10860
   ' OleObjectBlob removido na versao publica
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmVendaDiaria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ano As Long
Public mes As Long

Private Sub UserForm_initialize()

    With lvVendaDiaria
    
    .View = lvwReport
    .Gridlines = True
    .AllowColumnReorder = True
    .HideSelection = False
    .FullRowSelect = True
    
    lvVendaDiaria.ColumnHeaders.Clear
    
    .ColumnHeaders.Add , , "DATA VENDA", 100
    .ColumnHeaders.Add , , "VALOR TOTAL", 100
    .ColumnHeaders.Add , , "CUSTO", 100
    .ColumnHeaders.Add , , "LUCRO BRUTO", 100
    .ColumnHeaders.Add , , "MARGEM BRUTA", 120
    
    End With

End Sub

Private Sub Userform_activate()

    CarregarVendasMes

End Sub

Private Sub CarregarVendasMes()

    Dim sql As String
    Dim rs As ADODB.Recordset
    Dim item As ListItem
    
sql = "SELECT V.DATAVENDA AS DATA, " & _
        "SUM(V.VALORFINAL) AS TOTAL, C.CUSTO AS CUSTO, SUM(V.VALORFINAL) - C.CUSTO AS LUCRO, " & _
        "round((((SUM(V.VALORFINAL) - C.CUSTO) / sum(V.VALORFINAL)) * 100),2) AS MARGEM " & _
        "FROM tab_vendas V JOIN(SELECT V.DATAVENDA AS DATA, SUM(IV.CUSTOUNITARIO * IV.QUANTIDADE) AS CUSTO " & _
        "FROM tab_itensvenda IV JOIN tab_vendas V ON V.IDVENDA = IV.IDVENDA GROUP BY 1)C ON DATE(V.DATAVENDA) = DATE(C.DATA) " & _
        "WHERE YEAR(V.DATAVENDA) = " & ano & " AND MONTH(V.DATAVENDA) = " & mes & _
        " GROUP BY DATAVENDA " & _
        "ORDER BY DATA "
        
    Set rs = Conn.Execute(sql)
    
    lvVendaDiaria.ListItems.Clear
    
Do While Not rs.EOF
   
   Set item = lvVendaDiaria.ListItems.Add(, , rs!Data)

   item.ListSubItems.Add , , Format(NzDbl(rs!Total), "R$   #,##0.00")
   item.ListSubItems.Add , , Format(NzDbl(rs!Custo), "R$   #,##0.00")
   item.ListSubItems.Add , , Format(NzDbl(rs!lucro), "R$   #,##0.00")
   item.ListSubItems.Add , , Format(NzDbl(rs!margem) / 100, "0.00%")
   
rs.MoveNext

Loop

rs.Close
Set rs = Nothing

End Sub
