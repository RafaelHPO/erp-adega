VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSenha 
   Caption         =   "SENHA DO USUARIO"
   ClientHeight    =   1710
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4320
   OleObjectBlob   =   "frmSenha.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSenha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public SenhaValida As Boolean
Public SenhaDigitada As String

Private Sub btnConfirmar_Click()

    If ValidarSenha(txtSenha.Value) Then
        SenhaValida = True
    Else
        SenhaValida = False
        MsgBox "Senha inválida", vbExclamation
    End If

    Me.Hide

End Sub

