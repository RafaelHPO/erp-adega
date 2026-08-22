VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLicenca 
   Caption         =   "LICENÇA"
   ClientHeight    =   5190
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5430
   OleObjectBlob   =   "frmLicenca.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmLicenca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AplicarPaleta()

    AplicarTema Me
    AplicarBotaoPrincipal btnSalvar
    
End Sub


Private Sub UserForm_initialize()

    lblTitulo.Caption = "ATIVAÇÃO DO SISTEMA"

    lblStatus.Caption = MotivoLicenca

    lblMensagem.Caption = MensagemLicenca


    If DataValidadeLicenca > 0 Then

        lblValidade.Caption = "Validade: " & _
                              Format(DataValidadeLicenca, "dd/mm/yyyy")

    Else

        lblValidade.Caption = ""

    End If


    lblVersao.Caption = "Versão " & BuscarConfig("VERSAO_SISTEMA")


    txtChave.Value = ""
    
    AplicarPaleta

End Sub

Private Sub btnAtivar_Click()

    Dim rs As ADODB.Recordset
    Dim sql As String

    If Trim(txtChave.Value) = "" Then

        MsgBox "Informe a chave de ativação.", vbExclamation
        txtChave.SetFocus
        Exit Sub

    End If


    sql = "CALL PROC_ATIVARLICENCA('" & Replace(txtChave.Value, "'", "''") & "')"


    Set rs = Conn.Execute(sql)


    MsgBox rs!MSG, IIf(rs!SUCESSO = 1, vbInformation, vbCritical)


    If rs!SUCESSO = 1 Then


        rs.Close
        Set rs = Nothing


        Unload Me

        frmLogin.Show


    Else


        txtChave.SelStart = 0
        txtChave.SelLength = Len(txtChave.Text)
        txtChave.SetFocus


        rs.Close
        Set rs = Nothing


    End If


End Sub

Private Sub txtChave_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    If KeyCode = vbKeyReturn Then

        KeyCode = 0

        btnAtivar_Click

    End If

End Sub

Private Sub btnFechar_Click()

    Call FecharSistema

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

  If CloseMode = vbFormControlMenu Then
       Cancel = True
       Call FecharSistema
   End If

End Sub
