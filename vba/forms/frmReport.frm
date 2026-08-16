VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmReport 
   Caption         =   "REGISTRO DE OCORRÊNCIAS / FALHAS"
   ClientHeight    =   7260
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6825
   OleObjectBlob   =   "frmReport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Origem As String
Public IDOcorrencia As Long
Public VersaoSistema

Private Sub btnCancelar_Click()

    Unload Me

End Sub

Private Sub btnEditar_Click()

    Dim SQL As String
    
        SQL = "UPDATE TAB_OCORRENCIAS SET " & _
             "DATA = " & SqlData(txtData) & ", " & _
             "OCORRENCIA = " & SqlTexto(txtOcorrencia) & ", " & _
             "TELA = " & SqlTexto(txtTela) & ", " & _
             "DESCRICAO = " & SqlTexto(txtDescricao) & ", " & _
             "STATUS = " & SqlTexto(txtStatus) & ", " & _
             "IDUSUARIO = " & SqlNumero(IDUsuarioLogado) & ", " & _
             "VERSAO_SISTEMA = " & SqlTexto(VersaoSistema) & " " & _
             "WHERE IDREPORT = " & IDOcorrencia
      
      Conn.Execute SQL
      
    MsgBox "Ocorrência atualizada!", vbInformation
    
Unload Me

End Sub

Private Sub btnSalvar_Click()
        
    Dim SQL As String
    
    SQL = " INSERT INTO TAB_OCORRENCIAS(DATA, OCORRENCIA, TELA, DESCRICAO, STATUS, IDUSUARIO, VERSAO_SISTEMA) " & _
          " VALUES(" & SqlData(txtData) & ", " & _
                     SqlTexto(txtOcorrencia) & "," & _
                     SqlTexto(txtTela) & "," & _
                     SqlTexto(txtDescricao) & "," & _
                     SqlTexto(txtStatus) & "," & _
                     SqlNumero(IDUsuarioLogado) & "," & _
                     SqlTexto(VersaoSistema) & ")"

    Conn.Execute SQL
    
    MsgBox "Ocorrência Registrada com Sucesso!", vbInformation, "Sucesso"
    
    Unload Me
    
End Sub

Private Sub txtData_Change()

    FormatarData txtData

End Sub

Private Sub userform_activate()

  '========================================
    ' NOVA OCORRÊNCIA
    '========================================
    
    If IDOcorrencia <= 0 Then
        
        txtData.Value = Format(Date, "dd/mm/yyyy")
        txtStatus.Value = "PENDENTE"
        
        'Campos liberados
        txtData.Locked = False
        txtOcorrencia.Locked = False
        txtTela.Locked = False
        txtDescricao.Locked = False
        txtResposta.Locked = False
        
        'Botões
        btnSalvar.Visible = True
        btnEditar.Visible = False
        btnCancelar.Visible = True
        
        'Resposta não existe enquanto pendente
        txtResposta.Visible = False
        Label4.Visible = False
        
        Exit Sub
        
    End If


    '========================================
    ' OCORRÊNCIA EXISTENTE
    '========================================
    
    CarregarReport


    '========================================
    ' ORIGEM
    '========================================
    
    If Trim(CStr(Origem)) <> "" Then
        txtTela.Value = Origem
    End If


    '========================================
    ' OCORRÊNCIA PENDENTE
    '========================================
    
    If UCase(Trim(CStr(txtStatus.Value))) = "PENDENTE" Then
        
        'Inicialmente somente leitura
        txtData.Locked = True
        txtOcorrencia.Locked = False
        txtTela.Locked = False
        txtDescricao.Locked = False
        txtResposta.Locked = True
        
        'Resposta não aparece enquanto não foi concluída
        txtResposta.Visible = False
        Label4.Visible = False
        
        'Pode editar
        btnEditar.Visible = True
        btnSalvar.Visible = False
        btnCancelar.Visible = True
        
    End If


    '========================================
    ' OCORRÊNCIA CONCLUÍDA
    '========================================
    
    If UCase(Trim(CStr(txtStatus.Value))) = "CONCLUIDO" Then
        
        'Tudo bloqueado
        txtData.Locked = True
        txtOcorrencia.Locked = True
        txtTela.Locked = True
        txtDescricao.Locked = True
        txtResposta.Locked = True
        
        'Mostra resposta
        txtResposta.Visible = True
        Label4.Visible = True
        
        'Somente visualização
        btnSalvar.Visible = False
        btnEditar.Visible = False
        btnCancelar.Visible = False
        
    End If
        txtStatus.Enabled = False
        VersaoSistema = BuscarConfig("VERSAO_SISTEMA")

End Sub


Private Sub CarregarReport()

    Dim rs As ADODB.Recordset
    Dim SQL As String
    
    SQL = " SELECT TELA, OCORRENCIA, DESCRICAO, STATUS, DATA, RESPOSTA " & _
          " FROM TAB_OCORRENCIAS WHERE IDREPORT = " & IDOcorrencia
    
    Set rs = Conn.Execute(SQL)
    
        If Not rs.EOF Then
        
            txtStatus = Nz(rs!STATUS)
            txtOcorrencia = Nz(rs!OCORRENCIA)
            txtTela = Nz(rs!tela)
            txtDescricao = Nz(rs!DESCRICAO)
            txtData = NzDate(rs!Data)
            txtResposta = Nz(rs!Resposta)
        
        End If
    
    rs.Close
    Set rs = Nothing
    
End Sub

