Attribute VB_Name = "Paletadecores"
Option Explicit

'==================================================
' PALETA GLOBAL ADEGA
' Módulo padrão: Paletadecores
'==================================================

'====================================================
' FUNDOS
'====================================================

' Cinza carvão escuro
' RGB(55, 58, 62)
' Hex: #373A3E
Public Const COR_FUNDO As Long = &H3E3A37

' Cinza grafite
' RGB(65, 68, 73)
' Hex: #414449
Public Const COR_FUNDO_SECUNDARIO As Long = &H494441

' Cinza ardósia
' RGB(78, 81, 87)
' Hex: #4E5157
Public Const COR_PAINEL As Long = &H57514E


'====================================================
' CAMPOS
'====================================================

' Cinza quase preto / carvão
' RGB(38, 39, 44)
' Hex: #26272C
Public Const COR_INPUT As Long = &H2C2726

' Branco
' RGB(255, 255, 255)
' Hex: #FFFFFF
Public Const COR_INPUT_TEXTO As Long = &HFFFFFF


'====================================================
' AÇÕES
'====================================================

' Rosa vinho / magenta escuro
' RGB(145, 55, 105)
' Hex: #913769
Public Const COR_VERMELHO As Long = &H693791

' Vinho escuro
' RGB(105, 40, 78)
' Hex: #69284E
Public Const COR_VERMELHO_ESCURO As Long = &H4E2869

' Magenta / rosa intenso
' RGB(175, 45, 125)
' Hex: #AF2D7D
Public Const COR_ACAO_SECUNDARIA As Long = &H7D2DAF


'====================================================
' TEXTOS
'====================================================

' Branco
' RGB(255, 255, 255)
' Hex: #FFFFFF
Public Const COR_TEXTO As Long = &HFFFFFF

' Roxo / magenta
' RGB(190, 195, 202)
' Hex: #BEC3CA
Public Const COR_TEXTO_SECUNDARIO As Long = &H7D2DAF
' Alternativa:
' RGB(202, 195, 190)
' Hex: #CAC3BE


'====================================================
' DESTAQUES
'====================================================

' Amarelo dourado
' RGB(255, 210, 85)
' Hex: #FFD255
Public Const COR_DESTAQUE As Long = &H55D2FF

' Azul claro / azul-ciano
' RGB(0, 168, 232)
' Hex: #00A8E8
Public Const COR_LABEL As Long = &HE8A800


'====================================================
' BORDAS
'====================================================

' Cinza médio
' RGB(105, 108, 113)
' Hex: #696C71
Public Const COR_BORDA As Long = &H716C69


'====================================================
' LISTVIEW
'====================================================

' Cinza muito claro
' RGB(225, 225, 225)
' Hex: #E1E1E1
Public Const COR_LISTVIEW_FUNDO As Long = &HE1E1E1

' Marrom muito escuro / quase preto
' RGB(30, 30, 36)
' Hex: #1E1E24
Public Const COR_LISTVIEW_TEXTO As Long = &H241E1E

'==================================================
' PALETA FLAME
'==================================================

Public Sub AplicarTema(ByVal formulario As Object)

    'Fundo do formulário
    formulario.BackColor = COR_FUNDO

    'Aplica o tema padrão recursivamente
    AplicarTemaContainer formulario

End Sub


Public Sub AplicarTemaContainer(ByVal container As Object)

    Dim controle As Object

    For Each controle In container.Controls

        'Processa somente filhos diretos do container atual
        If controle.Parent Is container Then

            Select Case TypeName(controle)

                Case "Frame"

                    controle.BackColor = COR_FUNDO_SECUNDARIO
                    controle.ForeColor = COR_TEXTO_SECUNDARIO

                    AplicarTemaContainer controle

                Case "CommandButton"

                    AplicarBotaoSecundario controle

                Case "TextBox"

                    AplicarTemaCampo controle

                Case "Label"

                    AplicarTemaLabel controle

                Case "ListBox", "ComboBox"

                    AplicarTemaCampo controle

                Case "CheckBox", "OptionButton", "ToggleButton"

                    controle.BackColor = COR_FUNDO_SECUNDARIO
                    controle.ForeColor = COR_TEXTO

                Case "Image"

                    'Não altera nenhuma propriedade das imagens.

            End Select

        End If

    Next controle

End Sub


Public Sub AplicarBotaoPrincipal(ByVal botao As Object)

    botao.BackColor = COR_VERMELHO
    botao.ForeColor = COR_TEXTO

End Sub


Public Sub AplicarBotaoSecundario(ByVal botao As Object)

    botao.BackColor = COR_PAINEL
    botao.ForeColor = COR_TEXTO

End Sub


Public Sub AplicarTemaCampo(ByVal controle As Object)

    controle.BackColor = COR_PAINEL
    controle.ForeColor = COR_TEXTO

    AplicarCorBorda controle, COR_BORDA

End Sub


Public Sub AplicarTemaLabel(ByVal rotulo As Object)

    rotulo.ForeColor = COR_TEXTO_SECUNDARIO

    'BackColor só terá efeito nos Labels
    'que já possuem BackStyle opaco.
    rotulo.BackStyle = fmBackStyleTransparent
    rotulo.BackColor = COR_FUNDO_SECUNDARIO

End Sub


Public Sub AplicarCabecalhoPainel(ByVal rotulo As Object)

    rotulo.BackColor = COR_FUNDO_SECUNDARIO
    rotulo.ForeColor = COR_TEXTO

End Sub


Public Sub AplicarTemaLv(ByVal listview As Object)

    listview.BackColor = COR_LISTVIEW_FUNDO
    listview.ForeColor = COR_LISTVIEW_TEXTO
    listview.Font.Size = 11
    listview.Gridlines = True

End Sub

Public Sub AplicarCorBorda( _
    ByVal controle As Object, _
    ByVal cor As Long)

    'Alguns tipos de controle não expõem BorderColor.
    On Error Resume Next

    controle.BorderColor = cor

    On Error GoTo 0

End Sub
