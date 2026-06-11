'=======================================================================
' MOD_ESTRUTURA
' Cria toda a estrutura da planilha de disparo de e-mails.
' Execute a macro CriarEstruturaPlanilha UMA vez apos colar os modulos.
'
' 100% compativel com Excel para Mac (Microsoft 365):
'   - Sem ActiveX (botoes criados como formas/Shapes com OnAction)
'   - Sem COM Automation / CreateObject
'   - Sem bibliotecas exclusivas do Windows
'=======================================================================
Option Explicit

' Nomes das abas (usados tambem pelos demais modulos)
Public Const ABA_LISTA As String = "Lista de Envios"
Public Const ABA_CONFIG As String = "Configurações"
Public Const ABA_HISTORICO As String = "Histórico"

'-----------------------------------------------------------------------
' CriarEstruturaPlanilha
' Cria (ou recria os cabecalhos de) todas as abas necessarias.
'-----------------------------------------------------------------------
Public Sub CriarEstruturaPlanilha()

    Application.ScreenUpdating = False

    CriarAbaLista
    CriarAbaConfiguracoes
    CriarAbaHistorico

    Worksheets(ABA_LISTA).Activate

    Application.ScreenUpdating = True

    MsgBox "Estrutura criada com sucesso!" & vbNewLine & vbNewLine & _
           "1) Preencha a aba '" & ABA_CONFIG & "'." & vbNewLine & _
           "2) Preencha a aba '" & ABA_LISTA & "'." & vbNewLine & _
           "3) Use o botao ENVIAR E-MAILS para disparar.", _
           vbInformation, "Disparo de E-mails"
End Sub

'-----------------------------------------------------------------------
' CriarAbaLista: aba principal com os destinatarios
'-----------------------------------------------------------------------
Private Sub CriarAbaLista()
    Dim ws As Worksheet
    Set ws = ObterOuCriarAba(ABA_LISTA)

    With ws
        ' Cabecalhos
        .Range("A1").Value = "Empresa"
        .Range("B1").Value = "Nome do contato"
        .Range("C1").Value = "E-mail principal"
        .Range("D1").Value = "E-mail CC"
        .Range("E1").Value = "Assunto"
        .Range("F1").Value = "Corpo do e-mail"
        .Range("G1").Value = "Status"
        .Range("H1").Value = "Data/Hora de envio"

        FormatarCabecalho .Range("A1:H1")

        ' Larguras de coluna
        .Columns("A").ColumnWidth = 22
        .Columns("B").ColumnWidth = 22
        .Columns("C").ColumnWidth = 30
        .Columns("D").ColumnWidth = 26
        .Columns("E").ColumnWidth = 35
        .Columns("F").ColumnWidth = 60
        .Columns("G").ColumnWidth = 28
        .Columns("H").ColumnWidth = 20

        ' Congela a linha de cabecalho
        .Activate
        .Range("A2").Select
        ActiveWindow.FreezePanes = True
    End With

    ' Botoes (formas com macro associada - sem ActiveX)
    CriarBotao ws, "btnEnviar", "ENVIAR E-MAILS", 540, 6, 130, 30, "EnviarEmailsEmMassa", RGB(0, 120, 60)
    CriarBotao ws, "btnLimpar", "Limpar Status", 680, 6, 110, 30, "LimparStatus", RGB(120, 120, 120)
End Sub

'-----------------------------------------------------------------------
' CriarAbaConfiguracoes: parametros gerais do disparo
'-----------------------------------------------------------------------
Private Sub CriarAbaConfiguracoes()
    Dim ws As Worksheet
    Set ws = ObterOuCriarAba(ABA_CONFIG)

    With ws
        .Range("A1").Value = "CONFIGURAÇÕES DE ENVIO"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14

        .Range("A3").Value = "Caminho do anexo (opcional):"
        .Range("A4").Value = "Modo de envio:"
        .Range("A5").Value = "Intervalo entre envios (segundos):"
        .Range("A6").Value = "Assinatura padrão:"

        .Range("A3:A6").Font.Bold = True
        .Columns("A").ColumnWidth = 34
        .Columns("B").ColumnWidth = 70

        ' Valores iniciais
        If .Range("B4").Value = "" Then .Range("B4").Value = "REVISAO"
        If .Range("B5").Value = "" Then .Range("B5").Value = 3

        ' Lista suspensa para o modo de envio (validacao de dados nativa)
        With .Range("B4").Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                 Formula1:="REVISAO,DIRETO"
            .InputMessage = "REVISAO = abre cada e-mail para conferência antes do envio." & _
                            " DIRETO = envia imediatamente."
            .ShowInput = True
        End With

        ' Notas de ajuda
        .Range("A8").Value = "Notas:"
        .Range("A8").Font.Bold = True
        .Range("A9").Value = "• Anexo: informe o caminho completo, ex.: /Users/seunome/Documents/proposta.pdf"
        .Range("A10").Value = "• Personalização: use {{EMPRESA}} e {{NOME}} no Assunto e no Corpo do e-mail."
        .Range("A11").Value = "• Vários destinatários no mesmo campo: separe com ponto e vírgula ( ; )."
        .Range("A12").Value = "• A assinatura é adicionada automaticamente ao final de cada e-mail."
    End With
End Sub

'-----------------------------------------------------------------------
' CriarAbaHistorico: log de todos os envios
'-----------------------------------------------------------------------
Private Sub CriarAbaHistorico()
    Dim ws As Worksheet
    Set ws = ObterOuCriarAba(ABA_HISTORICO)

    With ws
        .Range("A1").Value = "Número do envio"
        .Range("B1").Value = "Data/Hora"
        .Range("C1").Value = "Destinatário"
        .Range("D1").Value = "Assunto"
        .Range("E1").Value = "Status"
        .Range("F1").Value = "Lote de envio"

        FormatarCabecalho .Range("A1:F1")

        .Columns("A").ColumnWidth = 16
        .Columns("B").ColumnWidth = 20
        .Columns("C").ColumnWidth = 32
        .Columns("D").ColumnWidth = 40
        .Columns("E").ColumnWidth = 30
        .Columns("F").ColumnWidth = 24
    End With
End Sub

'-----------------------------------------------------------------------
' Funcoes auxiliares de construcao
'-----------------------------------------------------------------------

' Devolve a aba pelo nome; cria se nao existir
Private Function ObterOuCriarAba(nomeAba As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Worksheets(nomeAba)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = Worksheets.Add(After:=Worksheets(Worksheets.Count))
        ws.Name = nomeAba
    End If
    Set ObterOuCriarAba = ws
End Function

' Aplica formatacao padrao aos cabecalhos
Private Sub FormatarCabecalho(rng As Range)
    With rng
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 22
    End With
End Sub

' Cria um botao como forma (Shape) com macro associada - compativel com Mac
Private Sub CriarBotao(ws As Worksheet, nome As String, texto As String, _
                       esquerda As Single, topo As Single, _
                       largura As Single, altura As Single, _
                       macro As String, cor As Long)
    Dim shp As Shape

    ' Remove botao anterior com o mesmo nome, se existir
    On Error Resume Next
    ws.Shapes(nome).Delete
    On Error GoTo 0

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, esquerda, topo, largura, altura)
    With shp
        .Name = nome
        .Fill.ForeColor.RGB = cor
        .Line.Visible = msoFalse
        .TextFrame2.TextRange.Text = texto
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Size = 11
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame2.HorizontalAnchor = msoAnchorCenter
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .OnAction = macro
    End With
End Sub
