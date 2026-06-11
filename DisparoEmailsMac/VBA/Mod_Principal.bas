'=======================================================================
' MOD_PRINCIPAL
' Rotina principal de disparo de e-mails em massa.
'
' Fluxo:
'   1. Le as configuracoes (anexo, modo, intervalo, assinatura)
'   2. Valida o anexo (se informado) ANTES de iniciar o lote
'   3. Percorre a aba "Lista de Envios" linha a linha:
'        - pula linhas ja enviadas
'        - valida e-mails (vazios e formato)
'        - aplica personalizacao {{EMPRESA}} / {{NOME}}
'        - chama o Outlook para Mac via AppleScriptTask
'        - atualiza Status / Data-Hora e grava no Historico
'        - aguarda o intervalo configurado
'   4. Exibe o relatorio final (enviados x erros)
'
' Integracao com o Outlook: EXCLUSIVAMENTE via AppleScriptTask, que
' executa o arquivo EnviarEmailOutlook.applescript instalado em:
'   ~/Library/Application Scripts/com.microsoft.Excel/
' (Sem ActiveX, sem COM, sem CreateObject, sem APIs do Windows.)
'=======================================================================
Option Explicit

' Nome do arquivo AppleScript e do handler chamado dentro dele
Private Const ARQUIVO_APPLESCRIPT As String = "EnviarEmailOutlook.applescript"
Private Const HANDLER_APPLESCRIPT As String = "enviarEmail"

' Colunas da aba Lista de Envios
Private Const COL_EMPRESA As Long = 1   ' A
Private Const COL_NOME As Long = 2      ' B
Private Const COL_EMAIL As Long = 3     ' C
Private Const COL_CC As Long = 4        ' D
Private Const COL_ASSUNTO As Long = 5   ' E
Private Const COL_CORPO As Long = 6     ' F
Private Const COL_STATUS As Long = 7    ' G
Private Const COL_DATAHORA As Long = 8  ' H

'-----------------------------------------------------------------------
' EnviarEmailsEmMassa - macro principal (associada ao botao ENVIAR E-MAILS)
'-----------------------------------------------------------------------
Public Sub EnviarEmailsEmMassa()

    Dim wsLista As Worksheet
    Dim wsConfig As Worksheet

    ' --- Configuracoes ---
    Dim caminhoAnexo As String
    Dim modoEnvio As String
    Dim intervaloSegundos As Double
    Dim assinatura As String

    ' --- Controle do lote ---
    Dim lote As String
    Dim ultimaLinha As Long
    Dim linha As Long
    Dim totalProcessados As Long
    Dim totalEnviados As Long
    Dim totalErros As Long
    Dim totalPulados As Long

    ' --- Dados de cada linha ---
    Dim empresa As String, nomeContato As String
    Dim emailPrincipal As String, emailCC As String
    Dim assunto As String, corpo As String
    Dim statusAtual As String
    Dim parametros As String
    Dim resultado As String
    Dim statusLinha As String

    ' ================= 1. Validacoes iniciais =================
    If Not AbasExistem() Then
        MsgBox "Estrutura não encontrada. Execute primeiro a macro 'CriarEstruturaPlanilha'.", _
               vbExclamation, "Disparo de E-mails"
        Exit Sub
    End If

    Set wsLista = Worksheets(ABA_LISTA)
    Set wsConfig = Worksheets(ABA_CONFIG)

    ' Le as configuracoes
    caminhoAnexo = Trim(CStr(wsConfig.Range("B3").Value))
    modoEnvio = UCase(Trim(CStr(wsConfig.Range("B4").Value)))
    intervaloSegundos = Val(wsConfig.Range("B5").Value)
    assinatura = CStr(wsConfig.Range("B6").Value)

    If modoEnvio <> "DIRETO" Then modoEnvio = "REVISAO" ' padrao seguro

    ' Verifica se ha linhas para processar
    ultimaLinha = wsLista.Cells(wsLista.Rows.Count, COL_EMAIL).End(xlUp).Row
    If ultimaLinha < 2 Then
        ' Considera tambem linhas com empresa preenchida mas e-mail vazio
        ultimaLinha = wsLista.Cells(wsLista.Rows.Count, COL_EMPRESA).End(xlUp).Row
        If ultimaLinha < 2 Then
            MsgBox "Nenhum destinatário encontrado na aba '" & ABA_LISTA & "'.", _
                   vbExclamation, "Disparo de E-mails"
            Exit Sub
        End If
    Else
        ' Garante que linhas com empresa alem do ultimo e-mail tambem entrem
        If wsLista.Cells(wsLista.Rows.Count, COL_EMPRESA).End(xlUp).Row > ultimaLinha Then
            ultimaLinha = wsLista.Cells(wsLista.Rows.Count, COL_EMPRESA).End(xlUp).Row
        End If
    End If

    ' Verifica a existencia do anexo ANTES de iniciar o lote
    If caminhoAnexo <> "" Then
        If Not ArquivoExiste(caminhoAnexo) Then
            MsgBox "Anexo não encontrado:" & vbNewLine & caminhoAnexo & vbNewLine & vbNewLine & _
                   "Corrija o caminho na aba '" & ABA_CONFIG & "' (célula B3) ou deixe em branco " & _
                   "para enviar sem anexo.", vbCritical, "Anexo não encontrado"
            Exit Sub
        End If
    End If

    ' Confirmacao do usuario antes de iniciar
    If MsgBox("Iniciar o disparo?" & vbNewLine & vbNewLine & _
              "Linhas na lista: " & (ultimaLinha - 1) & vbNewLine & _
              "Modo de envio: " & IIf(modoEnvio = "DIRETO", "DIRETO (envia imediatamente)", "REVISÃO (abre rascunhos para conferência)") & vbNewLine & _
              "Anexo: " & IIf(caminhoAnexo = "", "(sem anexo)", caminhoAnexo) & vbNewLine & _
              "Intervalo entre envios: " & intervaloSegundos & " s", _
              vbYesNo + vbQuestion, "Confirmar disparo") = vbNo Then Exit Sub

    ' Identificador unico do lote (aparece no Historico)
    lote = "LOTE-" & Format(Now, "yyyymmdd-hhmmss")

    ' ================= 2. Loop de envio =================
    For linha = 2 To ultimaLinha

        ' Pula linhas totalmente vazias
        If Trim(CStr(wsLista.Cells(linha, COL_EMPRESA).Value)) = "" And _
           Trim(CStr(wsLista.Cells(linha, COL_EMAIL).Value)) = "" Then GoTo ProximaLinha

        ' Pula linhas ja enviadas (permite retomar um lote interrompido)
        statusAtual = CStr(wsLista.Cells(linha, COL_STATUS).Value)
        If InStr(1, statusAtual, "Enviado", vbTextCompare) > 0 Or _
           InStr(1, statusAtual, "revisão", vbTextCompare) > 0 Then
            totalPulados = totalPulados + 1
            GoTo ProximaLinha
        End If

        totalProcessados = totalProcessados + 1

        ' --- Le os dados da linha ---
        empresa = Trim(CStr(wsLista.Cells(linha, COL_EMPRESA).Value))
        nomeContato = Trim(CStr(wsLista.Cells(linha, COL_NOME).Value))
        emailPrincipal = Trim(CStr(wsLista.Cells(linha, COL_EMAIL).Value))
        emailCC = Trim(CStr(wsLista.Cells(linha, COL_CC).Value))
        assunto = CStr(wsLista.Cells(linha, COL_ASSUNTO).Value)
        corpo = CStr(wsLista.Cells(linha, COL_CORPO).Value)

        ' --- Validacao: e-mail vazio ---
        If emailPrincipal = "" Then
            statusLinha = "Erro: e-mail principal vazio"
            GoTo RegistrarResultado
        End If

        ' --- Validacao: formato do e-mail principal ---
        If Not ValidarEmails(emailPrincipal) Then
            statusLinha = "Erro: e-mail principal inválido"
            GoTo RegistrarResultado
        End If

        ' --- Validacao: formato do CC (apenas se preenchido) ---
        If emailCC <> "" Then
            If Not ValidarEmails(emailCC) Then
                statusLinha = "Erro: e-mail CC inválido"
                GoTo RegistrarResultado
            End If
        End If

        ' --- Validacao: assunto e corpo ---
        If Trim(assunto) = "" Then
            statusLinha = "Erro: assunto vazio"
            GoTo RegistrarResultado
        End If
        If Trim(corpo) = "" Then
            statusLinha = "Erro: corpo do e-mail vazio"
            GoTo RegistrarResultado
        End If

        ' --- Personalizacao {{EMPRESA}} / {{NOME}} ---
        assunto = AplicarPersonalizacao(assunto, empresa, nomeContato)
        corpo = AplicarPersonalizacao(corpo, empresa, nomeContato)

        ' --- Acrescenta a assinatura padrao ---
        If Trim(assinatura) <> "" Then
            corpo = corpo & vbLf & vbLf & assinatura
        End If

        ' --- Monta o pacote de parametros para o AppleScript ---
        parametros = PrepararTextoParaEnvio(emailPrincipal, False) & DELIMITADOR & _
                     PrepararTextoParaEnvio(emailCC, False) & DELIMITADOR & _
                     PrepararTextoParaEnvio(assunto, False) & DELIMITADOR & _
                     PrepararTextoParaEnvio(corpo, True) & DELIMITADOR & _
                     PrepararTextoParaEnvio(caminhoAnexo, False) & DELIMITADOR & _
                     modoEnvio

        ' --- Chama o Outlook para Mac via AppleScript ---
        resultado = ExecutarEnvioAppleScript(parametros)

        If resultado = "OK" Then
            If modoEnvio = "DIRETO" Then
                statusLinha = "Enviado"
            Else
                statusLinha = "Aberto para revisão"
            End If
        Else
            statusLinha = resultado ' ja vem no formato "ERRO: ..."
        End If

RegistrarResultado:
        ' --- Atualiza Status e Data/Hora na lista ---
        wsLista.Cells(linha, COL_STATUS).Value = statusLinha
        wsLista.Cells(linha, COL_DATAHORA).Value = Format(Now, "dd/mm/yyyy hh:mm:ss")

        ' Cor do status: verde para sucesso, vermelho para erro
        If InStr(1, statusLinha, "Erro", vbTextCompare) > 0 Then
            wsLista.Cells(linha, COL_STATUS).Font.Color = RGB(180, 0, 0)
            totalErros = totalErros + 1
        Else
            wsLista.Cells(linha, COL_STATUS).Font.Color = RGB(0, 130, 0)
            totalEnviados = totalEnviados + 1
        End If

        ' --- Registra no Historico ---
        RegistrarHistorico emailPrincipal, assunto, statusLinha, lote

        ' --- Intervalo entre envios (so apos tentativa real de envio) ---
        If InStr(1, statusLinha, "Erro", vbTextCompare) = 0 And linha < ultimaLinha Then
            Pausar intervaloSegundos
        End If

ProximaLinha:
    Next linha

    ' ================= 3. Relatorio final =================
    MsgBox "Disparo concluído!" & vbNewLine & vbNewLine & _
           "Lote: " & lote & vbNewLine & _
           "Processados nesta execução: " & totalProcessados & vbNewLine & _
           IIf(modoEnvio = "DIRETO", "Enviados: ", "Abertos para revisão: ") & totalEnviados & vbNewLine & _
           "Erros: " & totalErros & vbNewLine & _
           "Pulados (já enviados): " & totalPulados & vbNewLine & vbNewLine & _
           IIf(totalErros > 0, "Verifique a coluna Status e a aba '" & ABA_HISTORICO & "' para detalhes dos erros.", _
               "Detalhes completos na aba '" & ABA_HISTORICO & "'."), _
           IIf(totalErros > 0, vbExclamation, vbInformation), "Relatório final"
End Sub

'-----------------------------------------------------------------------
' ExecutarEnvioAppleScript
' Encapsula a chamada AppleScriptTask com tratamento de erros.
' Retorna "OK" ou "ERRO: <descricao>".
'-----------------------------------------------------------------------
Private Function ExecutarEnvioAppleScript(parametros As String) As String
    Dim retorno As String

    On Error GoTo TratarErro

    retorno = AppleScriptTask(ARQUIVO_APPLESCRIPT, HANDLER_APPLESCRIPT, parametros)

    If Trim(retorno) = "" Then
        ExecutarEnvioAppleScript = "ERRO: o AppleScript não retornou resposta"
    Else
        ExecutarEnvioAppleScript = Trim(retorno)
    End If
    Exit Function

TratarErro:
    ' Erro tipico: o arquivo .applescript nao esta na pasta exigida pelo Excel
    ExecutarEnvioAppleScript = "ERRO: falha ao executar o AppleScript (" & Err.Description & "). " & _
        "Verifique se o arquivo '" & ARQUIVO_APPLESCRIPT & "' está em " & _
        "~/Library/Application Scripts/com.microsoft.Excel/ e se o Excel tem permissão " & _
        "de Automação sobre o Outlook (Ajustes do Sistema > Privacidade e Segurança > Automação)."
End Function

'-----------------------------------------------------------------------
' AbasExistem: confirma que as tres abas necessarias existem
'-----------------------------------------------------------------------
Private Function AbasExistem() As Boolean
    Dim ws As Worksheet
    Dim achouLista As Boolean, achouConfig As Boolean, achouHist As Boolean

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = ABA_LISTA Then achouLista = True
        If ws.Name = ABA_CONFIG Then achouConfig = True
        If ws.Name = ABA_HISTORICO Then achouHist = True
    Next ws

    AbasExistem = achouLista And achouConfig And achouHist
End Function
