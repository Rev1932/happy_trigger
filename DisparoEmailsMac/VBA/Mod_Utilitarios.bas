'=======================================================================
' MOD_UTILITARIOS
' Funcoes de apoio: validacao, arquivos, texto, pausa e historico.
'
' 100% compativel com Excel para Mac:
'   - Verificacao de arquivo com Dir() nativo do VBA (sem FileSystemObject)
'   - Pausa com Timer/DoEvents (sem APIs do Windows)
'=======================================================================
Option Explicit

' Delimitador usado para empacotar os parametros enviados ao AppleScript.
' Deve ser identico ao usado no arquivo EnviarEmailOutlook.applescript.
Public Const DELIMITADOR As String = "||#||"

'-----------------------------------------------------------------------
' ValidarEmails
' Valida um campo que pode conter um ou mais e-mails separados por ";".
' Retorna True somente se TODOS os enderecos tiverem formato valido.
'-----------------------------------------------------------------------
Public Function ValidarEmails(campoEmail As String) As Boolean
    Dim partes() As String
    Dim i As Long
    Dim endereco As String

    If Trim(campoEmail) = "" Then
        ValidarEmails = False
        Exit Function
    End If

    partes = Split(campoEmail, ";")
    For i = LBound(partes) To UBound(partes)
        endereco = Trim(partes(i))
        ' Ignora itens vazios gerados por ";" sobrando no fim
        If endereco <> "" Then
            ' Verificacao basica de formato: algo@algo.algo, sem espacos
            If Not (endereco Like "?*@?*.?*") Or InStr(endereco, " ") > 0 Then
                ValidarEmails = False
                Exit Function
            End If
        End If
    Next i

    ValidarEmails = True
End Function

'-----------------------------------------------------------------------
' ArquivoExiste
' Verifica se um arquivo existe no caminho informado (caminho POSIX,
' ex.: /Users/seunome/Documents/proposta.pdf). Usa Dir(), nativo do
' VBA e funcional no macOS.
'-----------------------------------------------------------------------
Public Function ArquivoExiste(caminho As String) As Boolean
    On Error Resume Next
    ArquivoExiste = (Dir(caminho) <> "")
    On Error GoTo 0
End Function

'-----------------------------------------------------------------------
' PrepararTextoParaEnvio
' Prepara um campo de texto antes de enviar ao AppleScript:
'   1. Remove o delimitador, caso apareca no texto (evita quebra do pacote)
'   2. Converte quebras de linha em <br> (o corpo da mensagem e HTML)
'-----------------------------------------------------------------------
Public Function PrepararTextoParaEnvio(texto As String, converterQuebras As Boolean) As String
    Dim resultado As String
    resultado = Replace(texto, DELIMITADOR, " ")

    If converterQuebras Then
        resultado = Replace(resultado, vbCrLf, "<br>")
        resultado = Replace(resultado, vbCr, "<br>")
        resultado = Replace(resultado, vbLf, "<br>")
    Else
        ' Campos de linha unica (assunto, e-mails): remove quebras
        resultado = Replace(resultado, vbCrLf, " ")
        resultado = Replace(resultado, vbCr, " ")
        resultado = Replace(resultado, vbLf, " ")
    End If

    PrepararTextoParaEnvio = Trim(resultado)
End Function

'-----------------------------------------------------------------------
' AplicarPersonalizacao
' Substitui os campos de personalizacao pelo conteudo da linha:
'   {{EMPRESA}} -> coluna Empresa
'   {{NOME}}    -> coluna Nome do contato
' A substituicao ignora maiusculas/minusculas no marcador.
'-----------------------------------------------------------------------
Public Function AplicarPersonalizacao(texto As String, empresa As String, nomeContato As String) As String
    Dim resultado As String
    resultado = texto
    resultado = Replace(resultado, "{{EMPRESA}}", empresa, 1, -1, vbTextCompare)
    resultado = Replace(resultado, "{{NOME}}", nomeContato, 1, -1, vbTextCompare)
    AplicarPersonalizacao = resultado
End Function

'-----------------------------------------------------------------------
' Pausar
' Aguarda o numero de segundos informado sem travar o Excel.
' Usa Timer + DoEvents (sem Sleep da API do Windows).
'-----------------------------------------------------------------------
Public Sub Pausar(segundos As Double)
    Dim inicio As Double

    If segundos <= 0 Then Exit Sub

    inicio = Timer
    Do While Timer < inicio + segundos
        ' Se passar da meia-noite, Timer reinicia; encerra a pausa
        If Timer < inicio Then Exit Do
        DoEvents
    Loop
End Sub

'-----------------------------------------------------------------------
' RegistrarHistorico
' Acrescenta uma linha na aba Historico com os dados do envio.
'-----------------------------------------------------------------------
Public Sub RegistrarHistorico(destinatario As String, assunto As String, _
                              statusEnvio As String, lote As String)
    Dim wsHist As Worksheet
    Dim proximaLinha As Long

    Set wsHist = Worksheets(ABA_HISTORICO)
    proximaLinha = wsHist.Cells(wsHist.Rows.Count, "A").End(xlUp).Row + 1

    With wsHist
        .Cells(proximaLinha, 1).Value = proximaLinha - 1          ' Numero sequencial do envio
        .Cells(proximaLinha, 2).Value = Format(Now, "dd/mm/yyyy hh:mm:ss")
        .Cells(proximaLinha, 3).Value = destinatario
        .Cells(proximaLinha, 4).Value = assunto
        .Cells(proximaLinha, 5).Value = statusEnvio
        .Cells(proximaLinha, 6).Value = lote
    End With
End Sub

'-----------------------------------------------------------------------
' LimparStatus
' Limpa as colunas Status e Data/Hora da aba Lista de Envios,
' permitindo reenviar a lista (o Historico e preservado).
'-----------------------------------------------------------------------
Public Sub LimparStatus()
    Dim wsLista As Worksheet
    Dim ultimaLinha As Long

    Set wsLista = Worksheets(ABA_LISTA)
    ultimaLinha = wsLista.Cells(wsLista.Rows.Count, "C").End(xlUp).Row

    If ultimaLinha >= 2 Then
        If MsgBox("Limpar o Status e a Data/Hora de todas as linhas?" & vbNewLine & _
                  "(O Histórico será preservado.)", vbYesNo + vbQuestion, _
                  "Limpar Status") = vbYes Then
            wsLista.Range("G2:H" & ultimaLinha).ClearContents
        End If
    End If
End Sub
