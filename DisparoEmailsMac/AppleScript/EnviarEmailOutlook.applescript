-- =====================================================================
--  EnviarEmailOutlook.applescript
--  Integracao Excel para Mac  ->  Microsoft Outlook para Mac
--
--  Este arquivo DEVE ser salvo (como texto puro, extensao .applescript) em:
--      ~/Library/Application Scripts/com.microsoft.Excel/
--
--  Ele e chamado pelo VBA atraves de:
--      AppleScriptTask("EnviarEmailOutlook.applescript", "enviarEmail", parametros)
--
--  O VBA envia UMA string com os campos separados pelo delimitador "||#||",
--  na seguinte ordem:
--      1. Destinatario(s)  (varios separados por ";")
--      2. CC               (vazio ou varios separados por ";")
--      3. Assunto
--      4. Corpo (HTML - quebras de linha ja convertidas em <br>)
--      5. Caminho POSIX do anexo (vazio = sem anexo)
--      6. Modo de envio: "DIRETO" (envia) ou "REVISAO" (abre rascunho)
--
--  Retorno para o VBA:
--      "OK"          -> sucesso
--      "ERRO: ..."   -> descricao do erro (o VBA registra no Status/Historico)
-- =====================================================================

-- Handler principal chamado pelo Excel (AppleScriptTask)
on enviarEmail(paramString)
	try
		-- ----- 1. Separa os parametros recebidos do VBA -----
		set delimitadorAntigo to AppleScript's text item delimiters
		set AppleScript's text item delimiters to "||#||"
		set listaParametros to text items of paramString
		set AppleScript's text item delimiters to delimitadorAntigo

		if (count of listaParametros) < 6 then
			return "ERRO: parametros insuficientes recebidos do Excel"
		end if

		set textoDestinatarios to item 1 of listaParametros
		set textoCC to item 2 of listaParametros
		set assuntoEmail to item 3 of listaParametros
		set corpoEmail to item 4 of listaParametros
		set caminhoAnexo to item 5 of listaParametros
		set modoEnvio to item 6 of listaParametros

		-- ----- 2. Valida destinatario -----
		if textoDestinatarios is "" then
			return "ERRO: destinatario vazio"
		end if

		-- ----- 3. Cria e configura a mensagem no Outlook -----
		tell application "Microsoft Outlook"
			-- Cria a mensagem (content aceita HTML, por isso o corpo usa <br>)
			set novaMensagem to make new outgoing message with properties {subject:assuntoEmail, content:corpoEmail}

			-- Destinatarios principais (suporta varios, separados por ";")
			repeat with enderecoAtual in my separarEmails(textoDestinatarios)
				make new recipient at novaMensagem with properties {email address:{address:(contents of enderecoAtual)}}
			end repeat

			-- Destinatarios em copia (CC), se informados
			if textoCC is not "" then
				repeat with enderecoAtual in my separarEmails(textoCC)
					make new cc recipient at novaMensagem with properties {email address:{address:(contents of enderecoAtual)}}
				end repeat
			end if

			-- Anexo (caminho POSIX), se informado
			if caminhoAnexo is not "" then
				make new attachment at novaMensagem with properties {file:(POSIX file caminhoAnexo)}
			end if

			-- ----- 4. Envia ou abre para revisao -----
			if modoEnvio is "DIRETO" then
				send novaMensagem
			else
				-- Modo REVISAO: abre a janela do e-mail para conferencia manual
				open novaMensagem
				activate
			end if
		end tell

		return "OK"

	on error mensagemErro number numeroErro
		return "ERRO: " & mensagemErro & " (codigo " & numeroErro & ")"
	end try
end enviarEmail

-- ---------------------------------------------------------------------
-- separarEmails: recebe "a@x.com; b@y.com" e devolve lista de enderecos
-- limpos (sem espacos e sem itens vazios)
-- ---------------------------------------------------------------------
on separarEmails(textoEmails)
	set delimitadorAntigo to AppleScript's text item delimiters
	set AppleScript's text item delimiters to ";"
	set partes to text items of textoEmails
	set AppleScript's text item delimiters to delimitadorAntigo

	set listaLimpa to {}
	repeat with parteAtual in partes
		set enderecoLimpo to my removerEspacos(contents of parteAtual)
		if enderecoLimpo is not "" then
			set end of listaLimpa to enderecoLimpo
		end if
	end repeat
	return listaLimpa
end separarEmails

-- ---------------------------------------------------------------------
-- removerEspacos: remove espacos no inicio e no fim de um texto
-- ---------------------------------------------------------------------
on removerEspacos(textoOriginal)
	set textoAtual to textoOriginal
	repeat while textoAtual begins with " "
		if (length of textoAtual) is 1 then return ""
		set textoAtual to text 2 thru -1 of textoAtual
	end repeat
	repeat while textoAtual ends with " "
		if (length of textoAtual) is 1 then return ""
		set textoAtual to text 1 thru -2 of textoAtual
	end repeat
	return textoAtual
end removerEspacos
