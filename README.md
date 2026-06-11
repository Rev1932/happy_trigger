# Disparo de E-mails em Massa — Excel para Mac + Outlook para Mac

Solução 100% macOS para envio de e-mails em massa pelo **Microsoft Outlook para Mac
(aplicativo local)**, controlado por uma planilha **Excel para Mac (Microsoft 365)**.

A integração Excel → Outlook é feita **exclusivamente via AppleScript**
(`AppleScriptTask`), o mecanismo oficial e sancionado pela Microsoft para
automação no Excel para Mac.

**Sem nenhuma dependência de Windows:**

| Restrição | Como foi atendida |
|---|---|
| Sem ActiveX | Botões são formas (Shapes) com macro associada |
| Sem COM Automation / `CreateObject` | Toda a comunicação com o Outlook é via `AppleScriptTask` |
| Sem bibliotecas exclusivas do Windows | Nenhum `Declare`, nenhum `FileSystemObject`, nenhum `Shell` do Windows |
| Verificação de arquivos | `Dir()` nativo do VBA (funciona no macOS) |
| Pausa entre envios | `Timer` + `DoEvents` (sem `Sleep` da API do Windows) |

---

## 1. Arquivos da solução

```
DisparoEmailsMac/
├── README.md                                ← este guia
├── VBA/
│   ├── Mod_Estrutura.bas                    ← cria as 3 abas e os botões
│   ├── Mod_Principal.bas                    ← rotina de disparo em massa
│   └── Mod_Utilitarios.bas                  ← validações, histórico, pausa
└── AppleScript/
    └── EnviarEmailOutlook.applescript       ← integração com o Outlook
```

---

## 2. Instalação passo a passo (no MacBook)

### Passo 1 — Instalar o AppleScript na pasta exigida pelo Excel

Por segurança (sandbox do macOS), o Excel só executa AppleScripts que estejam
na pasta `~/Library/Application Scripts/com.microsoft.Excel/`.

Abra o **Terminal** e execute:

```bash
mkdir -p "$HOME/Library/Application Scripts/com.microsoft.Excel"
cp "EnviarEmailOutlook.applescript" "$HOME/Library/Application Scripts/com.microsoft.Excel/"
```

(Ajuste o caminho de origem do `cp` para onde você salvou o arquivo.)

> Alternativa sem Terminal: abra o **Editor de Script** (Script Editor),
> cole o conteúdo do arquivo, e salve com **Formato de Arquivo: Texto**,
> nome `EnviarEmailOutlook.applescript`, dentro da pasta acima
> (no Finder: `Cmd+Shift+G` → cole o caminho da pasta).

**Importante:** o arquivo deve ser **texto puro** com extensão `.applescript`
(não salve como `.scpt` compilado nem como aplicativo).

### Passo 2 — Criar a pasta de trabalho com macros

1. Abra o Excel e crie uma pasta de trabalho nova.
2. **Arquivo > Salvar Como** → formato **Pasta de Trabalho Habilitada para Macro do Excel (.xlsm)**.
   Sugestão de nome: `Disparo de E-mails.xlsm`.

### Passo 3 — Colar os módulos VBA

1. Menu **Ferramentas > Macro > Editor do Visual Basic** (ou `Fn+Option+F11`).
2. No painel de projeto, clique com o botão direito no projeto da sua pasta de
   trabalho → **Inserir > Módulo**. Repita 3 vezes.
3. Cole em cada módulo o conteúdo de um arquivo `.bas`
   (`Mod_Estrutura`, `Mod_Principal`, `Mod_Utilitarios`).
   *Ao colar, ignore a primeira linha `Attribute VB_Name = ...` se ela vier junto —
   nos arquivos entregues ela não existe, então é só colar tudo.*
4. Opcional: renomeie cada módulo (janela Propriedades, F4) para
   `Mod_Estrutura`, `Mod_Principal` e `Mod_Utilitarios`.

### Passo 4 — Criar a estrutura da planilha

1. Volte ao Excel, menu **Ferramentas > Macro > Macros…**
2. Execute **`CriarEstruturaPlanilha`**.
3. Serão criadas as abas **Lista de Envios**, **Configurações** e **Histórico**,
   com cabeçalhos formatados e os botões **ENVIAR E-MAILS** e **Limpar Status**.

### Passo 5 — Primeiro envio (teste)

1. Preencha a aba **Configurações** (veja seção 4).
2. Preencha 1 linha na **Lista de Envios** com o **seu próprio e-mail**.
3. Deixe o modo de envio em **REVISAO**.
4. Clique em **ENVIAR E-MAILS**.
5. No primeiro disparo, o macOS perguntará se o Excel pode controlar o Outlook —
   clique em **Permitir** (veja seção 3).
6. O e-mail abrirá no Outlook para conferência. Estando tudo certo, mude para
   **DIRETO** quando quiser envio automático.

---

## 3. Permissões do macOS

### 3.1 Automação (obrigatória)

No primeiro envio, aparecerá o aviso:
*"Microsoft Excel deseja controlar o Microsoft Outlook"* → clique em **Permitir**.

Se você negou por engano, ou para conferir depois:

**Ajustes do Sistema > Privacidade e Segurança > Automação >
Microsoft Excel > ativar "Microsoft Outlook"**.

### 3.2 Pasta do AppleScript (obrigatória)

O arquivo `EnviarEmailOutlook.applescript` precisa estar exatamente em:

```
~/Library/Application Scripts/com.microsoft.Excel/
```

Fora dessa pasta o `AppleScriptTask` falha (é uma exigência do sandbox do Excel).

### 3.3 Acesso a arquivos (para anexos)

Quem lê o arquivo do anexo é o **Outlook** (não o Excel). Recomendações:

- Prefira anexos dentro de **Documentos** ou **Downloads**.
- Se o anexo estiver em pasta protegida (Mesa/Desktop, iCloud, discos externos)
  e ocorrer erro de permissão, conceda acesso ao Outlook em
  **Ajustes do Sistema > Privacidade e Segurança > Arquivos e Pastas**
  (ou **Acesso Total ao Disco**, se necessário).

### 3.4 Outlook "Novo" vs. "Legado"

O AppleScript usado (`make new outgoing message`, `send`) é suportado pelo
Outlook para Mac. Se sua versão do **"Novo Outlook"** retornar erro de
AppleScript ao enviar:

1. Atualize o Outlook para a versão mais recente (o suporte a AppleScript no
   Novo Outlook foi ampliado nas versões recentes); **ou**
2. Alterne temporariamente para o Outlook Legado:
   menu **Outlook > desativar "Novo Outlook"** (ou **Ajuda > Voltar ao Outlook Legado**).

---

## 4. Estrutura e preenchimento da planilha

### Aba "Lista de Envios"

| Coluna | Campo | Observações |
|---|---|---|
| A | Empresa | usada no marcador `{{EMPRESA}}` |
| B | Nome do contato | usado no marcador `{{NOME}}` |
| C | E-mail principal | **obrigatório**; vários: separe com `;` |
| D | E-mail CC | opcional; vários: separe com `;` |
| E | Assunto | aceita `{{EMPRESA}}` e `{{NOME}}` |
| F | Corpo do e-mail | aceita `{{EMPRESA}}`, `{{NOME}}` e quebras de linha (`Option+Enter` na célula) |
| G | Status | preenchido automaticamente |
| H | Data/Hora de envio | preenchido automaticamente |

Exemplo de corpo:

```
Olá {{NOME}},

Segue em anexo a proposta comercial preparada para a {{EMPRESA}}.

Fico à disposição para qualquer dúvida.
```

### Aba "Configurações"

| Célula | Campo | Valores |
|---|---|---|
| B3 | Caminho do anexo (opcional) | caminho POSIX completo, ex.: `/Users/victor/Documents/proposta.pdf` — vazio = sem anexo. Aceita PDF, DOCX, XLSX ou qualquer arquivo válido |
| B4 | Modo de envio | `REVISAO` (abre cada e-mail para conferência) ou `DIRETO` (envia imediatamente) — lista suspensa |
| B5 | Intervalo entre envios | segundos de espera entre um envio e outro (ex.: 3) |
| B6 | Assinatura padrão | texto adicionado ao final de todos os e-mails (aceita quebras de linha) |

### Aba "Histórico"

Preenchida automaticamente a cada tentativa de envio:
**Número do envio | Data/Hora | Destinatário | Assunto | Status | Lote de envio**.

Cada execução do disparo gera um identificador de lote único
(ex.: `LOTE-20260611-143055`), permitindo filtrar os envios de cada rodada.

---

## 5. Funcionamento e regras

1. **Envio em massa**: percorre todas as linhas da Lista de Envios.
2. **Retomada de lote**: linhas com status "Enviado" ou "Aberto para revisão"
   são **puladas** — se um lote for interrompido, basta clicar de novo em
   ENVIAR E-MAILS para continuar de onde parou. Para reenviar tudo, use o
   botão **Limpar Status**.
3. **Validações por linha** (linhas inválidas recebem status de erro e o lote continua):
   - e-mail principal vazio;
   - formato inválido do e-mail principal ou do CC;
   - assunto ou corpo vazios.
4. **Anexo**: a existência do arquivo é verificada **antes de iniciar o lote**;
   se não existir, o disparo nem começa.
5. **Tratamento de erros**: qualquer erro do AppleScript/Outlook é capturado e
   gravado na coluna Status (em vermelho) e no Histórico — um erro em uma linha
   não interrompe as demais.
6. **Relatório final**: ao terminar, uma janela mostra o lote, o total
   processado, enviados, erros e pulados.

---

## 6. Explicação de cada módulo

### `Mod_Estrutura.bas`
Construtor da planilha. `CriarEstruturaPlanilha` cria/formata as três abas,
adiciona a lista suspensa do modo de envio e cria os botões como **formas**
(Shapes com `OnAction`) — substituindo o ActiveX, que não existe no Mac.
É seguro executar de novo: não apaga dados já digitados, apenas refaz
cabeçalhos e botões.

### `Mod_Principal.bas`
O coração da solução. `EnviarEmailsEmMassa` lê as configurações, valida o
anexo, pede confirmação, e processa linha a linha: valida, personaliza,
empacota os 6 parâmetros (destinatário, CC, assunto, corpo, anexo, modo) numa
única string separada por `||#||` e chama
`AppleScriptTask("EnviarEmailOutlook.applescript", "enviarEmail", parametros)`.
O retorno (`OK` ou `ERRO: ...`) alimenta o Status, o Histórico e o relatório
final. `ExecutarEnvioAppleScript` isola a chamada com `On Error` e devolve uma
mensagem orientando sobre pasta do script e permissões caso o `AppleScriptTask`
falhe.

### `Mod_Utilitarios.bas`
Funções de apoio, todas nativas do VBA para Mac:
- `ValidarEmails` — valida um ou vários endereços separados por `;`;
- `ArquivoExiste` — checa o anexo com `Dir()`;
- `PrepararTextoParaEnvio` — remove o delimitador do texto e converte quebras
  de linha em `<br>` (o corpo da mensagem no Outlook é HTML);
- `AplicarPersonalizacao` — substitui `{{EMPRESA}}` e `{{NOME}}`;
- `Pausar` — intervalo entre envios com `Timer`/`DoEvents`;
- `RegistrarHistorico` — grava cada tentativa na aba Histórico;
- `LimparStatus` — limpa Status/Data-Hora para permitir reenvio.

### `EnviarEmailOutlook.applescript`
Único ponto de contato com o Outlook. O handler `enviarEmail` desmonta a
string de parâmetros, cria a mensagem (`make new outgoing message`), adiciona
destinatários (`recipient` / `cc recipient`, com suporte a vários endereços),
anexa o arquivo (`make new attachment` com `POSIX file`) e então **envia**
(`send`) ou **abre para revisão** (`open` + `activate`), conforme o modo.
Todo o bloco roda dentro de `try`, devolvendo `OK` ou `ERRO: <descrição>`
para o VBA.

---

## 7. Solução de problemas rápidos

| Sintoma | Causa provável | Correção |
|---|---|---|
| "ERRO: falha ao executar o AppleScript" | Arquivo `.applescript` fora da pasta exigida ou com nome diferente | Refaça o Passo 1 da instalação |
| Erro -1743 ("not authorized") | Permissão de Automação negada | Ajustes do Sistema > Privacidade e Segurança > Automação |
| Erro ao anexar arquivo | Caminho errado ou pasta protegida | Use caminho POSIX completo; mova o anexo para Documentos |
| Erro de AppleScript só no "Novo Outlook" | Versão sem suporte completo a AppleScript | Atualize o Outlook ou alterne para o modo Legado |
| Macro não aparece / botão não funciona | Arquivo salvo como .xlsx | Salve como **.xlsm** e reabra habilitando macros |
