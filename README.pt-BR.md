# Claude para Calc (extensão do LibreOffice)

Painel lateral de chat dentro do Calc, no estilo do "Claude para Excel": você pergunta sobre a planilha
aberta e o Claude lê células, explica fórmulas, escreve/corrige fórmulas, formata, ordena e cria gráficos
e tabelas dinâmicas. Cada alteração é uma ação normal do LibreOffice e pode ser desfeita com **Ctrl+Z**.

## Instalação (Windows 11 + LibreOffice 26.2)
1. No Calc: **Ferramentas ▸ Gerenciador de Extensões ▸ Adicionar…** e escolha `ClaudeCalc.oxt`.
2. Reinicie o LibreOffice.
3. Abra uma planilha e use o menu **Claude ▸ Abrir painel do Claude** (ou o ícone laranja na barra lateral direita).

## Como funciona a autenticação
A extensão **não lê nem guarda credenciais**. Ela executa o **Claude Code** (`claude -p`) que já está no seu
computador e que usa o login que você já tem. Ordem de detecção do executável:
configurado em Config. → `claude` no PATH → `%USERPROFILE%\.local\bin\claude.exe` → `%APPDATA%\npm\claude.cmd`
→ Claude Code embutido no app Claude Desktop (`%APPDATA%\Claude\claude-code\<versão>\claude.exe`, inclusive a
variante MSIX em `%LOCALAPPDATA%\Packages\Claude_*\LocalCache\Roaming\...`).
O executável do próprio app Claude Desktop (WindowsApps\Claude.exe) é ignorado de propósito.

- Se o painel disser **"Não conectado"**, clique em **Entrar**: abre o login do próprio Claude no navegador.
- Alternativa sem login de assinatura: defina a variável de ambiente `ANTHROPIC_API_KEY` (chave do Console).

## O que o Claude pode fazer (ferramentas)
Ler: `get_workbook_info`, `read_range`, `find_cells`, `find_errors`, `get_format`.
Alterar: `write_range` (valores e fórmulas estilo Excel), `clear_range`, `format_range`, `copy_range`,
`insert_delete`, `resize`, `sort_range`, `manage_sheet`, `freeze_panes`, `add_comment`, `define_name`,
`create_chart`, `create_pivot_table`. Mostrar: `select_range` (seleciona a célula na tela).
O Claude só enxerga estas ferramentas: **sem terminal, sem arquivos, sem internet**.

Exemplos: "Explique como chegamos ao total em F20" · "Ache as células com erro e corrija" ·
"Monte um resumo por região em outra aba, com fórmulas" · "Formate o cabeçalho e congele a primeira linha" ·
"Crie um gráfico de colunas das vendas por mês".

## Segurança e privacidade
- As ferramentas são servidas por um servidor local em `127.0.0.1` (porta aleatória), protegido por token
  por painel e conferência do cabeçalho Host. Nada fica exposto na rede.
- O conteúdo das células que o Claude **lê** é enviado à Anthropic, como no Claude para Excel.
- **Config. ▸ "Pedir confirmação antes de cada alteração"** mostra uma caixa Sim/Não antes de cada escrita.

## Limitações conhecidas
- Chat em texto simples (sem links clicáveis para células; o Claude usa `select_range` para apontá-las).
- Ainda não há formatação condicional, validação de dados nem macros.
- Funções novas do Excel dependem de existirem no LibreOffice (se der `#NAME?`, o Claude troca por outra).
- Leitura limitada a 5000 células por chamada (ajustável em Config.); o Claude lê em partes.

## Solução de problemas
- Log: `%APPDATA%\ClaudeCalc\claudecalc.log`. Configurações: `%APPDATA%\ClaudeCalc\settings.json`.
- "Claude Code não encontrado": Config. ▸ informe o caminho do `claude.exe` ▸ Detectar.
- Erro de login: clique em Entrar ou rode `claude auth login` em um terminal.
- Windows sem Git for Windows pode impedir o Claude Code de iniciar; instale-o se aparecer erro sobre git-bash.

## Código-fonte
O `.oxt` é um zip: `python/claude_calc.py` (componente UNO) e `python/claudecalc/` (tools.py, panel.py,
mcp_server.py, claude_cli.py, executor.py…). Para recompilar: `zip -r ClaudeCalc.oxt META-INF description
description.xml images python *.xcu`.

## Desenvolvimento
- `./build.sh` gera `dist/ClaudeCalc.oxt` a partir de `src/`.
- `tests/test_tools.py` testa as ferramentas contra um LibreOffice real (`soffice --headless --accept=socket,host=127.0.0.1,port=2002;urp;`).
- Aviso: projeto da comunidade, não é um produto oficial da Anthropic.
