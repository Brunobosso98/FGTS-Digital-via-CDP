# Automação com Chrome + CDP (Playwright)

## Objetivo
Este projeto usa o Playwright para controlar uma janela do Chrome que foi aberta manualmente com depuração remota (CDP).

## Como funciona
1. O Chrome é iniciado com `--remote-debugging-port=9222`.
2. Isso expõe uma API local (`http://127.0.0.1:9222`) com informações e controle das abas.
3. O Playwright conecta nessa instância com `connectOverCDP(...)`.
4. Após conectado, ele usa a aba aberta para clicar/preencher/ler HTML.

## Comando para abrir o Chrome (Windows)
No CMD:

```bat
"%ProgramFiles%\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="%TEMP%\chrome-cdp"
```

Alternativa (se não existir no caminho acima):

```bat
"%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="%TEMP%\chrome-cdp"
```

## Como validar se o CDP está ativo
Com o Chrome aberto nesse modo, acesse:

- `http://127.0.0.1:9222/json/version`
- `http://127.0.0.1:9222/json/list`

Se aparecer JSON com `webSocketDebuggerUrl`, a conexão está pronta.

## Fluxo usado neste projeto
Arquivo: `index.js`

- Conecta ao CDP em `http://127.0.0.1:9222`
- Obtém o primeiro contexto e a primeira aba
- Traz a aba para frente
- Aguarda o botão `Trocar Perfil`
- Clica no botão

## Execução
No diretório do projeto:

```bash
node index.js
```

## Observações importantes
- Esse método controla a mesma sessão da janela aberta manualmente.
- Se fechar o Chrome iniciado com CDP, a automação perde a conexão.
- A porta de depuração dá controle local do navegador; use apenas em ambiente confiável.
