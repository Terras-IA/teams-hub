# Teams Hub

Várias contas do **Microsoft Teams** e do **Outlook** abertas ao mesmo tempo,
numa janela só, no Linux. Feito pela [TerrasIA](https://www.terrasia.app).

O Teams para Linux não deixa ficar logado em mais de uma conta. O Teams Hub
abre cada conta numa sessão separada, então todas ficam conectadas juntas e
continuam logadas depois de reiniciar.

## O que faz

- **Uma aba por conta do Teams**, cada uma com o próprio login.
- **Outlook de cada conta** (✉ embaixo da conta), na mesma sessão do Teams,
  sem pedir login de novo.
- **Quem chamou:** o número de não lidas aparece em cada aba, e o ícone da
  bandeja ganha um ponto vermelho quando chega mensagem no Teams.
- **Avisos com o nome da conta**, por exemplo `[Empresa ABC] Fulano: oi`.
  Clicar abre a aba certa.
- **Compartilhar tela** em reunião, com seletor de tela ou janela.
- **Fica na bandeja:** fechar a janela só esconde; os avisos continuam.
  Para sair, use o menu da bandeja → **Sair**.
- **Contas pela interface:** adicione, renomeie, troque a cor ou remova na
  aba ⚙ **Contas**.

## Instalar

Ubuntu, Debian, Linux Mint e derivados, 64 bits.

1. Baixe o `.deb` mais recente em [Releases](https://github.com/Terras-IA/teams-hub/releases).
2. Instale com dois cliques no arquivo, ou pelo terminal:

   ```bash
   sudo apt install ./teams-hub_*_amd64.deb
   ```

3. Abra **Teams Hub** no menu de aplicativos. Na aba **Contas**, preencha
   sigla, nome e cor e clique em **Adicionar** para cada conta. Depois entre
   com a conta da Microsoft na aba que aparecer.

Para remover: `sudo apt remove teams-hub`.

## Dicas

- **Teams travado** (botão que não responde depois de horas aberto): clique
  com o botão direito na conta, na barra lateral, e escolha recarregar.
- **Aviso de e-mail novo:** ligue as notificações no Outlook de cada conta
  (⚙ → Geral → Notificações).
- **Problema?** O app grava o que deu errado em
  `~/.cache/teams-hub/diagnostico.log`. Abra uma
  [issue](https://github.com/Terras-IA/teams-hub/issues) com o trecho do log.

## Privacidade

O Teams Hub não tem servidor. Ele só abre o Teams e o Outlook da Microsoft
dentro dele, como um navegador. Logins e configurações ficam no seu
computador (`~/.config/teams-hub` e `~/.config/Teams Hub`).

O Teams Hub não é um produto da Microsoft nem é afiliado a ela. Microsoft
Teams e Outlook são marcas da Microsoft.

## Gerar a partir do código

Requer Node.js 20 ou mais novo.

```bash
npm install
npm start          # roda em modo de desenvolvimento
npm run dist       # gera dist/teams-hub_<versão>_amd64.deb
```

## Licença

[MIT](LICENSE) © 2026 TerrasIA
