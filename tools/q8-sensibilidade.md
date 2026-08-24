# Q8 — verificar a SENSIBILIDADE do token da janela

**Quem faz:** você, Ricardo. Leva uns 3 minutos e **não** exige esperar
download nenhum.

## Por que isto está pendente

O Iris identifica o ambiente por uma impressão digital, e a janela de
sincronização faz parte dela — porque a janela muda **o que existe**, não só o
que custa (§18.4). Como o OOM não expõe a janela (§22.3), a impressão digital
usa o valor cru `00036601` do perfil como token opaco.

Um token opaco serve se tiver duas propriedades:

| Propriedade | O que significa | Estado |
|---|---|---|
| **Estável** | não muda sozinho | **parcial** — §22.4 mediu 5 leituras idênticas em ~1,2 s, na mesma sessão. Reinício não foi coberto. |
| **Sensível** | muda quando a janela muda | **NÃO MEDIDO** |

Se ele não for sensível, todo o mecanismo é decoração: o Iris não distinguiria
o universo de antes do de depois, e as autorizações concedidas ao antigo
continuariam valendo no novo sem ninguém notar.

Eu tinha suposto que medir isto custava GB de download. Custa não: o que custa
GB é medir o **universo resultante**. Ler o token é instantâneo.

## O protocolo

Rode o script antes de cada passo e anote a linha do token:

```bash
powershell -NoProfile -File tools/q8-token.ps1
```

O valor que importa é o da chave **da conta**
(`Outlook\<guid>`), **não** o do `GroupsStore` — são dois valores diferentes,
e hoje eles são `84-09-00-00` e `80-01-00-00`.

1. **Leia** o token. Anote. (Hoje: `84-09-00-00`.)
2. No Outlook: *Arquivo → Configurações de Conta → Configurações de Conta →*
   dê duplo clique na conta → mova o cursor **"Baixar email do passado"** para
   um valor **diferente** → OK.
3. **Leia** o token de novo. Anote.
4. Volte o cursor para **onde estava**. OK.
5. **Leia** o token uma terceira vez. Anote.
6. Feche e reabra o Outlook, **leia** de novo.
7. Se der, reinicie o computador e **leia** uma última vez — este passo é o
   que fecha a estabilidade que a §22.4 deixou pela metade.

Se quiser evitar o download, pode cancelar/pausar a sincronização depois do
passo 2. **Eu não sei em que momento o registro é escrito** — pode ser no OK,
ao fechar o diálogo, no reinício, ou em outro ponto. É justamente isso que o
protocolo vai revelar: leia depois de cada passo e veja quando o valor muda.

## O que cada resultado significa

- **Passo 3 diferente do passo 1, e passo 5 igual ao passo 1** → o token é
  sensível e reversível. É o resultado esperado. Marque `TokenValidado:=True`
  na linha da matriz em `src/Iris.Sync/EnvironmentPolicy.vb` **desde que** o
  passo 7 também tenha dado igual ao 5, e registre todos os valores na §22.4.

  Se o passo 7 não for feito, **não marque**: `TokenValidado` exige as duas
  propriedades — sensibilidade e estabilidade através de reinício —, e a §22.4
  só mediu leituras numa sessão curta. Um token que mude sozinho ao reiniciar
  faria o Iris reconciliar a caixa inteira toda vez que você abrisse o Outlook.

- **Passo 3 igual ao passo 1 em todas as leituras seguintes** → o token não é
  sensível **onde foi lido**. O mecanismo não serve como está e a impressão
  digital precisa de outra fonte. Uma alternativa **a investigar** — não
  demonstrada, e com problema conhecido — é usar o alcance (a mensagem mais
  antiga alcançável) como parte da identidade: é caro, e pior, o alcance muda
  sem a configuração mudar, porque correio novo chega e a mensagem mais antiga
  pode ser excluída ou movida. Não marque nada; me avise.

- **Passo 6 diferente do passo 5** → o token muda com reinício, ou seja, não é
  estável do jeito que a §22.4 mediu, e a medição de estabilidade precisa ser
  refeita cobrindo reinício. Também me avise.

## Por que eu não fiz isto sozinho

Mover o cursor da janela é mudança de configuração da sua conta e dispara
download. Está na lista do que eu não faço sem você pedir.
