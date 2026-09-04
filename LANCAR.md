# Lançar uma versão do Iris

Como uma versão sai daqui e chega na outra máquina. Escrito para ser lido
daqui a seis meses, quando ninguém lembrar por que as coisas estão nesta
ordem.

## Por que não há login

A pergunta que originou tudo isto era se seria preciso um sistema de login
para distribuir versões. Não é, e a razão importa: **login autentica quem
baixa**. O que precisa ser garantido é o contrário — que o pacote veio de
quem diz ter vindo e não foi trocado no caminho. Isso é **assinatura**.

E aqui é sério de um jeito que numa atualização comum não é: o Iris lê o
e-mail do dono. Um atualizador é um canal de execução de código apontado
para *dentro* desse programa. Ele merece o mesmo rigor que o portão de
divulgação tem para o que sai.

## Uma vez, e nunca mais

```powershell
.\tools\gerar-chave-de-assinatura.ps1
```

Gera o par ECDSA P-256. A privada vai para `%USERPROFILE%\.iris\`, com ACL
só para você — o arquivo nasce vazio, já restrito, e só depois recebe conteúdo.
A pública é impressa na tela.

**A pasta importa tanto quanto o arquivo**, e por um motivo que não é óbvio: a
ACL de um arquivo não protege o *nome* dele dentro da pasta. Quem pode criar e
apagar entradas ali troca o arquivo depois que ele recebeu a ACL, e a chave
acaba escrita num objeto de outra pessoa.

Por isso o script **recusa** uma pasta onde alguma identidade que ele consegue
resolver — um usuário, um grupo local ou de domínio — possa escrever, e nomeia
quais são. Ele **avisa**, sem recusar, quando a identidade não resolve para
conta nenhuma: perfis do Windows costumam carregar SIDs órfãos de instalações
anteriores, que ninguém pode usar para entrar. A ressalva é que um SID de
domínio também deixa de resolver com o controlador fora do ar — nesse caso vira
aviso, e a decisão é sua.

`-AceitarPastaCompartilhada` passa por cima da recusa. Existe porque uma
conferência sem saída é uma conferência que alguém apaga.

A criptografia não acontece no PowerShell, e sim em
[`tools/Iris.Assinatura`](tools/Iris.Assinatura), que é .NET 10. Os dois
scripts faziam tudo em PowerShell na primeira versão e **não rodavam**:
`powershell.exe` nesta máquina é o Windows PowerShell 5.1, sobre .NET
Framework, onde `ExportPkcs8PrivateKeyPem` e `ImportFromPem` não existem, e
não há `pwsh` instalado. A descoberta teria sido na hora de gerar a chave.

Se você perder a chave pública mas ainda tiver a privada, recupere-a sem gerar
par novo:

```powershell
.\tools\publicar-versao.ps1 -MostrarChavePublica
```

Cole a pública em [`ChaveDeAtualizacao.vb`](src/Iris.App/ChaveDeAtualizacao.vb),
junto com o endereço do manifesto:

```
https://github.com/<dono>/<repo>/releases/latest/download/iris.json
```

`latest/download` é um endereço estável: não é preciso saber o número da
última versão para perguntar qual é a última versão.

Com uma ressalva que vale saber: *latest* é um **metadado do GitHub**, e não
"a maior versão". Rascunhos e pré-lançamentos ficam de fora, e uma release
pode ser marcada como *latest* à mão. Por isso o script publica com `--latest`
explícito — deixar implícito seria deixar a conta certa para o acaso.

**Guarde uma cópia da chave privada fora desta máquina.** Perdê-la não
quebra o Iris que já está instalado — quebra a sua capacidade de publicar
qualquer atualização que ele aceite. O conserto seria distribuir à mão uma
versão nova, com outra chave pública embutida.

Enquanto `ChaveDeAtualizacao.vb` estiver vazio, a tela diz *"a verificação
de versões ainda não foi configurada"* — e não *"a assinatura não
confere"*, que seria acusar um ataque onde houve uma pendência.

## A cada versão

```powershell
.\tools\publicar-versao.ps1 -Versao 0.2.0 -Repositorio dono/repo -Notas "O que mudou."
```

O script publica autocontido num `.exe` só, calcula o SHA-256 **do arquivo
que acabou de sair**, escreve `iris.json` com esse hash, assina o `iris.json`
inteiro — e **só então** grava a versão em `Directory.Build.props`.

Essa ordem é o conserto de um defeito real: a versão era gravada primeiro, e
uma falha em qualquer etapa seguinte a deixava gravada. A reexecução com o
mesmo número era então recusada pela própria conferência de "a versão tem de
subir", e sair daquilo exigia editar o arquivo à mão. Agora a versão vai para
o build por `-p:Version=`, e o arquivo só muda **no fim de tudo** — depois do
`gh release create`, quando ele acontece, porque o ponto de não-retorno é a
release existir. Sem `-Publicar`, o script diz como desfazer.

Ele **não publica**. Deixa os arquivos em `artefatos/<versao>/` e imprime o
comando `gh release create`. Subir a release é um ato seu, e um ato público.
Com `-Publicar`, ele sobe.

Depois: **commite o `Directory.Build.props`**. O número da versão vive lá,
e só lá.

### A versão tem de subir

O Iris compara com `<=` e trata "igual" como já estar em dia. Republicar o
mesmo número produz uma release que ninguém baixa, sem erro em lugar
nenhum — a descoberta seria alguém reclamando que a atualização não chega.
O script recusa antes de compilar.

E "maior" não é capricho: um manifesto antigo, legitimamente assinado, pode
ser reapresentado por quem controle o caminho da rede. Se "diferente"
bastasse, isso rebaixaria a instalação para uma versão com um defeito já
corrigido — assinada, e portanto aceita.

## O que o Iris faz do outro lado

"Verificar atualizações" **nunca instala sozinho**. Ele baixa, confere a
assinatura e o hash, e diz onde o arquivo ficou. O clique duplo é seu.

Trocar um executável que o Windows tem aberto exige um instante em que não
existe Iris nenhum no disco; dá para fazer direito, com um segundo
executável que faz a troca, e é bem mais máquina do que duas máquinas
justificam.

E ele **não pergunta sozinho**: não há verificação no arranque nem
temporizador. Um programa que fala com um servidor sem ninguém pedir
anuncia, a cada abertura, que aquela máquina existe.

## O que a assinatura não compra

Ela prova **origem**, e não **qualidade**: uma versão ruim assinada por você
é uma versão ruim que o Iris aceita.

E não é o certificado de código do Windows. Um executável sem assinatura
Authenticode, baixado da Internet, **pode** provocar o aviso do SmartScreen —
"aplicativo não reconhecido", com um "Mais informações" → "Executar assim
mesmo". Digo *pode* e não *vai*: depende da marca de origem que o navegador
gravou no arquivo, da política do Windows naquela máquina e da reputação que
aquele binário ou aquele editor já tenha. Cada versão nova começa sem
reputação nenhuma.

Um certificado de assinatura de código comum já dá identidade ao editor e
permite acumular reputação; o EV historicamente recebia tratamento melhor.
Nenhum dos dois é garantia de silêncio, e os dois têm custo anual. Para duas
máquinas suas, não vale.

## O primeiro push, e o que ele deixou resolvido

**Feito em 03/09/2026.** O repositório é
[github.com/fricardo-dot/iris](https://github.com/fricardo-dot/iris), público —
e público é requisito, não preferência: o Iris não manda credencial nenhuma,
então `latest/download` só responde em repositório aberto.

Duas coisas foram resolvidas antes de ele existir, e as duas ficariam caras
depois:

1. **O e-mail de trabalho saiu do histórico.** Ele estava em dois arquivos — o
   caminho fixo de um `.ost` num script, e um log colado num documento; o nome
   de um `.ost` *é* o endereço da conta. Os arquivos foram limpos e o histórico
   reescrito, com a árvore do topo saindo idêntica.
2. **O autor dos commits passou a ser o `noreply` do GitHub.** Eles estavam com
   um endereço fictício, que funcionava e não associava nada à conta. Agora são
   `294185678+fricardo-dot@users.noreply.github.com` — o GitHub reconhece o
   autor, e nenhum endereço real fica público. A configuração ficou no
   repositório, e não na global, para os commits seguintes não voltarem ao
   antigo.

### E o ciclo foi conferido contra a release de verdade

Na 0.2.0, com o código do cliente e a chave embutida no executável — não com
dublê:

| Versão instalada | Desfecho |
|---|---|
| 0.1.0 | há versão nova, oferece a 0.2.0 |
| 0.2.0 | já está em dia |
| 9.9.9 | já está em dia — **o rebaixamento é recusado** |

Depois o download inteiro: 66.422.849 bytes, SHA-256 igual ao do manifesto
assinado, promovido com o nome final, nenhum `.parcial` sobrando. Os três
hashes conferem — o do manifesto, o do arquivo compilado e o do baixado.

**Duas coisas só aconteceram aí, e nenhum teste as tinha exercitado:** um
redirecionamento HTTPS de verdade — o `latest/download` redireciona para
`objects.githubusercontent.com` —, e a assinatura atravessando a rede. Até
então, "as duas pontas usam o mesmo formato" era prova de laboratório, com
assinador e leitor no mesmo processo.

### O que ainda falta

**Executar o `Iris-0.2.0.exe` na segunda máquina.** É o único passo que fecha a
distribuição de fato: um pacote baixado e executado por quem não o compilou. O
Windows vai avisar — *"aplicativo não reconhecido"*, com **Mais informações →
Executar assim mesmo** —, pelo motivo explicado acima, na seção do que a
assinatura não compra.

Lá, o botão de versões deve dizer *"Você já está na versão 0.2.0"*.

## Onde as peças moram

| O quê | Onde |
|---|---|
| Gerar a chave (você roda) | [`gerar-chave-de-assinatura.ps1`](tools/gerar-chave-de-assinatura.ps1) |
| Publicar uma versão (você roda) | [`publicar-versao.ps1`](tools/publicar-versao.ps1) |
| Escrever o `iris.json` | [`montar-manifesto.ps1`](tools/montar-manifesto.ps1) |
| Compilar o assinador | [`montar-assinador.ps1`](tools/montar-assinador.ps1) |
| O número da versão | [`Directory.Build.props`](Directory.Build.props) |
| A chave pública e o endereço | [`ChaveDeAtualizacao.vb`](src/Iris.App/ChaveDeAtualizacao.vb) |
| Ler e conferir o manifesto | [`ManifestoDeVersao.vb`](src/Iris.Update/ManifestoDeVersao.vb) |
| Perguntar e baixar | [`ProcuraDeVersao.vb`](src/Iris.Update/ProcuraDeVersao.vb) |
| A tela | [`AtualizacaoViewModel.vb`](src/Iris.App/ViewModels/AtualizacaoViewModel.vb) |
| Os testes, com as sabotagens | [`AtualizacaoTests.vb`](tests/Iris.Tests/AtualizacaoTests.vb) |
| Assinar (o que os scripts chamam) | [`Assinador.vb`](tools/Iris.Assinatura/Assinador.vb) |
| A linha de comando do assinador | [`Program.vb`](tools/Iris.Assinatura/Program.vb) |

O `Iris.Update` é o **segundo** assembly de produção com rede, e o
`EgressArquiteturaTests` cobra os dois pelo nome. A diferença entre eles está
escrita lá: o de IA *manda* conteúdo do dono, e por isso tem portão, cofre e
diário; este só emite `GET`, com corpo vazio, e o que ele tem no lugar é
assinatura, porque o risco dele é o que chega.

"Arquivo público" é a implantação prevista, e não algo que o código imponha:
ele exige `https` e mais nada. Uma chave legítima pode assinar um manifesto
apontando para qualquer endereço `https`. Quem protege o conteúdo é o SHA-256
de dentro do manifesto assinado, e não o nome do servidor.
