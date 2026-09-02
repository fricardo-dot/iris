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
só para você; a pública é impressa na tela.

Cole a pública em [`ChaveDeAtualizacao.vb`](src/Iris.App/ChaveDeAtualizacao.vb),
junto com o endereço do manifesto:

```
https://github.com/<dono>/<repo>/releases/latest/download/iris.json
```

`latest/download` é um endereço estável: não é preciso saber o número da
última versão para perguntar qual é a última versão.

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

O script grava a versão em `Directory.Build.props`, publica autocontido
num `.exe` só, calcula o SHA-256 **do arquivo que acabou de sair**, escreve
`iris.json` com esse hash, e assina o `iris.json` inteiro.

Ele **não publica**. Deixa os três arquivos em `artefatos/<versao>/` e
imprime o comando `gh release create`. Subir a release é um ato seu, e um
ato público. Com `-Publicar`, ele sobe.

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

E não é o certificado de código do Windows. O SmartScreen vai avisar na
primeira execução de cada versão nova — "aplicativo não reconhecido", com
um "Mais informações" → "Executar assim mesmo". Some depois que o
executável ganha reputação, ou nunca, se cada versão for baixada por duas
pessoas. Resolver isso custa um certificado EV de uma autoridade, com
custo anual; para duas máquinas suas, não vale.

## Antes do primeiro push

Nada foi empurrado para lugar nenhum: os commits estão só aqui. Duas coisas
para decidir antes de o repositório existir:

1. **O histórico carrega o seu e-mail de trabalho.** Os arquivos foram
   limpos (commit `800fc45`), mas `64da727` e `16dcbbb` ainda têm a string.
   Como nada foi publicado, reescrever é barato — e depois do primeiro push
   deixa de ser.

2. **Repositório público ou privado.** As releases precisam ser públicas
   para `latest/download` funcionar sem token. Repositório privado exigiria
   hospedar os arquivos em outro lugar.

## Onde as peças moram

| O quê | Onde |
|---|---|
| O número da versão | [`Directory.Build.props`](Directory.Build.props) |
| A chave pública e o endereço | [`ChaveDeAtualizacao.vb`](src/Iris.App/ChaveDeAtualizacao.vb) |
| Ler e conferir o manifesto | [`ManifestoDeVersao.vb`](src/Iris.Update/ManifestoDeVersao.vb) |
| Perguntar e baixar | [`ProcuraDeVersao.vb`](src/Iris.Update/ProcuraDeVersao.vb) |
| A tela | [`AtualizacaoViewModel.vb`](src/Iris.App/ViewModels/AtualizacaoViewModel.vb) |
| Os testes, com as sabotagens | [`AtualizacaoTests.vb`](tests/Iris.Tests/AtualizacaoTests.vb) |

O `Iris.Update` é o **segundo** assembly de produção com rede, e o
`EgressArquiteturaTests` cobra os dois pelo nome. A diferença entre eles
está escrita lá: o de IA *manda* conteúdo do dono, e por isso tem portão,
cofre e diário; este só *busca* arquivo público, e o que ele tem no lugar é
assinatura, porque o risco dele é o que chega.
