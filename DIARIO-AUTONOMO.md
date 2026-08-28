# Diário do trabalho autônomo — a partir de 28/08/2026

O usuário autorizou executar as fases restantes sem ele na máquina, validando os
pontos críticos com o Codex e anotando os bloqueios. Este arquivo é o registro
corrido; o relatório final sai dele.

## Limites que eu declarei e não cruzo

1. **Nenhuma cerimônia é executada por mim.** Nem `--autorizar` do ambiente, nem
   alargamento da lista de pastas da IA, nem ACL. Construo e deixo sem apertar.
2. **Nada é enviado por e-mail**, em fase nenhuma.
3. **Nenhum conteúdo real vai para provedor externo.** O que precisar de
   provedor é provado contra `127.0.0.1`.
4. **Mutação no Outlook do usuário só em pasta de teste criada por mim**, nunca
   na Caixa de Entrada e nunca em item dele.

Cruzar qualquer um destes seria decidir no lugar dele enquanto ele não está.

## Ordem de execução escolhida

As dívidas primeiro, porque são pré-condição escrita e porque fecham a Fase 2 de
verdade; depois as medições, que destravam a pré-condição da Fase 4; depois as
fases novas.

| # | Bloco | Estado |
|---|---|---|
| 1 | Nove caminhos de guarda sem teste (4 categorias) | em andamento |
| 2 | Medir efeito da janela de sincronização | a fazer |
| 3 | Medir qualidade e utilidade do acervo parcial | a fazer |
| 4 | Falha rara da suíte: explicar ou encerrar | a fazer |
| 5 | Dívidas: broker/fechamento, auxiliar `Unico` | a fazer |
| 6 | Fase 4 — triagem e busca semântica, fechada | a fazer |
| 7 | Fase 5 — tarefas | a fazer |
| 8 | Fase 6 — calendário | a fazer |
| 9 | Fase 7 — contatos | a fazer |

---

## Registro

### 28/08 — bloco 1: guardas sem teste

**Categoria 1 — descarte do assistente com transmissão em voo.** Fechada. A
condição era um transmissor que ignore o token, e ela existe agora
(`ProvedorControlado.IgnorarCancelamento`). Três testes, controle negativo dos
dois lados. O número mediu diferente do que eu supus: desligar as duas guardas
do `Finally` produz **7** notificações, não 2 — escrever `Ocupado` reavalia a
cadeia inteira de comandos.

**Categoria 2 — troca de sessão durante a expansão da árvore.** Fechada.
`FakeBroker` ganhou stores e filhas configuráveis com travas; nasceu
`ArvoreDePastasTests`. Achado: os `Atual(geracao)` são **correntes de
conferências independentemente suficientes** — remover uma passa, remover todas
falha. Os testes provam a propriedade da cadeia, e está escrito assim.

**Categoria 3 — guardas de descarte da janela principal.** Três de quatro
caminhos fechados. A razão real de não haver primeiro teste não era a receita:
`MainViewModel` **não podia ser construído** numa suíte, porque o construtor
abria o cache do usuário. O caminho virou parâmetro opcional; produção continua
sem escolher.

- abertura pendente → **coberto**, controle negativo confirmado
- restauração de pasta vencida por sessão nova → **coberto**, controle negativo
  confirmado
- recarga da árvore concluindo depois do fechamento → **cerca de regressão, não
  cobertura**. Medi: removendo o `Folders.Clear()` do `Dispose` o teste continua
  verde, e o estado do broker é idêntico. Quem segura é a cadeia do
  `FolderTreeViewModel`, já coberta em outro arquivo. Vai para o relatório.
- recarga de contas pendente → **não isolável** pelo mesmo motivo: depois do
  `Dispose` a seleção já é `Nothing`, então o `Apontar` não é chamado com ou sem
  guarda. Vai para o relatório.

Suíte: **816**, 0 falhas, 0 pulados.
