# Iris — Escopo do Projeto

**Status:** rascunho para discussão
**Data:** 2026-08-22

---

## 1. Visão

Cliente de produtividade pessoal semelhante ao Outlook, em Visual Basic, com
inteligência artificial integrada ao fluxo de trabalho — não como um chat
lateral, mas dentro das ações do dia a dia: ler, triar, responder e encontrar.

Uso pessoal, máquina única (Windows).

---

## 2. Restrição fundadora

**Microsoft Graph está descartado.** O acesso exige consentimento de
administrador do tenant corporativo, que não está disponível.

Consequência: o Iris não se autentica contra servidor nenhum. Ele usa o
**Outlook clássico já instalado e autenticado** na máquina como camada de
acesso aos dados, via automação COM (Outlook Object Model).

Isso elimina de uma vez: OAuth2, registro de aplicativo, tokens, refresh,
IMAP/SMTP, senhas de aplicativo e armazenamento de credenciais de e-mail.

---

## 3. Arquitetura

```
+---------------------------------------------------+
|  Iris (aplicativo VB, janela própria)             |
|                                                   |
|  UI  ->  Núcleo  ->  Cache local (SQLite)         |
|              |                                    |
|              +--> Outlook Object Model (COM)      |
|              |        |                           |
|              |        +--> Outlook clássico       |
|              |             (instalado e ABERTO)   |
|              |                                    |
|              +--> API de IA (HTTPS)               |
+---------------------------------------------------+
```

**Papéis:**

- **Outlook clássico** — motor de dados e transporte. Envia, recebe,
  sincroniza com o servidor. O Iris nunca fala com o Exchange diretamente.
- **Iris** — interface própria e toda a camada de inteligência.
- **Cache local (SQLite)** — espelho dos itens para busca, triagem e
  embeddings. Existe porque percorrer milhares de itens via COM é lento
  demais para uma UI responsiva.
- **API de IA** — chamadas HTTPS para os recursos de IA.

**Requisito de operação:** o Outlook precisa estar instalado e em execução.
O Iris depende dele.

---

## 4. Módulos

Todos os quatro pilares do Outlook, entregues em fases.

| Módulo | Objeto COM | Fase |
|---|---|---|
| E-mail | `MailItem` | 1 |
| Tarefas | `TaskItem` | 2 |
| Calendário | `AppointmentItem` | 3 |
| Contatos | `ContactItem` | 3 |

### Fase 1 — E-mail (núcleo utilizável)
- Listar pastas e mensagens da caixa de entrada
- Ler mensagem e thread
- Escrever, responder, encaminhar
- Anexos
- Busca por texto
- Cache local funcionando

### Fase 2 — IA + tarefas
- Recursos de IA sobre e-mail (seção 5)
- Tarefas, incluindo extração de tarefas a partir de e-mails

### Fase 3 — Calendário e contatos
- Agenda, compromissos, convites
- Catálogo de contatos, integrado ao remetente da mensagem

---

## 5. Recursos de IA

Todos os quatro foram marcados como prioritários. Ordem sugerida por
relação valor/custo:

1. **Resumir thread ou mensagem** — condensa conversas longas.
   Maior impacto imediato, menor complexidade.
2. **Redigir e responder** — gera rascunho a partir de instrução curta,
   com o contexto da thread. Rascunho sempre revisável; nunca envia sozinho.
3. **Triagem automática** — classifica por prioridade e categoria, sinaliza
   o que exige ação. Precisa de critérios ajustáveis pelo usuário.
4. **Busca semântica** — encontra por sentido, não por palavra exata.
   O mais caro: exige gerar embeddings de cada item e guardá-los localmente.
   Depende do cache da Fase 1 estar sólido.

---

## 6. Fora de escopo

- Múltiplos usuários, perfis ou contas simultâneas
- Instalador e mecanismo de atualização
- Versão web ou mobile
- Substituir o Outlook como cliente de sincronização
- Suporte ao "novo Outlook" (ver riscos)

---

## 7. Riscos

**R1 — Migração para o "novo Outlook".** O novo cliente da Microsoft não
expõe o Object Model via COM. Se a TI forçar a migração e desinstalar o
Outlook clássico, o Iris para de funcionar por completo.
*Impacto: fatal. Mitigação: isolar todo o acesso COM atrás de uma interface
única, para que uma futura fonte de dados alternativa seja substituível.*

**R2 — Guarda de segurança do Object Model.** O Outlook pode exibir avisos
ou bloquear acesso programático a endereços de destinatários e ao envio de
mensagens, dependendo da política e do antivírus registrado no sistema.
*Mitigação: validar cedo, na Fase 1, com um teste real de envio.*

**R3 — Desempenho do COM.** Iterar coleções grandes item a item é lento.
*Mitigação: usar consultas restritas do próprio Outlook e sincronizar para
o cache local; a UI lê do cache, nunca do COM.*

**R4 — Privacidade.** Os recursos de IA enviam conteúdo de e-mails a um
serviço externo. Uso pessoal, decisão consciente do usuário.
*Mitigação: enviar apenas o necessário; recurso de IA sempre acionado
explicitamente, nunca em varredura automática silenciosa.*

**R5 — Escopo grande.** Quatro módulos e quatro recursos de IA é bastante
trabalho. *Mitigação: as fases acima; cada fase precisa ser utilizável
sozinha.*

---

## 8. Decisões pendentes

- [ ] Confirmar a abordagem COM com Outlook clássico
- [ ] Escolher a stack (ver seção 9)
- [ ] Escolher o provedor de IA e o modelo
- [ ] Definir o visual: parecido com o Outlook ou identidade própria

---

## 9. Stack — proposta

A restrição da seção 2 já elimina boa parte das opções: o app precisa
consumir COM e falar HTTPS/JSON com uma API de IA.

| Opção | COM | HTTPS + JSON | Veredito |
|---|---|---|---|
| VB6 | nativo | sofrível — TLS moderno e JSON são dor | descartado |
| VBA dentro do Outlook | nativo | ruim, e sem janela própria | descartado |
| VB.NET no .NET Framework 4.8 | excelente | bom | viável |
| VB.NET no .NET 8 | bom | excelente | viável |

Recomendação em discussão. Ver conversa.
