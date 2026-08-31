<#
.SYNOPSIS
    Monta o texto da cerimonia de ativacao da IA, com os identificadores da
    pasta ja preenchidos. So grava com -Salvar.

.DESCRIPTION
    A ativacao precisa do StoreID e do EntryID da pasta autorizada, e esses
    identificadores nao aparecem em lugar nenhum da interface do Outlook.
    Transcrever a mao um identificador de 140 caracteres hexadecimais e um
    convite ao erro -- e o erro seria silencioso: a ativacao carregaria,
    autorizando uma pasta que nao existe, e a IA negaria tudo sem dizer por que.

    Este roteiro le o Outlook (SO LEITURA, nada e criado, movido ou apagado),
    acha a pasta pelo nome, e IMPRIME o JSON pronto na tela.

    Por padrao NAO grava: a ativacao e o unico ponto do Iris em que voce
    assume, por escrito, que conteudo da sua caixa pode sair da maquina.

    Com -Salvar ele grava em %ProgramData%\Iris\ativacao.json e provisiona a
    pasta com ACL propria. O ato deliberado continua sendo seu -- pasta,
    modelo e prazo saem destes parametros --, e o que -Salvar tira e a
    transcricao a mao, que so acrescentava chance de errar.

.EXAMPLE
    Caminho ABSOLUTO, que e o caso normal: um PowerShell novo abre em
    C:\WINDOWS\system32, e caminho relativo ao repositorio nao existe la.

    powershell -ExecutionPolicy Bypass -File "C:\Users\Ricardo\Documents\Iris\tools\preparar-ativacao.ps1" -Pasta "Iris-Teste"

.EXAMPLE
    Ja dentro do repositorio:

    .\tools\preparar-ativacao.ps1 -Pasta "Iris-Teste" -Modelo "anthropic/claude-sonnet-5" -Dias 30
#>
[CmdletBinding()]
param(
    # UMA OU VARIAS. Autorizar tres pastas eram tres execucoes, e cada uma
    # reescrevia a ativacao inteira -- entao a terceira apagava as duas
    # primeiras, em silencio. Nao era so trabalho repetido: era uma armadilha.
    #
    # O ato deliberado nao muda: os nomes continuam saindo daqui, escritos por
    # voce. O que sai e a repeticao, que so acrescentava chance de errar.
    [Parameter(Mandatory = $true)]
    [string[]] $Pasta,

    [string] $Modelo = "google/gemini-3.7-flash",

    # Os provedores subjacentes que o pedido vai aceitar. Gateway roteia: sem
    # isto, quem processa o e-mail pode ser outro, com retencao propria.
    #
    # SEM PADRAO, de proposito. O padrao era "google", que nao e slug de
    # provedor nenhum -- e o valor errado so apareceu no primeiro pedido de
    # verdade. Um padrao aqui e uma escolha de roteamento feita por quem
    # escreveu o roteiro, e nao por quem assina a autorizacao.
    [Parameter(Mandatory = $true)]
    [string[]] $Provedores,

    # QUAIS OPERACOES a autorizacao cobre.
    #
    # Eram duas, fixas no corpo do roteiro: Resumir e Redigir. Classificar
    # entrou em 31/08/2026 e NAO foi para o padrao, de proposito -- ela e a
    # unica que manda a pasta em LOTES e grava o resultado no cache, onde ele
    # sobrevive a sessao. Quem quiser assina de novo, com ela na lista.
    #
    # SEM [ValidateSet], pelo mesmo motivo de -Leituras: o ValidateSet roda no
    # binding, antes de a divisao por virgula acontecer.
    [string[]] $Operacoes = @("Resumir", "Redigir"),

    # O Iris recusa prazo acima de 90 dias.
    [ValidateRange(1, 90)]
    [int] $Dias = 30,

    # Quais RESULTADOS de leitura de rotulo a ativacao aceita.
    #
    # ERA FIXO EM "Absent", e isso custou uma investigacao inteira. Medido em
    # 30/08/2026 numa caixa real: as 13 mensagens de uma pasta devolviam
    # MSIP_Labels como STRING VAZIA, que o Iris classifica como "Blank" --
    # propriedade existe, valor vazio. "Absent" e outra coisa: a propriedade
    # nao foi encontrada.
    #
    # Sao os dois jeitos de "esta mensagem nao tem rotulo", e a ativacao
    # autorizava um so. O portao entao recusava tudo, e a tela dizia "nao foi
    # possivel classificar" -- que era falso: classificou muito bem.
    #
    # O padrao inclui os dois. Present, HistoricalOnly e o resto continuam de
    # fora: autorizar mensagem CLASSIFICADA e outra decisao, e ela e sua.
    # SEM [ValidateSet], e a razao e do PowerShell: ValidateSet roda no
    # BINDING do parametro, antes de o corpo do roteiro rodar -- entao a
    # divisao por virgula (necessaria com `-File`, ver acima) nunca chega a
    # acontecer, e "Absent,Blank" e rejeitado como se fosse um valor unico
    # invalido.
    #
    # Terceira variacao da mesma armadilha neste arquivo, e a unica em que o
    # remedio das outras duas nao serve. A conferencia foi para o corpo, logo
    # depois da divisao.
    [string[]] $Leituras = @("Absent", "Blank"),

    # Voce verificou a politica corporativa aplicavel? O padrao e a resposta
    # honesta para quem nao verificou.
    [switch] $PoliticaVerificada,

    # Grava direto no lugar certo, em vez de so imprimir.
    #
    # A primeira versao so imprimia, para o ato ser deliberado. Na pratica o
    # deliberado ficou intacto -- quem escolhe pasta, modelo e prazo e voce,
    # nestes parametros -- e o que sobrou foi transcricao: o texto de
    # orientacao acabou colado dentro do JSON, e o arquivo foi parar no
    # diretorio do repositorio. Copiar a mao nunca foi a cerimonia; era so
    # uma chance a mais de errar.
    #
    # Continua opt-in: sem esta chave, nada e gravado.
    [switch] $Salvar
,

    # Substitui uma ativacao que ja existe.
    #
    # A recusa em sobrescrever e proposital: trocar em silencio a lista de
    # pastas autorizadas seria trocar o que voce assinou sem lhe dizer. Mas o
    # unico caminho que sobrava era APAGAR o arquivo a mao -- e apagar destroi
    # o registro do que estava autorizado antes, que e justamente o que se
    # quer poder conferir depois.
    #
    # Com esta chave a antiga vira ativacao-AAAAMMDD-HHMMSS.json ao lado, e a
    # nova entra. O ato continua deliberado: sem a chave, nada e trocado.
    [switch] $Substituir
)

$ErrorActionPreference = 'Stop'

# VIRGULA SEPARA, MESMO CHAMADO COM -File.
#
# Com `powershell -File roteiro.ps1 -Provedores a,b` os argumentos chegam
# LITERAIS: o PowerShell nao interpreta a virgula, e $Provedores vira um
# arranjo de UM elemento com a string "a,b" dentro. O roteiro entao reclamava
# que "a,b" nao existe enquanto imprimia "a" e "b" como disponiveis -- uma
# mensagem que acusa o usuario de um erro que e do roteiro.
#
# Dividir aqui faz as duas formas de chamada valerem igual.
$Provedores = @($Provedores |
    ForEach-Object { $_ -split ',' } |
    ForEach-Object { $_.Trim() } |
    Where-Object { $_ })

# E O MESMO VALE PARA -Pasta, que virou lista depois do -Provedores.
#
# Eu escrevi o comentario acima, e nao apliquei a licao ao parametro que
# acabei de transformar em arranjo. O sintoma foi identico e igualmente
# acusador: "Nao achei nenhuma pasta chamada 'Caixa de Entrada,0. E-mails
# Lidos,1. Backup'" -- o roteiro imprimindo o proprio defeito como se fosse
# erro de quem digitou.
#
# RESSALVA QUE NAO DA PARA CONTORNAR AQUI: nome de pasta PODE conter virgula,
# e nesse caso esta divisao quebra o nome em dois e nenhum dos dois existe. O
# roteiro vai dizer que nao achou, e o motivo sera este. Chamar com
# `-File` e virgula nao tem como distinguir os dois casos -- quem tiver pasta
# com virgula no nome precisa chamar sem `-File`:
#   powershell -Command "& '...\preparar-ativacao.ps1' -Pasta 'A, com virgula'"
$Pasta = @($Pasta |
    ForEach-Object { $_ -split ',' } |
    ForEach-Object { $_.Trim() } |
    Where-Object { $_ })

$Leituras = @($Leituras |
    ForEach-Object { $_ -split ',' } |
    ForEach-Object { $_.Trim() } |
    Where-Object { $_ })

$Operacoes = @($Operacoes |
    ForEach-Object { $_ -split ',' } |
    ForEach-Object { $_.Trim() } |
    Where-Object { $_ })

# A conferencia que o ValidateSet faria, aqui -- DEPOIS da divisao.
$operacoesValidas = @("Resumir", "Redigir", "Classificar")
$forasDeOperacao = @($Operacoes | Where-Object { $operacoesValidas -notcontains $_ })
# DOIS ERROS DIFERENTES, DUAS MENSAGENS. Juntos, a lista vazia imprimia
# "Operacao(oes) que nao existem:" sem nada depois dos dois pontos -- um
# diagnostico que manda procurar o valor errado onde nao ha valor nenhum.
if ($Operacoes.Count -eq 0) {
    Write-Host "Nenhuma operacao autorizada." -ForegroundColor Red
    Write-Host ("Escolha entre: {0}" -f ($operacoesValidas -join ", ")) -ForegroundColor Red
    Write-Host ""
    Write-Host "Uma ativacao sem operacao nenhuma nao serve para nada: o portao" -ForegroundColor Yellow
    Write-Host "recusa tudo e a tela nao explica por que." -ForegroundColor Yellow
    exit 1
}
if ($forasDeOperacao.Count -gt 0) {
    Write-Host ("Operacao(oes) que nao existem: {0}" -f ($forasDeOperacao -join ", ")) -ForegroundColor Red
    Write-Host ("Validas: {0}" -f ($operacoesValidas -join ", ")) -ForegroundColor Red
    Write-Host ""
    Write-Host "Nenhuma operacao autorizada e uma ativacao que nao serve para nada," -ForegroundColor Yellow
    Write-Host "e operacao inventada nao vira erro no Iris: vira portao fechado" -ForegroundColor Yellow
    Write-Host "sem ninguem entender por que." -ForegroundColor Yellow
    exit 1
}

$leiturasValidas = @("Absent", "Blank", "Present", "HistoricalOnly")
$forasDeSet = @($Leituras | Where-Object { $leiturasValidas -notcontains $_ })
if ($forasDeSet.Count -gt 0) {
    Write-Host ("Leitura(s) que nao existem: {0}" -f ($forasDeSet -join ", ")) -ForegroundColor Red
    Write-Host ("Validas: {0}" -f ($leiturasValidas -join ", ")) -ForegroundColor Red
    Write-Host ""
    Write-Host "Absent e Blank sao os dois jeitos de 'sem rotulo': a propriedade" -ForegroundColor Yellow
    Write-Host "nao existir, e ela existir vazia. Present e mensagem CLASSIFICADA," -ForegroundColor Yellow
    Write-Host "e autoriza-la a sair e outra decisao." -ForegroundColor Yellow
    exit 1
}

if (-not $Provedores) {
    Write-Host "Nenhum provedor informado." -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host ("Procurando {0} pasta(s) no Outlook (somente leitura): {1}" -f `
    $Pasta.Count, (($Pasta | ForEach-Object { "'$_'" }) -join ", ")) -ForegroundColor Cyan

$outlook = $null
$ns = $null
try {
    $outlook = New-Object -ComObject Outlook.Application
    $ns = $outlook.GetNamespace("MAPI")
} catch {
    Write-Host "Nao consegui falar com o Outlook. Ele esta aberto?" -ForegroundColor Red
    exit 1
}

# Varre a arvore inteira. Nao usa $pid como variavel: no PowerShell ele e
# somente-leitura (e o PID do processo) e a atribuicao aborta a execucao.
$achadas = New-Object System.Collections.ArrayList

# RAMO QUE NAO FOI VISTO. A travessia engolia falha ao abrir as filhas, e o
# "nao achei nenhuma pasta" logo abaixo virava afirmacao sobre uma arvore
# que nao foi percorrida inteira.
$script:ramosCegos = 0

function Percorrer($pastas, $trilha) {
    foreach ($f in $pastas) {
        $caminho = if ($trilha) { "$trilha\$($f.Name)" } else { $f.Name }
        if ($Pasta -contains $f.Name) {
            [void]$achadas.Add([pscustomobject]@{
                Caminho = $caminho
                EntryId = $f.EntryID
                StoreId = $f.StoreID
                Itens   = $f.Items.Count
            })
        }
        try { Percorrer $f.Folders $caminho } catch { $script:ramosCegos++ }
    }
}

Percorrer $ns.Folders ""

# CADA NOME PEDIDO TEM DE RESOLVER PARA EXATAMENTE UMA PASTA.
#
# Conferido nome a nome, e nao no total: com dois nomes pedidos, um achado
# duas vezes e outro nenhuma, o total daria dois e pareceria certo. Somar
# antes de conferir esconde os dois erros de uma vez.
$alvos = @()
$problema = $false

foreach ($nome in $Pasta) {
    $doNome = @($achadas | Where-Object { $_.Caminho -eq $nome -or $_.Caminho.EndsWith("\$nome") })

    if ($doNome.Count -eq 0) {
        if ($script:ramosCegos -gt 0) {
            Write-Host ("Nao achei pasta chamada '{0}' -- e {1} ramo(s) da arvore nao" -f $nome, $script:ramosCegos) -ForegroundColor Red
            Write-Host "  pude percorrer. Ela pode estar num deles." -ForegroundColor Red
        } else {
            Write-Host "Nao achei nenhuma pasta chamada '$nome'." -ForegroundColor Red
        }
        $problema = $true
        continue
    }

    if ($doNome.Count -gt 1) {
        Write-Host "Achei MAIS DE UMA pasta chamada '$nome':" -ForegroundColor Yellow
        $doNome | Format-Table Caminho, Itens -AutoSize
        Write-Host "Renomeie uma delas, ou me diga qual. Escolher por voce seria" -ForegroundColor Yellow
        Write-Host "autorizar uma pasta que talvez nao seja a que voce quis." -ForegroundColor Yellow
        $problema = $true
        continue
    }

    $alvos += $doNome[0]
}

# NENHUMA PASTA SE UMA FALHOU.
#
# Autorizar o subconjunto que deu certo seria pior que falhar: voce pediu
# tres, o arquivo sai com duas, e a ativacao passa a dizer menos do que voce
# assinou -- sem que nada na tela mostre a diferenca depois.
if ($problema) {
    Write-Host ""
    Write-Host "Nenhuma pasta foi autorizada. Corrija os nomes acima e rode de novo." -ForegroundColor Red
    exit 1
}

foreach ($a in $alvos) {
    Write-Host "Achei: $($a.Caminho)  ($($a.Itens) itens)" -ForegroundColor Green
}
Write-Host ""

# ---------------------------------------------------------------------------
# OS SLUGS SAO CONFERIDOS ANTES DE VIRAREM AUTORIZACAO.
#
# Esta conferencia existe porque o padrao era "google", e slug "google" NAO
# EXISTE -- os de verdade sao "google-vertex" e "google-ai-studio". "Google" e
# o nome de EXIBICAO. A ativacao passou a exigir um filtro que nao casava com
# nada, e o primeiro pedido de verdade morreu com o OpenRouter recusando.
#
# A falha foi fechada, porque allow_fallbacks e falso. Mas uma restricao que
# nao restringe coisa nenhuma nao devia chegar a ser escrita num arquivo de
# autorizacao: o lugar de pegar isso e aqui, na cerimonia, e nao no primeiro
# e-mail.
#
# O carregador do Iris NAO faz esta conferencia de proposito: ele le um arquivo
# e nao toca na rede. Validar slug exige perguntar ao provedor.
Write-Host "Conferindo os provedores contra os endpoints de '$Modelo'..." -ForegroundColor Cyan
try {
    $eps = (Invoke-RestMethod -Method Get -TimeoutSec 30 `
        -Uri "https://openrouter.ai/api/v1/models/$Modelo/endpoints").data.endpoints
} catch {
    Write-Host "Nao consegui consultar os endpoints do modelo." -ForegroundColor Red
    Write-Host $_.Exception.Message
    exit 1
}

if (-not $eps) {
    Write-Host "O OpenRouter nao lista endpoint nenhum para '$Modelo'." -ForegroundColor Red
    Write-Host "Confira o nome do modelo." -ForegroundColor Yellow
    exit 1
}

# O slug de roteamento e a PRIMEIRA parte do tag: "google-vertex/global" ->
# "google-vertex". Slug base casa todas as variantes.
$slugs = $eps | ForEach-Object { ($_.tag -split '/')[0] } | Select-Object -Unique | Sort-Object

Write-Host "  slugs disponiveis: $($slugs -join ', ')"
$maus = $Provedores | Where-Object { $slugs -notcontains $_ }
if ($maus) {
    Write-Host ""
    Write-Host "NAO EXISTE endpoint com o(s) slug(s): $($maus -join ', ')" -ForegroundColor Red
    Write-Host ""
    Write-Host "Uma lista que nao casa com nada faz o pedido ser RECUSADO -- e" -ForegroundColor Yellow
    Write-Host "isso e o desfecho certo, mas descobrir no primeiro e-mail e caro." -ForegroundColor Yellow
    Write-Host "Repita com, por exemplo:" -ForegroundColor Yellow
    Write-Host ("  -Provedores " + ($slugs -join ',')) -ForegroundColor Green
    exit 1
}
Write-Host "  ok: todos os slugs pedidos existem." -ForegroundColor Green
Write-Host ""

$agora = (Get-Date).ToUniversalTime()
$ate = $agora.AddDays($Dias)
$fmt = "yyyy-MM-ddTHH:mm:ssZ"

$json = [ordered]@{
    id                            = "ativacao-" + $agora.ToString("yyyy-MM-dd")
    versao                        = 1
    autoridade                    = "$env:USERNAME, dono da caixa"
    politicaCorporativaVerificada = [bool]$PoliticaVerificada
    quando                        = $agora.ToString($fmt)
    ate                           = $ate.ToString($fmt)
    provedor                      = "openrouter"
    endpoint                      = "https://openrouter.ai/api/v1/chat/completions"
    modelo                        = $Modelo
    regiao                        = "nao imposta pelo pedido"
    retencaoAceita                = "retencao zero exigida no proprio pedido"
    exigirRetencaoZero            = $true
    provedoresPermitidos          = $Provedores
    operacoes                     = @($Operacoes)
    pastas                        = @($alvos | ForEach-Object { @{ storeId = $_.StoreId; entryId = $_.EntryId } })
    rotulos                       = @()
    leituras                      = $Leituras
    contentBits                   = @(0)
} | ConvertTo-Json -Depth 5

# %ProgramData% e nao %LOCALAPPDATA%: a conferencia de permissao olha a
# pasta-mae, e num perfil real ela nao passa -- ACE herdada de sobra. Ver o
# doc de ActivationLoader.CaminhoPadrao.
#
# E resolvido pela MESMA API do produto, e nao por $env:ProgramData: a
# variavel de ambiente pode divergir da pasta que o .NET resolve, e a
# ferramenta escreveria num lugar que o Iris nao le.
$raiz = [Environment]::GetFolderPath([Environment+SpecialFolder]::CommonApplicationData)
$destino = Join-Path $raiz "Iris\ativacao.json"

Write-Host "----- inicio do JSON -----" -ForegroundColor Cyan
Write-Host $json
Write-Host "----- fim do JSON --------" -ForegroundColor Cyan
Write-Host ""

if ($Salvar) {
    if ((Test-Path $destino) -and -not $Substituir) {
        Write-Host "JA EXISTE uma ativacao em $destino." -ForegroundColor Red
        Write-Host "Nao vou sobrescrever em silencio: a lista de pastas autorizadas" -ForegroundColor Red
        Write-Host "e o que voce assinou, e troca-la calado seria trocar a sua" -ForegroundColor Red
        Write-Host "assinatura." -ForegroundColor Red
        Write-Host ""
        Write-Host "Repita com -Substituir. A antiga NAO e apagada: vira" -ForegroundColor Yellow
        Write-Host "ativacao-AAAAMMDD-HHMMSS.json ao lado, para voce poder" -ForegroundColor Yellow
        Write-Host "conferir depois o que estava autorizado antes." -ForegroundColor Yellow
        exit 1
    }

    if ((Test-Path $destino) -and $Substituir) {
        # GUARDA A ANTIGA ANTES DE ESCREVER, e so escreve se a copia deu certo.
        # Perder o registro do que estava autorizado seria perder a unica coisa
        # que responde "o que eu tinha assinado antes?".
        $carimbo = (Get-Date).ToString("yyyyMMdd-HHmmss")
        $guardada = Join-Path (Split-Path $destino) "ativacao-$carimbo.json"
        try {
            Copy-Item $destino $guardada -ErrorAction Stop
            Write-Host "A ativacao anterior foi guardada em:" -ForegroundColor Cyan
            Write-Host "  $guardada" -ForegroundColor Cyan
        } catch {
            Write-Host "NAO consegui guardar a ativacao anterior: $($_.Exception.Message)" -ForegroundColor Red
            Write-Host "Nao vou substituir sem ter salvo a antiga." -ForegroundColor Red
            exit 1
        }
    }
    $pastaDestino = Split-Path $destino
    New-Item -ItemType Directory -Force $pastaDestino | Out-Null

    # A PASTA NASCE COM ACL PROPRIA. Em %ProgramData% a heranca traz Users com
    # direito de criar, e o Iris recusa isso na pasta que contem a ativacao --
    # quem cria ali dentro troca o arquivo.
    $eu = [System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value
    $saida = & icacls $pastaDestino /grant:r `
        "*${eu}:(OI)(CI)(M)" "*S-1-5-18:(OI)(CI)(F)" "*S-1-5-32-544:(OI)(CI)(F)" `
        /inheritance:r 2>&1

    # FALHA DO icacls TEM DE PARAR AQUI, ANTES DE GRAVAR.
    #
    # No PowerShell 5.1 um executavel nativo que falha nao levanta erro: o
    # roteiro imprimiria "GRAVADO" sobre uma pasta que continua com a heranca
    # do %ProgramData% -- e Users pode criar ali dentro, o que e exatamente o
    # que o Iris recusa. Gravar antes de conferir seria anunciar sucesso sobre
    # uma ativacao que nao vai carregar.
    if ($LASTEXITCODE -ne 0) {
        Write-Host "NAO consegui ajustar a permissao da pasta. Nada foi gravado." -ForegroundColor Red
        Write-Host $saida
        exit 1
    }

    [IO.File]::WriteAllText($destino, $json, (New-Object Text.UTF8Encoding $false))
    Write-Host "GRAVADO em: $destino" -ForegroundColor Green
    Write-Host "Confira com:  dotnet run --project tools\Iris.CrashHarness -- ativacao" -ForegroundColor Green
} else {
    Write-Host "NAO gravei nada." -ForegroundColor Yellow
    Write-Host "Para gravar em $destino, repita o comando com -Salvar." -ForegroundColor Yellow
    Write-Host "Se for copiar a mao, copie SO o que esta entre as duas marcas." -ForegroundColor Yellow
}
Write-Host ""
Write-Host "Antes de salvar, leia:" -ForegroundColor Yellow
Write-Host "  * 'operacoes' autoriza: $($Operacoes -join ', ')."
if ($Operacoes -contains "Classificar") {
    Write-Host "  * CLASSIFICAR manda a pasta em LOTES e grava o rotulo no cache." -ForegroundColor Yellow
    Write-Host "    E diferente de resumir, que e um pedido por vez com o resultado" -ForegroundColor Yellow
    Write-Host "    na tela e nada gravado." -ForegroundColor Yellow
}
Write-Host "  * 'leituras' aceita: $($Leituras -join ', ')."
Write-Host "  * a autorizacao vence em $($ate.ToLocalTime().ToString('dd/MM/yyyy')) e"
Write-Host "    depois disso a IA volta a ficar desligada, de proposito."
if (-not $PoliticaVerificada) {
    Write-Host "  * politicaCorporativaVerificada = false. A faixa vai dizer isso" -ForegroundColor Yellow
    Write-Host "    o tempo todo, ate voce verificar." -ForegroundColor Yellow
}
Write-Host ""
Write-Host "E a chave, uma vez so (o Iris nunca a ve; quem guarda e o Windows):"
Write-Host "  cmdkey /generic:Iris/OpenRouter /user:iris /pass" -ForegroundColor Green
Write-Host ""
