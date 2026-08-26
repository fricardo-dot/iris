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

    Com -Salvar ele grava em %ProgramData%\Iristivacao.json e provisiona a
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
    [Parameter(Mandatory = $true)]
    [string] $Pasta,

    [string] $Modelo = "google/gemini-3.7-flash",

    # Os provedores subjacentes que o pedido vai aceitar. Gateway roteia: sem
    # isto, quem processa o e-mail pode ser outro, com retencao propria.
    [string[]] $Provedores = @("google"),

    # O Iris recusa prazo acima de 90 dias.
    [ValidateRange(1, 90)]
    [int] $Dias = 30,

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
)

$ErrorActionPreference = 'Stop'

Write-Host ""
Write-Host "Procurando a pasta '$Pasta' no Outlook (somente leitura)..." -ForegroundColor Cyan

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

function Percorrer($pastas, $trilha) {
    foreach ($f in $pastas) {
        $caminho = if ($trilha) { "$trilha\$($f.Name)" } else { $f.Name }
        if ($f.Name -eq $Pasta) {
            [void]$achadas.Add([pscustomobject]@{
                Caminho = $caminho
                EntryId = $f.EntryID
                StoreId = $f.StoreID
                Itens   = $f.Items.Count
            })
        }
        try { Percorrer $f.Folders $caminho } catch { }
    }
}

Percorrer $ns.Folders ""

if ($achadas.Count -eq 0) {
    Write-Host "Nao achei nenhuma pasta chamada '$Pasta'." -ForegroundColor Red
    exit 1
}

if ($achadas.Count -gt 1) {
    Write-Host "Achei MAIS DE UMA pasta com esse nome:" -ForegroundColor Yellow
    $achadas | Format-Table Caminho, Itens -AutoSize
    Write-Host "Renomeie uma delas, ou me diga qual. Escolher por voce seria" -ForegroundColor Yellow
    Write-Host "autorizar uma pasta que talvez nao seja a que voce quis." -ForegroundColor Yellow
    exit 1
}

$alvo = $achadas[0]
Write-Host "Achei: $($alvo.Caminho)  ($($alvo.Itens) itens)" -ForegroundColor Green
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
    operacoes                     = @("Resumir", "Redigir")
    pastas                        = @(@{ storeId = $alvo.StoreId; entryId = $alvo.EntryId })
    rotulos                       = @()
    leituras                      = @("Absent")
    contentBits                   = @(0)
} | ConvertTo-Json -Depth 5

# %ProgramData% e nao %LOCALAPPDATA%: a conferencia de permissao olha a
# pasta-mae, e num perfil real ela nao passa -- ACE herdada de sobra. Ver o
# doc de ActivationLoader.CaminhoPadrao.
$destino = Join-Path $env:ProgramData "Iris\ativacao.json"

Write-Host "----- inicio do JSON -----" -ForegroundColor Cyan
Write-Host $json
Write-Host "----- fim do JSON --------" -ForegroundColor Cyan
Write-Host ""

if ($Salvar) {
    if (Test-Path $destino) {
        Write-Host "JA EXISTE uma ativacao em $destino." -ForegroundColor Red
        Write-Host "Nao vou sobrescrever: apague a antiga se for essa a intencao." -ForegroundColor Red
        exit 1
    }
    $pastaDestino = Split-Path $destino
    New-Item -ItemType Directory -Force $pastaDestino | Out-Null

    # A PASTA NASCE COM ACL PROPRIA. Em %ProgramData% a heranca traz Users com
    # direito de criar, e o Iris recusa isso na pasta que contem a ativacao --
    # quem cria ali dentro troca o arquivo.
    $eu = [System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value
    & icacls $pastaDestino /grant:r `
        "*${eu}:(OI)(CI)(M)" "*S-1-5-18:(OI)(CI)(F)" "*S-1-5-32-544:(OI)(CI)(F)" `
        /inheritance:r | Out-Null

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
Write-Host "  * 'operacoes' autoriza resumir E redigir. Tire um se quiser menos."
Write-Host "  * 'leituras' aceita SO mensagem sem rotulo de sensibilidade."
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
