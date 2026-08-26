<#
.SYNOPSIS
    Monta o texto da cerimonia de ativacao da IA, com os identificadores da
    pasta ja preenchidos. NAO grava nada.

.DESCRIPTION
    A ativacao precisa do StoreID e do EntryID da pasta autorizada, e esses
    identificadores nao aparecem em lugar nenhum da interface do Outlook.
    Transcrever a mao um identificador de 140 caracteres hexadecimais e um
    convite ao erro -- e o erro seria silencioso: a ativacao carregaria,
    autorizando uma pasta que nao existe, e a IA negaria tudo sem dizer por que.

    Este roteiro le o Outlook (SO LEITURA, nada e criado, movido ou apagado),
    acha a pasta pelo nome, e IMPRIME o JSON pronto na tela.

    NAO GRAVA O ARQUIVO DE PROPOSITO. A ativacao e o unico ponto do Iris em que
    voce assume, por escrito, que conteudo da sua caixa pode sair da maquina.
    Um roteiro que gravasse sozinho transformaria isso num comando entre
    outros. Leia o que saiu, confira, e salve voce mesmo.

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
    [switch] $PoliticaVerificada
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

$destino = Join-Path $env:LOCALAPPDATA "Iris\ativacao.json"

Write-Host "----- confira, e salve em: $destino -----" -ForegroundColor Cyan
Write-Host ""
Write-Host $json
Write-Host ""
Write-Host "----------------------------------------------------------------" -ForegroundColor Cyan
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
