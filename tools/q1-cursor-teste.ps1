# Testa o algoritmo de paginacao — O MESMO que o script real usa.
#
# Nao toca no Outlook.
#
# Duas coisas que este teste faz e a versao anterior nao fazia:
#
#   1. Chama Invoke-PaginacaoPorCursor, de paginacao.ps1 — o mesmo codigo
#      que q1-cursor.ps1 roda contra o Outlook. Antes o teste tinha uma
#      copia do algoritmo que avancava com "< T" enquanto o real usava
#      "<= T", entao os cenarios passavam provando outra coisa.
#
#   2. Roda os CONTROLES NEGATIVOS de verdade, com os defeitos ligados por
#      parametro. Antes eu editava o arquivo a mao e anotava o resultado —
#      numero anotado nao e regressao verificavel.
#
# A caixa real tem no maximo 3 itens no mesmo segundo, entao o caso que
# quebra — grupo empatado MAIOR que a pagina — so da para exercitar aqui.

. (Join-Path $PSScriptRoot "paginacao.ps1")

function New-Fonte {
    param([object[]]$Itens, [switch]$OrdemInstavel)
    return [pscustomobject]@{ Itens = $Itens; Instavel = [bool]$OrdemInstavel }
}

function Abrir-Fonte {
    param($Fonte, $Fronteira, [bool]$Inclusivo)

    $conjunto = if ($null -eq $Fronteira) { $Fonte.Itens } else {
        if ($Inclusivo) { $Fonte.Itens | Where-Object { $_.Quando -le $Fronteira } }
        else            { $Fonte.Itens | Where-Object { $_.Quando -lt $Fronteira } }
    }

    $ordenado = @($conjunto | Sort-Object -Property @{Expression='Quando'; Descending=$true})

    # Dentro do empate a ordem e EMBARALHADA a cada abertura: o OOM nao
    # promete ordem estavel ali, e algoritmo que dependa dela esta errado
    # mesmo passando.
    if ($Fonte.Instavel) {
        $novo = @()
        foreach ($g in ($ordenado | Group-Object Quando | Sort-Object { [datetime]$_.Name } -Descending)) {
            $novo += ($g.Group | Sort-Object { Get-Random })
        }
        $ordenado = $novo
    }

    return [pscustomobject]@{ Linhas = $ordenado; Pos = 0 }
}

function Ler-Fonte {
    param($Cursor, [int]$Quantas)
    $fim = [Math]::Min($Cursor.Pos + $Quantas, $Cursor.Linhas.Count)
    $r = @()
    for ($i = $Cursor.Pos; $i -lt $fim; $i++) { $r += $Cursor.Linhas[$i] }
    $Cursor.Pos = $fim
    return ,$r
}

function Montar {
    param([int]$Antes, [int]$Empatados, [int]$Depois)
    $base = Get-Date "2026-08-01 12:00:00"
    $itens = @()
    for ($i = 0; $i -lt $Antes; $i++) {
        $itens += [pscustomobject]@{ Id = "A$i"; Quando = $base.AddSeconds(-$i) }
    }
    $t = $base.AddSeconds(-$Antes)
    for ($i = 0; $i -lt $Empatados; $i++) {
        $itens += [pscustomobject]@{ Id = "E$i"; Quando = $t }
    }
    for ($i = 0; $i -lt $Depois; $i++) {
        $itens += [pscustomobject]@{ Id = "B$i"; Quando = $t.AddSeconds(-1 - $i) }
    }
    return ,$itens
}

function Rodar {
    param([object[]]$Itens, [int]$Pagina, [switch]$Instavel,
          [switch]$SemDrenagem, [switch]$FronteiraInclusiva)

    $fonte = New-Fonte $Itens -OrdemInstavel:$Instavel
    return Invoke-PaginacaoPorCursor `
        -Abrir  { param($f, $i) Abrir-Fonte $fonte $f $i } `
        -Ler    { param($c, $n) Ler-Fonte $c $n } `
        -Fechar { param($c) } `
        -TamanhoDaPagina $Pagina `
        -SemDrenagem:$SemDrenagem -FronteiraInclusiva:$FronteiraInclusiva
}

$casos = @(
    @{ Nome = "sem empate";                        Itens = (Montar 300 0 0)     },
    @{ Nome = "empate de 3 (como a caixa real)";   Itens = (Montar 150 3 150)   },
    @{ Nome = "empate de 50 (= pagina) + antigos"; Itens = (Montar 100 50 200)  },
    @{ Nome = "empate de 100 (2x) + antigos";      Itens = (Montar 100 100 200) },
    @{ Nome = "empate de 500 (10x) + antigos";     Itens = (Montar 100 500 300) },
    @{ Nome = "tudo no mesmo segundo";             Itens = (Montar 0 200 0)     },
    @{ Nome = "empate no FIM da pasta";            Itens = (Montar 200 100 0)   },
    @{ Nome = "empate de 200, ordem INSTAVEL";     Itens = (Montar 100 200 100); Instavel = $true }
)

Write-Host ("{0,-40} | {1,5} | {2,8} | {3,10} | {4,10}" -f `
    "cenario", "total", "correto", "sem drenar", "inclusiva")
Write-Host ("-" * 88)

$tudoOk = $true
$pegouSemDrenagem = $false
$pegouInclusiva = $false

foreach ($c in $casos) {
    $inst = [bool]$c.Instavel
    $total = $c.Itens.Count

    $bom  = (Rodar $c.Itens 50 -Instavel:$inst).Lidos
    $sem  = (Rodar $c.Itens 50 -Instavel:$inst -SemDrenagem).Lidos
    $incl = (Rodar $c.Itens 50 -Instavel:$inst -FronteiraInclusiva).Lidos

    if ($bom -ne $total) { $tudoOk = $false }

    # Os DOIS defeitos precisam ser discriminados, cada um por si. Um
    # guarda que aceitasse "algum dos dois" deixaria passar o dia em que um
    # deles parasse de perder item — e ai metade do controle negativo
    # viraria decoracao sem ninguem notar.
    if ($sem  -ne $total) { $pegouSemDrenagem = $true }
    if ($incl -ne $total) { $pegouInclusiva = $true }

    Write-Host ("{0,-40} | {1,5} | {2,8} | {3,10} | {4,10}" -f `
        $c.Nome, $total,
        $(if ($bom -eq $total) { "OK" } else { "PERDEU $($total-$bom)" }),
        $(if ($sem -eq $total) { "-" } else { "perde $($total-$sem)" }),
        $(if ($incl -eq $total) { "-" } else { "perde $($total-$incl)" }))
}

Write-Host ""
if (-not $tudoOk) {
    Write-Host "FALHA: o algoritmo correto perdeu item."
    exit 1
}
$faltou = @()
if (-not $pegouSemDrenagem) { $faltou += "-SemDrenagem" }
if (-not $pegouInclusiva)   { $faltou += "-FronteiraInclusiva" }

if ($faltou.Count -gt 0) {
    Write-Host ("FALHA: nenhum cenario discriminou {0}." -f ($faltou -join " e "))
    Write-Host "Controle negativo que nao perde item nao esta controlando nada."
    exit 1
}
Write-Host "O algoritmo correto leu tudo, e os DOIS defeitos foram pegos."
