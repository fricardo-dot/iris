# Testa o algoritmo de paginacao — O MESMO que o script real usa.
#
# Nao toca no Outlook.
#
# A versao anterior deste teste avancava a fronteira com AddSeconds(-1),
# que e "< T", enquanto o script real reabria com "<= T". Os cenarios
# passavam provando um algoritmo MELHOR do que o implementado — teste que
# nao testa o que diz. Agora os dois chamam Invoke-PaginacaoPorCursor, em
# paginacao.ps1: a unica diferenca e de onde as linhas vem.
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

    # Dentro do empate a ordem e EMBARALHADA quando OrdemInstavel: o OOM
    # nao promete ordem estavel num empate, e um algoritmo que dependa
    # dela esta errado mesmo passando.
    $ordenado = @($conjunto | Sort-Object -Property @{Expression='Quando'; Descending=$true})
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

function Testar {
    param([string]$Nome, [object[]]$Itens, [int]$Pagina, [switch]$Instavel)

    $fonte = New-Fonte $Itens -OrdemInstavel:$Instavel
    $r = Invoke-PaginacaoPorCursor `
        -Abrir  { param($f, $i) Abrir-Fonte $fonte $f $i } `
        -Ler    { param($c, $n) Ler-Fonte $c $n } `
        -Fechar { param($c) } `
        -TamanhoDaPagina $Pagina

    $ok = ($r.Lidos -eq $Itens.Count)
    Write-Host ("{0,-40} | {1,5} | {2,5} | {3,9} | {4}" -f `
        $Nome, $Itens.Count, $r.Lidos, $r.Consultas,
        $(if ($ok) { "OK" } else { "PERDEU $($Itens.Count - $r.Lidos)" }))
    return $ok
}

Write-Host ("{0,-40} | {1,5} | {2,5} | {3,9} | {4}" -f "cenario","total","lidos","consultas","resultado")
Write-Host ("-" * 84)

$ok = $true
$ok = (Testar "sem empate"                        (Montar 300 0 0)     50) -and $ok
$ok = (Testar "empate de 3 (como a caixa real)"   (Montar 150 3 150)   50) -and $ok

# O CASO QUE FALTAVA: grupo empatado >= pagina, com itens MAIS ANTIGOS
# depois dele. Sem o avanco ESTRITO apos drenar, a paginacao para aqui.
$ok = (Testar "empate de 50 (= pagina) + antigos" (Montar 100 50 200)  50) -and $ok
$ok = (Testar "empate de 100 (2x) + antigos"      (Montar 100 100 200) 50) -and $ok
$ok = (Testar "empate de 500 (10x) + antigos"     (Montar 100 500 300) 50) -and $ok
$ok = (Testar "tudo no mesmo segundo"             (Montar 0 200 0)     50) -and $ok
$ok = (Testar "empate no FIM da pasta"            (Montar 200 100 0)   50) -and $ok
$ok = (Testar "empate de 200, ordem INSTAVEL"     (Montar 100 200 100) 50 -Instavel) -and $ok

Write-Host ""
if ($ok) { Write-Host "Todos os cenarios leram tudo." }
else { Write-Host "ALGUM CENARIO PERDEU ITEM."; exit 1 }
