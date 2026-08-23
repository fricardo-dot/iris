# Q1: testa o ALGORITMO de paginacao contra uma tabela SINTETICA.
#
# Nao toca no Outlook.
#
# Motivo: a caixa real tem no maximo 3 itens no mesmo segundo, e o caso que
# quebra a paginacao e uma pagina INTEIRA empatada. Com pagina de 50 e 80
# itens no mesmo instante:
#
#   1. a pagina 1 devolve 50 dos 80;
#   2. a consulta seguinte usa "<= T";
#   3. dentro do empate a ordem nao e contratual;
#   4. GetArray(50) pode devolver OS MESMOS 50;
#   5. nenhum item novo -> o algoritmo declara fim e perde 30, em silencio.
#
# A tabela sintetica devolve, de proposito, a ordem MAIS HOSTIL possivel
# dentro de um empate: sempre a mesma. Se o algoritmo sobrevive a isso,
# sobrevive a uma ordem instavel.

# ---------------------------------------------------------------
# Tabela de mentira: mesmo contrato do Table do OOM.
# ---------------------------------------------------------------
function New-TabelaFalsa {
    param([object[]]$Itens)   # cada item: @{ Id; Quando }
    $estado = [pscustomobject]@{
        Todos    = $Itens
        Filtrado = @()
        Pos      = 0
        Chamadas = 0
    }
    return $estado
}

function Abrir-TabelaFalsa {
    param($Tabela, [Nullable[datetime]]$AteInclusive)

    # Ordena por Quando decrescente. Dentro do empate mantem a ordem
    # original — a ordem "estavel hostil" descrita acima.
    $conjunto = if ($null -ne $AteInclusive) {
        $Tabela.Todos | Where-Object { $_.Quando -le $AteInclusive }
    } else { $Tabela.Todos }

    $Tabela.Filtrado = @($conjunto | Sort-Object -Property @{Expression='Quando'; Descending=$true})
    $Tabela.Pos = 0
    $Tabela.Chamadas++
}

function Get-LinhasFalsas {
    param($Tabela, [int]$Quantas)
    $fim = [Math]::Min($Tabela.Pos + $Quantas, $Tabela.Filtrado.Count)
    $r = @()
    for ($i = $Tabela.Pos; $i -lt $fim; $i++) { $r += $Tabela.Filtrado[$i] }
    $Tabela.Pos = $fim
    return $r
}

# ---------------------------------------------------------------
# O ALGORITMO em teste.
#
# Regra: quando a pagina termina no instante T, DRENA todo o grupo de T
# antes de avancar a fronteira para "< T". Sem drenar, um grupo maior que
# a pagina trava o avanco.
# ---------------------------------------------------------------
function Paginar {
    param($Tabela, [int]$TamanhoDaPagina, [int]$MaxChamadas = 500)

    $vistos = @{}
    $ordem = @()
    $fronteira = $null
    $chamadas = 0

    while ($chamadas -lt $MaxChamadas) {
        Abrir-TabelaFalsa $Tabela $fronteira
        $chamadas++

        $pagina = Get-LinhasFalsas $Tabela $TamanhoDaPagina
        if ($pagina.Count -eq 0) { break }

        $novos = 0
        $ultimo = $null
        foreach ($it in $pagina) {
            $ultimo = $it.Quando
            if ($vistos.ContainsKey($it.Id)) { continue }
            $vistos[$it.Id] = $true
            $ordem += $it.Id
            $novos++
        }

        # A pagina terminou dentro de um empate? Drena o grupo inteiro
        # ANTES de mexer na fronteira.
        $drenou = $false
        while ($chamadas -lt $MaxChamadas) {
            $extra = Get-LinhasFalsas $Tabela $TamanhoDaPagina
            if ($extra.Count -eq 0) { break }
            $aindaNoGrupo = $false
            foreach ($it in $extra) {
                if ($it.Quando -ne $ultimo) { break }
                $aindaNoGrupo = $true
                if ($vistos.ContainsKey($it.Id)) { continue }
                $vistos[$it.Id] = $true
                $ordem += $it.Id
                $novos++
                $drenou = $true
            }
            if (-not $aindaNoGrupo) { break }
        }

        # Com o grupo drenado, a fronteira pode andar com seguranca.
        if ($novos -eq 0 -and -not $drenou) { break }
        $fronteira = $ultimo.AddSeconds(-1)
    }

    return [pscustomobject]@{ Lidos = $ordem.Count; Chamadas = $chamadas }
}

# ---------------------------------------------------------------
# Casos
# ---------------------------------------------------------------
function Cenario {
    param([string]$Nome, [int]$Total, [int]$NoMesmoSegundo, [int]$Pagina)

    $base = Get-Date "2026-08-01 12:00:00"
    $itens = @()

    # Um bloco empatado no meio, o resto com instantes distintos.
    $distintos = $Total - $NoMesmoSegundo
    $metade = [int]($distintos / 2)

    for ($i = 0; $i -lt $metade; $i++) {
        $itens += [pscustomobject]@{ Id = "A$i"; Quando = $base.AddSeconds(-$i) }
    }
    $instanteEmpate = $base.AddSeconds(-$metade)
    for ($i = 0; $i -lt $NoMesmoSegundo; $i++) {
        $itens += [pscustomobject]@{ Id = "E$i"; Quando = $instanteEmpate }
    }
    for ($i = 0; $i -lt ($distintos - $metade); $i++) {
        $itens += [pscustomobject]@{ Id = "B$i"; Quando = $instanteEmpate.AddSeconds(-1 - $i) }
    }

    $t = New-TabelaFalsa $itens
    $r = Paginar $t $Pagina

    $ok = ($r.Lidos -eq $Total)

    # Write-Host, e nao a string solta: uma funcao do PowerShell devolve
    # TUDO que ela escreve no fluxo de saida. Com a string solta, a linha
    # da tabela virava parte do valor de retorno, o "-and" consumia as
    # duas coisas, e o teste passava sem imprimir nem verificar nada.
    Write-Host ("{0,-34} | {1,5} | {2,5} | {3,8} | {4}" -f `
        $Nome, $Total, $r.Lidos, $r.Chamadas, $(if ($ok) { "OK" } else { "PERDEU $($Total - $r.Lidos)" }))
    return $ok
}

Write-Host "cenario                            | total | lidos | chamadas | resultado"
Write-Host "-----------------------------------|-------|-------|----------|----------"

$todosOk = $true
$todosOk = (Cenario "sem empate"                    300  0   50) -and $todosOk
$todosOk = (Cenario "empate de 3 (como a caixa)"    300  3   50) -and $todosOk
$todosOk = (Cenario "empate de 51 (> pagina)"       300  51  50) -and $todosOk
$todosOk = (Cenario "empate de 100 (2x a pagina)"   300  100 50) -and $todosOk
$todosOk = (Cenario "empate de 500 (10x)"           900  500 50) -and $todosOk
$todosOk = (Cenario "TUDO no mesmo segundo"         200  200 50) -and $todosOk

Write-Output ""
if ($todosOk) {
    Write-Output "Todos os cenarios leram tudo."
} else {
    Write-Output "ALGUM CENARIO PERDEU ITEM."
    exit 1
}
