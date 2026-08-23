# O ALGORITMO de paginacao por cursor. Um lugar so.
#
# Existe porque o teste sintetico anterior testava um algoritmo DIFERENTE
# do script real. Agora o real e o teste chamam ESTA funcao; a unica
# diferenca entre eles e de onde as linhas vem.
#
# ---------------------------------------------------------------------
# O PROBLEMA
#
# ReceivedTime nao e ordem total: varios itens compartilham o mesmo
# instante, e o OOM nao promete ordem estavel dentro do empate.
#
# Paginar com "< T" pula os empatados que ficaram para tras. Paginar com
# "<= T" rele o grupo inteiro, e se ele for maior que a pagina a paginacao
# trava: nenhum item novo, e a consulta seguinte devolve os mesmos.
#
# ---------------------------------------------------------------------
# O ALGORITMO, como ele e de fato
#
#   1. abre um cursor e le uma pagina;
#   2. DRENA o resto do grupo do ultimo instante NO MESMO cursor — sem
#      reabrir, entao nao ha filtro envolvido nessa parte;
#   3. so entao reabre com "< T" ESTRITO.
#
# Uma versao anterior tinha uma variavel "inclusivo" e reabria com "<=".
# Era codigo morto: a primeira fronteira e nula e toda drenagem bem
# sucedida deixava a fronteira estrita. Removida, porque descricao que nao
# corresponde ao codigo e pior que ausencia de descricao.
#
# ---------------------------------------------------------------------
# Modo QUEBRADO, de proposito
#
# Os parametros SemDrenagem e FronteiraInclusiva reproduzem os dois
# defeitos que este algoritmo ja teve. Existem para os controles negativos
# serem TESTE EXECUTAVEL, e nao um numero que alguem anotou depois de
# editar o arquivo a mao.

function Invoke-PaginacaoPorCursor {
    param(
        # Abre a fonte. Recebe (fronteira, inclusivo); devolve um cursor.
        [scriptblock]$Abrir,
        # Le ate N linhas do cursor. Recebe (cursor, n).
        # Cada linha: objeto com .Id e .Quando
        [scriptblock]$Ler,
        # Fecha o cursor. Recebe (cursor).
        [scriptblock]$Fechar,

        [int]$TamanhoDaPagina = 50,
        [int]$MaxAberturas = 1000,
        [scriptblock]$AoLerPagina = $null,

        # --- defeitos reproduziveis, para os controles negativos ---
        [switch]$SemDrenagem,
        [switch]$FronteiraInclusiva
    )

    $vistos = @{}
    $fronteira = $null
    $aberturas = 0
    $total = 0

    while ($aberturas -lt $MaxAberturas) {
        $cursor = & $Abrir $fronteira ([bool]$FronteiraInclusiva)
        $aberturas++

        $pagina = & $Ler $cursor $TamanhoDaPagina
        if (-not $pagina -or $pagina.Count -eq 0) { & $Fechar $cursor; break }

        $novos = 0
        $ultimo = $null
        foreach ($linha in $pagina) {
            $ultimo = $linha.Quando
            if ($vistos.ContainsKey($linha.Id)) { continue }
            $vistos[$linha.Id] = $true
            $novos++
        }

        # DRENA o grupo do ultimo instante, no MESMO cursor. Para no
        # primeiro instante diferente SEM consumi-lo: essas linhas voltam
        # na consulta seguinte.
        $grupoCompleto = $SemDrenagem.IsPresent
        if (-not $SemDrenagem) {
            while ($true) {
                $extra = & $Ler $cursor $TamanhoDaPagina
                if (-not $extra -or $extra.Count -eq 0) { $grupoCompleto = $true; break }

                $saiu = $false
                foreach ($linha in $extra) {
                    if ($linha.Quando -ne $ultimo) { $saiu = $true; break }
                    if ($vistos.ContainsKey($linha.Id)) { continue }
                    $vistos[$linha.Id] = $true
                    $novos++
                }
                if ($saiu) { $grupoCompleto = $true; break }
            }
        }

        & $Fechar $cursor
        $total += $novos
        if ($AoLerPagina) { & $AoLerPagina $novos $ultimo }

        # Sem nada novo e sem ter drenado, nao ha como avancar com
        # seguranca: parar e melhor que pular o resto do grupo.
        if ($novos -eq 0 -and -not $grupoCompleto) { break }
        if ($novos -eq 0 -and $FronteiraInclusiva) { break }

        if (-not $grupoCompleto) { break }
        $fronteira = $ultimo
    }

    return [pscustomobject]@{ Lidos = $total; Aberturas = $aberturas }
}
