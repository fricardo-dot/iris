# O ALGORITMO de paginacao por cursor. Um lugar so.
#
# Existe porque o teste sintetico anterior testava um algoritmo DIFERENTE
# do script real: o teste avancava com AddSeconds(-1), que e "< T", e o
# real reabria com "<= T". Os cenarios passavam provando uma versao melhor
# do que a implementada — teste que nao testa o que diz.
#
# Agora o real e o teste chamam ESTA funcao. A unica diferenca entre eles e
# de onde as linhas vem.
#
# ---------------------------------------------------------------------
# O problema que o algoritmo resolve
#
# ReceivedTime nao e ordem total: varios itens compartilham o mesmo
# instante. Paginar com "< T" pula os empatados; paginar com "<= T" relê o
# grupo inteiro, e se ele for maior que a pagina a paginacao trava — nenhum
# item novo, e a consulta seguinte devolve os mesmos.
#
# A saida tem tres partes, e falta UMA delas ja perde mensagem:
#   1. reabrir com "<=" para nao pular empatado;
#   2. DRENAR o grupo da fronteira antes de avancar;
#   3. depois de drenado, avancar com "<" ESTRITO — senao a consulta
#      seguinte recomeca no mesmo grupo e a paginacao declara fim.
#
# A parte 3 e a que faltava, e ela so aparece quando o grupo empatado e
# maior que a pagina.

function Invoke-PaginacaoPorCursor {
    param(
        # Abre a fonte. Recebe (fronteira, inclusivo) e devolve um cursor.
        [scriptblock]$Abrir,
        # Le ate N linhas do cursor. Recebe (cursor, n).
        # Cada linha: objeto com .Id e .Quando
        [scriptblock]$Ler,
        # Fecha o cursor. Recebe (cursor).
        [scriptblock]$Fechar,

        [int]$TamanhoDaPagina = 50,
        [int]$MaxConsultas = 1000,
        [scriptblock]$AoLerPagina = $null
    )

    $vistos = @{}
    $fronteira = $null
    $inclusivo = $true
    $consultas = 0
    $total = 0

    while ($consultas -lt $MaxConsultas) {
        $cursor = & $Abrir $fronteira $inclusivo
        $consultas++

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

        # DRENA o grupo da fronteira. Para no primeiro instante diferente
        # SEM consumi-lo: essas linhas voltam na consulta seguinte.
        $grupoCompleto = $false
        while ($consultas -lt $MaxConsultas) {
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

        & $Fechar $cursor
        $total += $novos
        if ($AoLerPagina) { & $AoLerPagina $novos $ultimo }

        # Grupo drenado: a fronteira anda, e agora ESTRITA. Sem isto a
        # consulta seguinte recomeca no mesmo grupo, nada e novo, e a
        # paginacao declara fim com itens mais antigos por ler.
        if ($grupoCompleto) {
            $fronteira = $ultimo
            $inclusivo = $false
        } else {
            # Nao deu para provar que o grupo acabou (limite de consultas).
            # Nao avanca: avancar aqui pularia o resto do grupo.
            break
        }
    }

    return [pscustomobject]@{ Lidos = $total; Consultas = $consultas }
}
