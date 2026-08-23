# Q1: paginacao por CURSOR com a Table, contra a caixa REAL.
#
# SOMENTE LEITURA.
#
# O algoritmo vive em paginacao.ps1, e e o MESMO que q1-cursor-teste.ps1
# exercita contra tabela sintetica. A unica diferenca entre os dois e de
# onde as linhas vem — antes eram algoritmos diferentes, e o teste passava
# provando uma versao melhor do que a implementada aqui.
#
# TRES ARMADILHAS, todas descobertas medindo, todas SILENCIOSAS:
#
#   1. FUSO. O filtro DASL interpreta a data como UTC; o ReceivedTime que a
#      tabela devolve vem LOCAL. Filtrar com hora local pulava uma janela
#      do tamanho do fuso em CADA fronteira: 803 de 1003, e a paginacao
#      ainda terminava cedo, parecendo ter acabado.
#
#   2. EMPATE. ReceivedTime nao e ordem total. Sem drenar o grupo da
#      fronteira, um empate maior que a pagina trava o avanco.
#
#   3. FRONTEIRA INCLUSIVA APOS DRENAR. Reabrir com "<=" depois de drenar
#      recomeca no mesmo grupo: nada e novo e a paginacao declara fim.
#      Tem de virar "<" estrito.

param([int]$PastaId = 6, [int]$TamanhoDaPagina = 50)

. (Join-Path $PSScriptRoot "paginacao.ps1")

$ol = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
$pasta = $ol.GetNamespace("MAPI").GetDefaultFolder($PastaId)

$colunas = @(
    "EntryID", "Subject", "SenderName", "ReceivedTime", "Size", "UnRead",
    "MessageClass",
    "http://schemas.microsoft.com/mapi/proptag/0x0E1B000B",
    "http://schemas.microsoft.com/mapi/proptag/0x300B0102",
    "http://schemas.microsoft.com/mapi/proptag/0x1035001E"
)

$i = $pasta.Items; $total = $i.Count
[void][Runtime.InteropServices.Marshal]::ReleaseComObject($i)

Write-Host "pasta: $($pasta.Name) | $total itens | pagina: $TamanhoDaPagina"
Write-Host ""
Write-Host "pagina | novos | ms"
Write-Host "-------|-------|-----"

$script:pagina = 0
$script:anterior = 0.0
$cronometro = [Diagnostics.Stopwatch]::StartNew()

$abrir = {
    param($fronteira, $inclusivo)

    $filtro = $null
    if ($null -ne $fronteira) {
        # UTC e cultura invariante. Hora local aqui perde mensagem.
        $op = if ($inclusivo) { "<=" } else { "<" }
        $quando = ([datetime]$fronteira).ToUniversalTime().ToString(
            "yyyy-MM-dd HH:mm:ss", [Globalization.CultureInfo]::InvariantCulture)
        $filtro = "@SQL=" + '"' + "urn:schemas:httpmail:datereceived" + '"' + " $op '$quando'"
    }

    $t = if ($filtro) { $pasta.GetTable($filtro) } else { $pasta.GetTable() }
    $t.Columns.RemoveAll()
    foreach ($c in $colunas) { [void]$t.Columns.Add($c) }
    $t.Sort("ReceivedTime", $true)
    return $t
}

$ler = {
    param($t, $n)
    $a = $t.GetArray($n)
    $linhas = @()
    if ($a) {
        for ($r = $a.GetLowerBound(0); $r -le $a.GetUpperBound(0); $r++) {
            $linhas += [pscustomobject]@{
                Id     = "$($a.GetValue($r, 0))"
                Quando = $a.GetValue($r, 3)
            }
        }
    }
    return ,$linhas
}

$fechar = { param($t) [void][Runtime.InteropServices.Marshal]::ReleaseComObject($t) }

$aoLer = {
    param($novos, $ultimo)
    $script:pagina++
    $agora = $cronometro.Elapsed.TotalMilliseconds
    Write-Host ("{0,6} | {1,5} | {2,3}" -f $script:pagina, $novos, [int]($agora - $script:anterior))
    $script:anterior = $agora
}

$r = Invoke-PaginacaoPorCursor -Abrir $abrir -Ler $ler -Fechar $fechar `
                               -TamanhoDaPagina $TamanhoDaPagina -AoLerPagina $aoLer
$cronometro.Stop()

Write-Host ""
Write-Host ("lidos: {0} de {1}   |   consultas: {2}" -f $r.Lidos, $total, $r.Consultas)
Write-Host ("tempo: {0} ms   =>   {1:N2} ms/item" -f `
    [int]$cronometro.Elapsed.TotalMilliseconds,
    ($cronometro.Elapsed.TotalMilliseconds / [Math]::Max($r.Lidos, 1)))

if ($r.Lidos -ne $total) {
    Write-Host ""
    Write-Host ("ATENCAO: faltaram {0} itens." -f ($total - $r.Lidos))
}

[void][Runtime.InteropServices.Marshal]::ReleaseComObject($pasta)
