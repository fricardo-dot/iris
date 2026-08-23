# Q1: paginacao por CURSOR com a Table — versao correta.
#
# SOMENTE LEITURA.
#
# DUAS ARMADILHAS, as duas descobertas medindo, e as duas SILENCIOSAS:
#
#   1. FUSO. O filtro DASL interpreta a string de data como UTC; o
#      ReceivedTime que a tabela devolve vem em hora LOCAL. Usar a hora
#      local no filtro pulava uma janela do tamanho do offset do fuso em
#      CADA fronteira — 200 de 1003 itens perdidos, e a paginacao ainda
#      terminava cedo, parecendo ter acabado.
#
#   2. EMPATE. O filtro estrito "<" pula itens com o MESMO segundo do
#      ultimo da pagina anterior. E "<=" com deduplicacao NAO basta: se o
#      grupo empatado for maior que a pagina, a consulta seguinte pode
#      devolver os MESMOS itens, nenhum ser novo, e a paginacao declarar
#      fim — perdendo o resto do grupo em silencio.
#
#      A saida e DRENAR o grupo da fronteira antes de avancar. Testado
#      contra tabela sintetica em q1-cursor-teste.ps1: sem drenar, o
#      cenario "tudo no mesmo segundo" perde 150 de 200.
#
# ReceivedTime nao e ordem total. Como cursor, ele precisa de desempate.

param([int]$PastaId = 6, [int]$TamanhoDaPagina = 50, [int]$MaxPaginas = 60)

$ol = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
$pasta = $ol.GetNamespace("MAPI").GetDefaultFolder($PastaId)

$colunas = @(
    "EntryID", "Subject", "SenderName", "ReceivedTime", "Size", "UnRead",
    "MessageClass", "LastModificationTime",
    "http://schemas.microsoft.com/mapi/proptag/0x0E1B000B",
    "http://schemas.microsoft.com/mapi/proptag/0x300B0102",
    "http://schemas.microsoft.com/mapi/proptag/0x1035001E"
)

$i = $pasta.Items; $total = $i.Count
[void][Runtime.InteropServices.Marshal]::ReleaseComObject($i)

Write-Output "pasta: $($pasta.Name) | $total itens | pagina: $TamanhoDaPagina"
Write-Output ""
Write-Output "pagina | novos | repetidos | ms  | ms/item"
Write-Output "-------|-------|-----------|-----|--------"

$vistos = @{}
$fronteira = $null
$totalMs = 0
$novosTotal = 0
$repetidosTotal = 0

for ($p = 1; $p -le $MaxPaginas; $p++) {
    $sw = [Diagnostics.Stopwatch]::StartNew()

    # <= e nao <, para nao pular empate. E a data vai em UTC.
    $filtro = if ($fronteira) {
        "@SQL=" + '"' + "urn:schemas:httpmail:datereceived" + '"' + " <= '" +
            $fronteira.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss") + "'"
    } else { $null }

    $t = if ($filtro) { $pasta.GetTable($filtro) } else { $pasta.GetTable() }
    $t.Columns.RemoveAll()
    foreach ($c in $colunas) { [void]$t.Columns.Add($c) }
    $t.Sort("ReceivedTime", $true)

    # SomenteInstante: durante a drenagem, so consome linhas do grupo
    # empatado. Sem isso a drenagem marcava como vistos itens de FORA do
    # grupo e os descartava — a pagina seguinte os encontrava repetidos,
    # concluia que nao havia nada novo e parava com 50 de 1003.
    function Consumir($linhas, $SomenteInstante) {
        $res = @{ Novos = 0; Repetidos = 0; Ultima = $null; Qtd = 0; Saiu = $false }
        if (-not $linhas) { return $res }
        for ($r = $linhas.GetLowerBound(0); $r -le $linhas.GetUpperBound(0); $r++) {
            $quando = $linhas.GetValue($r, 3)
            if ($null -ne $SomenteInstante -and $quando -ne $SomenteInstante) {
                $res.Saiu = $true
                break
            }
            $res.Qtd++
            $id = "$($linhas.GetValue($r, 0))"
            $res.Ultima = $quando
            if ($vistos.ContainsKey($id)) { $res.Repetidos++; continue }
            $vistos[$id] = $true
            $null = [pscustomobject]@{
                EntryID   = $id
                Subject   = $linhas.GetValue($r, 1)
                Sender    = $linhas.GetValue($r, 2)
                Recebido  = $res.Ultima
                Tamanho   = $linhas.GetValue($r, 4)
                NaoLida   = $linhas.GetValue($r, 5)
                Classe    = $linhas.GetValue($r, 6)
                Modif     = $linhas.GetValue($r, 7)
                TemAnexo  = $linhas.GetValue($r, 8)
                SearchKey = $linhas.GetValue($r, 9)
                MsgId     = $linhas.GetValue($r, 10)
            }
            $res.Novos++
        }
        return $res
    }

    $res = Consumir $t.GetArray($TamanhoDaPagina) $null
    $novos = $res.Novos; $repetidos = $res.Repetidos; $ultima = $res.Ultima

    # DRENA o grupo da fronteira: enquanto as linhas seguintes tiverem o
    # MESMO instante, elas fazem parte do grupo que a pagina cortou ao
    # meio. Avancar a fronteira sem esvaziar o grupo perde o resto dele.
    $drenou = $false
    if ($null -ne $ultima) {
        while ($true) {
            $extra = Consumir $t.GetArray($TamanhoDaPagina) $ultima
            $novos += $extra.Novos; $repetidos += $extra.Repetidos
            if ($extra.Novos -gt 0) { $drenou = $true }

            # Saiu do grupo, ou a tabela acabou.
            if ($extra.Saiu -or $extra.Qtd -eq 0) { break }
        }
    }

    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($t)
    $ms = $sw.Elapsed.TotalMilliseconds
    $totalMs += $ms

    if ($novos -eq 0 -and -not $drenou) { Write-Output "  (fim)"; break }

    $novosTotal += $novos; $repetidosTotal += $repetidos
    "{0,6} | {1,5} | {2,9} | {3,3} | {4,7}" -f $p, $novos, $repetidos, [int]$ms, `
        [math]::Round($ms / [Math]::Max($novos,1), 2)

    $fronteira = [datetime]$ultima
}

Write-Output ""
Write-Output ("lidos: {0} de {1}   |   releituras por empate: {2}" -f $novosTotal, $total, $repetidosTotal)
Write-Output ("tempo: {0} ms   =>   {1:N2} ms/item" -f [int]$totalMs, ($totalMs / [Math]::Max($novosTotal,1)))
if ($novosTotal -ne $total) {
    Write-Output ""
    Write-Output ("ATENCAO: faltaram {0} itens. A paginacao ainda perde." -f ($total - $novosTotal))
}

[void][Runtime.InteropServices.Marshal]::ReleaseComObject($pasta)
