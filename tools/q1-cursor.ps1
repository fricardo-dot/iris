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
#      ultimo da pagina anterior. Aqui sao poucos (6 em 1003), mas "poucos"
#      nao e "nenhum", e uma mensagem sumida e uma mensagem sumida.
#      A saida e "<=" com deduplicacao por EntryID, aceitando reler alguns.
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

    $linhas = $t.GetArray($TamanhoDaPagina)
    $novos = 0; $repetidos = 0; $ultima = $null

    if ($linhas) {
        for ($r = $linhas.GetLowerBound(0); $r -le $linhas.GetUpperBound(0); $r++) {
            $id = "$($linhas.GetValue($r, 0))"
            $ultima = $linhas.GetValue($r, 3)
            if ($vistos.ContainsKey($id)) { $repetidos++; continue }
            $vistos[$id] = $true
            $null = [pscustomobject]@{
                EntryID   = $id
                Subject   = $linhas.GetValue($r, 1)
                Sender    = $linhas.GetValue($r, 2)
                Recebido  = $ultima
                Tamanho   = $linhas.GetValue($r, 4)
                NaoLida   = $linhas.GetValue($r, 5)
                Classe    = $linhas.GetValue($r, 6)
                Modif     = $linhas.GetValue($r, 7)
                TemAnexo  = $linhas.GetValue($r, 8)
                SearchKey = $linhas.GetValue($r, 9)
                MsgId     = $linhas.GetValue($r, 10)
            }
            $novos++
        }
    }
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($t)
    $ms = $sw.Elapsed.TotalMilliseconds
    $totalMs += $ms

    # Nenhum item novo com a fronteira andando significa fim de verdade.
    if ($novos -eq 0) { Write-Output "  (fim)"; break }

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
