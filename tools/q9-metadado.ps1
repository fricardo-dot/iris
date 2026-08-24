# Quanto custa, em disco, guardar SO o metadado de uma mensagem?
#
# SOMENTE LEITURA. Nada e gravado na caixa; o CSV de amostra vai para o
# TEMP e e apagado no fim.
#
# ------------------------------------------------------------------
# POR QUE ESTE NUMERO DECIDE A Q9'
#
# A caixa esta em Modo Cache. O OST guarda CONTEUDO — corpo, anexo,
# imagem — e por isso um mes custa 1,5 GB. O cache do Iris guardaria
# METADADO: quem, quando, assunto, chaves, estado de leitura.
#
# Se o metadado for barato o suficiente, o Iris pode ACUMULAR historico
# que o OST descarta, e devolver ao usuario o alcance que a restricao de
# memoria tirou. Se nao for, so resta ESPELHAR a janela.
#
# Eu tinha ESTIMADO ~600 bytes por item. Estimativa nao e medicao, e o
# EntryID sozinho tem 140 caracteres. Aqui se mede.
#
# O que se mede e o tamanho SERIALIZADO dos campos do MailSummary, mais os
# campos de correlacao que a Q2 mostrou serem necessarios. Nao inclui
# indice do SQLite nem overhead de pagina — por isso o numero final e
# apresentado com margem.

$ErrorActionPreference = "Stop"
$PT = "http://schemas.microsoft.com/mapi/proptag/"
$ol = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
$ns = $ol.GetNamespace("MAPI")
function Solta($o) { if ($null -ne $o) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } }

$in = $ns.GetDefaultFolder(6)
$t = $in.GetTable()
$cols = $t.Columns
try {
    $cols.RemoveAll()
    foreach ($c in @(
        ($PT + "0x66700102"),   # EntryID longo
        "Subject", "SenderName", "ReceivedTime", "Size", "UnRead", "MessageClass",
        ($PT + "0x0E1B000B"),   # PR_HASATTACH
        ($PT + "0x300B0102"),   # SearchKey        -> correlacao (Q2)
        ($PT + "0x1035001E"),   # Message-ID       -> correlacao (Q2)
        ($PT + "0x30080040")    # LastModification -> Q3
    )) { $x = $cols.Add($c); Solta $x }
} finally { Solta $cols }

$campos = @{}
$n = 0
$storeId = $in.StoreID
while (-not $t.EndOfTable) {
    $a = $t.GetArray(200)
    for ($r = $a.GetLowerBound(0); $r -le $a.GetUpperBound(0); $r++) {
        $n++
        for ($c = 0; $c -le 10; $c++) {
            $v = $a.GetValue($r, $c)
            $len = 0
            if ($null -ne $v) {
                if ($v -is [byte[]]) { $len = $v.Length * 2 }        # hex
                elseif ($v -is [string]) { $len = [Text.Encoding]::UTF8.GetByteCount($v) }
                elseif ($v -is [datetime]) { $len = 8 }
                elseif ($v -is [bool]) { $len = 1 }
                else { $len = [Text.Encoding]::UTF8.GetByteCount("$v") }
            }
            $campos[$c] = [int64]$campos[$c] + $len
        }
    }
}
Solta $t

$nomes = @("EntryID(longo)","Subject","SenderName","ReceivedTime","Size","UnRead",
           "MessageClass","HasAttach","SearchKey","Message-ID","LastModif")

Write-Host ("amostra: {0} itens da Caixa de Entrada" -f $n)
Write-Host ""
Write-Host ("{0,-16} | {1,10} | {2,8}" -f "campo", "total (B)", "media")
Write-Host ("-" * 42)
$soma = 0
for ($c = 0; $c -le 10; $c++) {
    $tot = [int64]$campos[$c]
    $soma += $tot
    Write-Host ("{0,-16} | {1,10:N0} | {2,8:N1}" -f $nomes[$c], $tot, ($tot / [Math]::Max($n,1)))
}
$storeBytes = [Text.Encoding]::UTF8.GetByteCount($storeId)
Write-Host ("-" * 42)
Write-Host ("{0,-16} | {1,10:N0} | {2,8:N1}" -f "SOMA dos campos", $soma, ($soma / [Math]::Max($n,1)))
Write-Host ""
Write-Host ("StoreID (uma vez por STORE, nao por item): {0} bytes" -f $storeBytes)
Write-Host "  -> guardar o StoreID em cada linha custaria mais que o resto junto."
Write-Host "     Ele vira tabela de stores, com um id interno pequeno. E o I1."
Write-Host ""

$media = $soma / [Math]::Max($n,1)
Write-Host ("=" * 60)
Write-Host ("MEDIA POR MENSAGEM: {0:N0} bytes de campo" -f $media)
Write-Host ("=" * 60)
foreach ($f in @(1.5, 2.0, 3.0)) {
    Write-Host ("  com fator {0:N1}x de overhead (indice, pagina, folga):" -f $f)
    foreach ($q in @(1004, 3000, 17668, 50000)) {
        Write-Host ("     {0,6} msgs -> {1,8:N1} MB" -f $q, ($q * $media * $f / 1MB))
    }
}
Write-Host ""
Write-Host ("Para comparar: o OST com UM MES de CONTEUDO ocupa 1.511 MB.")

Solta $in
