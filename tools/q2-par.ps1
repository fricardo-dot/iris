# Q2, parte 3: o par que colide e o mesmo item, ou dois?
#
# SOMENTE LEITURA.
#
# O ORACULO PRIMEIRO. Antes de olhar qualquer resultado eu escrevo o que
# significaria cada resposta:
#
#   Se os dois forem o ENVIADO e o RECEBIDO da mesma mensagem, entao eles
#   sao itens DISTINTOS que compartilham Message-ID E SearchKey. Unir os
#   dois e FALSO POSITIVO: o Outlook mostra os dois, o usuario tem estado
#   separado em cada um, e fundir faria o lido/triado de um vazar no outro.
#
#   Se for a mesma mensagem copiada, unir estaria certo.
#
# O discriminador: item enviado tem PR_CLIENT_SUBMIT_TIME e nao tem
# PR_MESSAGE_DELIVERY_TIME. Recebido e o inverso. Nao da para os dois serem
# a mesma manifestacao se um foi submetido e o outro entregue.

$PT = "http://schemas.microsoft.com/mapi/proptag/"
$props = [ordered]@{
    "SearchKey"      = "0x300B0102"
    "MessageID"      = "0x1035001E"
    "RecordKey"      = "0x0FF90102"
    "SubmitTime"     = "0x00390040"   # PR_CLIENT_SUBMIT_TIME  (enviado)
    "DeliveryTime"   = "0x0E060040"   # PR_MESSAGE_DELIVERY_TIME (recebido)
    "MsgFlags"       = "0x0E070003"   # PR_MESSAGE_FLAGS
    "TransportHdrs"  = "0x007D001E"   # PR_TRANSPORT_MESSAGE_HEADERS
    "ConvIndex"      = "0x00710102"
    "ConvTopic"      = "0x0070001E"
}

$ol = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
$ns = $ol.GetNamespace("MAPI")

function Achar($pastaId, [string]$rotulo, [string]$prefixo) {
    $pasta = $ns.GetDefaultFolder($pastaId)
    $t = $pasta.GetTable()
    $t.Columns.RemoveAll()
    [void]$t.Columns.Add("EntryID")
    [void]$t.Columns.Add("Subject")
    $achados = @()
    while (-not $t.EndOfTable) {
        $a = $t.GetArray(200)
        for ($r = $a.GetLowerBound(0); $r -le $a.GetUpperBound(0); $r++) {
            $s = "$($a.GetValue($r,1))"
            if ($s.StartsWith($prefixo, [StringComparison]::Ordinal)) {
                $achados += [pscustomobject]@{ Pasta=$rotulo; Id="$($a.GetValue($r,0))"; Assunto=$s }
            }
        }
    }
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($t)
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pasta)
    return ,$achados
}

$alvos = @()
$alvos += Achar 5 "Enviados"  "[IRIS-SPIKE-C]"
$alvos += Achar 3 "Excluidos" "[IRIS-SPIKE-C]"

Write-Host "itens do par: $($alvos.Count)"
Write-Host ""

function Hex($v) {
    if ($null -eq $v) { return "(nulo)" }
    if ($v -is [array]) { return (($v | ForEach-Object { $_.ToString("x2") }) -join "") }
    return "$v"
}

$linhas = @{}
foreach ($alvo in $alvos) {
    $item = $ns.GetItemFromID($alvo.Id)
    $pa = $item.PropertyAccessor
    $col = [ordered]@{}
    $col["Pasta"]    = $alvo.Pasta
    $col["Classe"]   = $item.MessageClass
    $col["Tamanho"]  = $item.Size
    $col["EntryID"]  = $alvo.Id.Substring($alvo.Id.Length - 16)
    foreach ($k in $props.Keys) {
        try {
            $v = $pa.GetProperty($PT + $props[$k])
            if ($k -eq "TransportHdrs") { $col[$k] = "presente ($($v.Length) chars)" }
            else { $col[$k] = Hex $v }
        } catch {
            $col[$k] = "(ausente)"
        }
    }
    $linhas[$alvo.Pasta] = $col
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pa)
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($item)
}

$chaves = @("Pasta","Classe","Tamanho","EntryID") + @($props.Keys)
$pastas = @($linhas.Keys)

Write-Host ("{0,-14} | {1,-42} | {2,-42} | igual?" -f "propriedade", $pastas[0], $pastas[1])
Write-Host ("-" * 118)
foreach ($k in $chaves) {
    $a = "$($linhas[$pastas[0]][$k])"
    $b = "$($linhas[$pastas[1]][$k])"
    $mostraA = if ($a.Length -gt 42) { $a.Substring(0,39) + "..." } else { $a }
    $mostraB = if ($b.Length -gt 42) { $b.Substring(0,39) + "..." } else { $b }
    $ig = if ($a -eq $b) { "SIM" } else { "nao" }
    Write-Host ("{0,-14} | {1,-42} | {2,-42} | {3}" -f $k, $mostraA, $mostraB, $ig)
}
