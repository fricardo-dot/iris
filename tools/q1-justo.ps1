# Q1: comparacao JUSTA entre Table e iteracao.
#
# SOMENTE LEITURA.
#
# A primeira comparacao nao era equivalente, e a revisao pegou:
#   - o lado da iteracao lia Permission e ABRIA a colecao Attachments;
#   - o lado da Table nao pegava Permission, usava PR_HASATTACH, e ainda
#     trazia tres colunas a mais.
#
# Aqui os dois lados fazem o MESMO trabalho: as oito propriedades escalares
# que os dois conseguem entregar. Depois, mede-se o que cada extra custa,
# em separado.

param([int]$PastaId = 6, [int]$TamanhoDaPagina = 50, [int]$Execucoes = 3)

$ol = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
$pasta = $ol.GetNamespace("MAPI").GetDefaultFolder($PastaId)

function Media($lista) { ($lista | Measure-Object -Average).Average }

# ---------------------------------------------------------------
# A) COMUM: 8 escalares, sem Permission, sem abrir Attachments.
# ---------------------------------------------------------------
$comuns = @("EntryID","Subject","SenderName","ReceivedTime","Size","UnRead",
            "MessageClass","LastModificationTime")

$tabComum = @(); $iterComum = @()

for ($e = 1; $e -le $Execucoes; $e++) {
    # Table
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $t = $pasta.GetTable()
    $t.Columns.RemoveAll()
    foreach ($c in $comuns) { [void]$t.Columns.Add($c) }
    $t.Sort("ReceivedTime", $true)
    $a = $t.GetArray($TamanhoDaPagina)
    $n = 0
    for ($r = $a.GetLowerBound(0); $r -le $a.GetUpperBound(0); $r++) {
        $null = [pscustomobject]@{
            A=$a.GetValue($r,0); B=$a.GetValue($r,1); C=$a.GetValue($r,2); D=$a.GetValue($r,3)
            E=$a.GetValue($r,4); F=$a.GetValue($r,5); G=$a.GetValue($r,6); H=$a.GetValue($r,7)
        }
        $n++
    }
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($t)
    $tabComum += $sw.Elapsed.TotalMilliseconds / $n

    # Iteracao — mesmas 8, e tambem materializando o objeto.
    $sw.Restart()
    $items = $pasta.Items
    $items.Sort("[ReceivedTime]", $true)
    $n = 0
    for ($k = 1; $k -le $TamanhoDaPagina; $k++) {
        $m = $items.Item($k)
        $null = [pscustomobject]@{
            A=$m.EntryID; B=$m.Subject; C=$m.SenderName; D=$m.ReceivedTime
            E=$m.Size; F=$m.UnRead; G=$m.MessageClass; H=$m.LastModificationTime
        }
        $n++
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($m)
    }
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($items)
    $iterComum += $sw.Elapsed.TotalMilliseconds / $n
}

# ---------------------------------------------------------------
# B) O que cada EXTRA custa, medido em separado.
# ---------------------------------------------------------------
$custoAnexoIter = @(); $custoPermIter = @(); $custoTresColunas = @()

for ($e = 1; $e -le $Execucoes; $e++) {
    $items = $pasta.Items
    $items.Sort("[ReceivedTime]", $true)

    # MARGINAL, e nao total: buscar o item de novo a cada medicao incluiria
    # o custo do Items.Item(k) em todo "extra", inflando o resultado. Aqui o
    # item e buscado UMA vez e a diferenca e cronometrada por dentro.
    $swAnexo = [Diagnostics.Stopwatch]::new()
    $swPerm  = [Diagnostics.Stopwatch]::new()

    for ($k = 1; $k -le $TamanhoDaPagina; $k++) {
        $m = $items.Item($k)

        $swAnexo.Start()
        $x = $m.Attachments; $null = $x.Count
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x)
        $swAnexo.Stop()

        $swPerm.Start()
        try { $null = $m.Permission } catch {}
        $swPerm.Stop()

        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($m)
    }
    $custoAnexoIter += $swAnexo.Elapsed.TotalMilliseconds / $TamanhoDaPagina
    $custoPermIter  += $swPerm.Elapsed.TotalMilliseconds / $TamanhoDaPagina
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($items)

    # As tres colunas extras da Table
    $sw.Restart()
    $t = $pasta.GetTable()
    $t.Columns.RemoveAll()
    foreach ($c in ($comuns + @(
        "http://schemas.microsoft.com/mapi/proptag/0x0E1B000B",
        "http://schemas.microsoft.com/mapi/proptag/0x300B0102",
        "http://schemas.microsoft.com/mapi/proptag/0x1035001E"))) { [void]$t.Columns.Add($c) }
    $t.Sort("ReceivedTime", $true)
    $a = $t.GetArray($TamanhoDaPagina)
    $n = 0
    for ($r = $a.GetLowerBound(0); $r -le $a.GetUpperBound(0); $r++) {
        $null = [pscustomobject]@{
            A=$a.GetValue($r,0); B=$a.GetValue($r,1); C=$a.GetValue($r,2); D=$a.GetValue($r,3)
            E=$a.GetValue($r,4); F=$a.GetValue($r,5); G=$a.GetValue($r,6); H=$a.GetValue($r,7)
            I=$a.GetValue($r,8); J=$a.GetValue($r,9); K=$a.GetValue($r,10)
        }
        $n++
    }
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($t)
    $custoTresColunas += $sw.Elapsed.TotalMilliseconds / $n
}

$tc = Media $tabComum
$ic = Media $iterComum
$ca = Media $custoAnexoIter
$cp = Media $custoPermIter
$c3 = Media $custoTresColunas

Write-Output "MESMO TRABALHO — 8 escalares, sem Permission, sem abrir Attachments"
Write-Output ("  Table    : {0,6:N2} ms/item" -f $tc)
Write-Output ("  Iteracao : {0,6:N2} ms/item" -f $ic)
Write-Output ("  ganho    : {0,6:N1}x" -f ($ic / $tc))
Write-Output ""
Write-Output "CUSTO DE CADA EXTRA, em separado"
Write-Output ("  iteracao: abrir Attachments + Count : {0,6:N2} ms/item  (marginal)" -f $ca)
Write-Output ("  iteracao: ler Permission            : {0,6:N2} ms/item  (marginal)" -f $cp)
Write-Output ("  Table   : 11 colunas em vez de 8    : {0,6:N2} ms/item  (delta {1:N2})" -f $c3, ($c3 - $tc))
Write-Output ""
Write-Output ("Iteracao com o DTO completo (8 + anexos + permissao): {0,6:N2} ms/item" -f ($ic + $ca + $cp))
Write-Output ("Table com as 11 colunas:                              {0,6:N2} ms/item" -f $c3)
Write-Output ("Ganho no caminho de PRODUCAO candidato:               {0,6:N1}x" -f (($ic + $ca + $cp) / $c3))

[void][Runtime.InteropServices.Marshal]::ReleaseComObject($pasta)
