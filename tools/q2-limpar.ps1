# Q2: conferir onde os itens do experimento pararam, e limpar.
#
# ESCREVE: so apaga a pasta temporaria, e so se ela estiver vazia.
# Se houver item la, MOVE de volta antes — nunca apaga item.
#
# Existe porque o q2-move.ps1 terminou dizendo "itens restantes: 2" depois
# de reportar os dois itens de volta na origem. Ou a contagem estava velha,
# ou sobrou coisa. Contagem em cache nao serve para decidir apagar pasta.

$ErrorActionPreference = "Stop"
$PT = "http://schemas.microsoft.com/mapi/proptag/"

$ol = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
$ns = $ol.GetNamespace("MAPI")

$raiz = $ns.GetDefaultFolder(6).Parent
$excluidos = $ns.GetDefaultFolder(3)
$lixo = $ns.GetDefaultFolder(23)

$temp = $null
foreach ($f in $raiz.Folders) {
    if ($f.Name -eq "Iris Q2 (temp)") { $temp = $f; break }
}
if ($null -eq $temp) {
    Write-Host "pasta 'Iris Q2 (temp)' nao existe. Nada a limpar."
    exit 0
}

$itens = $temp.Items
Write-Host ("Iris Q2 (temp) contem {0} item(ns):" -f $itens.Count)
Write-Host ""

$restantes = @()
for ($i = 1; $i -le $itens.Count; $i++) {
    $it = $itens.Item($i)
    $s = "$($it.Subject)"
    if ($s.Length -gt 50) { $s = $s.Substring(0,50) + "..." }
    Write-Host ("  [{0}] {1}" -f $i, $s)
    $restantes += $it
}

if ($restantes.Count -gt 0) {
    Write-Host ""
    Write-Host "Devolvendo cada um para a pasta de origem pelo ASSUNTO."
    foreach ($it in $restantes) {
        $s = "$($it.Subject)"
        $destino = if ($s.StartsWith("[IRIS-SPIKE-C]", [StringComparison]::Ordinal)) { $excluidos } else { $lixo }
        Write-Host ("  -> {0} : {1}" -f $destino.Name, $s.Substring(0, [Math]::Min(40, $s.Length)))
        $movido = $it.Move($destino)
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($movido)
    }
}
foreach ($it in $restantes) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($it) }
[void][Runtime.InteropServices.Marshal]::ReleaseComObject($itens)

# Reler do zero: contagem em cache foi o que gerou a duvida.
$novos = $temp.Items
$sobrou = $novos.Count
[void][Runtime.InteropServices.Marshal]::ReleaseComObject($novos)

Write-Host ""
Write-Host ("apos a devolucao, Iris Q2 (temp) tem {0} item(ns)" -f $sobrou)
if ($sobrou -eq 0) {
    $temp.Delete()
    Write-Host "pasta APAGADA (soft: vai para Itens Excluidos)."
} else {
    Write-Host "NAO apaguei: ainda ha item la dentro."
}
[void][Runtime.InteropServices.Marshal]::ReleaseComObject($temp)

Write-Host ""
Write-Host "CONFERENCIA FINAL das pastas de origem:"
Write-Host ("  Itens Excluidos : {0} itens" -f $ns.GetDefaultFolder(3).Items.Count)
Write-Host ("  Lixo Eletronico : {0} itens" -f $ns.GetDefaultFolder(23).Items.Count)
