# Q2: remover a copia que o experimento deixou no Lixo Eletronico.
#
# ESCREVE: um Delete() SOFT, que move para Itens Excluidos. Reversivel.
# NUNCA DeleteAllPermanently, nunca esvaziar nada.
#
# Situacao: o Copy() do controle negativo deixou uma copia em cada pasta.
# As duas sao artefato meu.
#
#   - No Lixo Eletronico ha 2 itens identicos. Removo UM, por RecordKey
#     exata, para nao encostar em nenhum outro item.
#   - Em Itens Excluidos a copia JA esta em Itens Excluidos. Delete() ali
#     pode ser permanente, e exclusao permanente sem consentimento explicito
#     esta proibida neste projeto. Fica onde esta, junto dos outros
#     artefatos de spike, e registrada.
#
# Os dois itens do Lixo sao copias identicas da mesma mensagem de lixo
# eletronico, entao qual dos dois sai e indiferente para o usuario.

$ErrorActionPreference = "Stop"
$PT = "http://schemas.microsoft.com/mapi/proptag/"
$ALVO_SK = "6e27c63c795e034da35ae7e823bed6e1"

$ol = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
$ns = $ol.GetNamespace("MAPI")

$lixo = $ns.GetDefaultFolder(23)
Write-Host ("Lixo Eletronico antes: {0} itens" -f $lixo.Items.Count)

# Localizar por SearchKey, sem depender de assunto.
$t = $lixo.GetTable()
$cols = $t.Columns
try {
    $cols.RemoveAll()
    [void]$cols.Add("EntryID")
    [void]$cols.Add("Subject")
    [void]$cols.Add(($PT + "0x300B0102"))
    [void]$cols.Add(($PT + "0x0FF90102"))
} finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($cols) }

$cands = @()
while (-not $t.EndOfTable) {
    $a = $t.GetArray(200)
    for ($r = $a.GetLowerBound(0); $r -le $a.GetUpperBound(0); $r++) {
        $sk = $a.GetValue($r, 2)
        if (-not ($sk -is [byte[]])) { continue }
        $hex = (($sk | ForEach-Object { $_.ToString("x2") }) -join "")
        if ($hex -ne $ALVO_SK) { continue }
        $rk = $a.GetValue($r, 3)
        $cands += [pscustomobject]@{
            Id = "$($a.GetValue($r,0))"
            Assunto = "$($a.GetValue($r,1))"
            Rk = (($rk | ForEach-Object { $_.ToString("x2") }) -join "")
        }
    }
}
[void][Runtime.InteropServices.Marshal]::ReleaseComObject($t)

Write-Host ("itens com a SearchKey do experimento: {0}" -f $cands.Count)
foreach ($c in $cands) {
    Write-Host ("   RecordKey ...{0}  |  {1}" -f $c.Rk.Substring($c.Rk.Length - 16),
                $c.Assunto.Substring(0, [Math]::Min(44, $c.Assunto.Length)))
}

if ($cands.Count -ne 2) {
    Write-Host ""
    Write-Host "ABORTADO: esperava exatamente 2. Nao mexo em nada sem certeza."
    exit 1
}

# O de RecordKey MENOR foi alocado antes. Tanto faz qual sai — sao copias
# identicas —, mas escolher por criterio fixo torna o script repetivel.
$sai = ($cands | Sort-Object Rk)[0]
Write-Host ""
Write-Host ("removendo (soft) o de RecordKey ...{0}" -f $sai.Rk.Substring($sai.Rk.Length - 16))

$item = $ns.GetItemFromID($sai.Id)
try {
    $item.Delete()      # soft: vai para Itens Excluidos
} finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($item) }

$depois = $ns.GetDefaultFolder(23)
Write-Host ("Lixo Eletronico depois: {0} itens" -f $depois.Items.Count)
[void][Runtime.InteropServices.Marshal]::ReleaseComObject($depois)
[void][Runtime.InteropServices.Marshal]::ReleaseComObject($lixo)
