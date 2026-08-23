# Q2, rodada 2: conferir a Table contra o PropertyAccessor, e o corpus
# contra ele mesmo.
#
# SOMENTE LEITURA.
#
# Duas duvidas que precisam morrer antes de qualquer numero virar requisito
# de arquitetura:
#
#   1. MARSHALING. Tudo que eu afirmei sai de Table.GetArray com colunas de
#      proptag binario. Se o marshaling devolver algo diferente do valor
#      real da propriedade, os 2281 "distintos" nao valem nada. Confiro
#      contra PropertyAccessor.GetProperty, que e outro caminho.
#
#   2. DERIVA. Table nao e snapshot. Se a caixa mudou durante a varredura,
#      o corpus e uma mistura de dois instantes. Leio duas vezes e comparo.

$ErrorActionPreference = "Stop"
$PT = "http://schemas.microsoft.com/mapi/proptag/"
$props = @(
    @{ Nome = "SearchKey"; Dasl = $PT + "0x300B0102" },
    @{ Nome = "RecordKey"; Dasl = $PT + "0x0FF90102" },
    @{ Nome = "MessageID"; Dasl = $PT + "0x1035001E" }
)

$ol = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
$ns = $ol.GetNamespace("MAPI")

function Hex($v) {
    if ($null -eq $v) { return $null }
    if ($v -is [byte[]]) {
        if ($v.Length -eq 0) { return $null }
        return (($v | ForEach-Object { $_.ToString("x2") }) -join "")
    }
    if ($v -is [string]) { $s = $v.Trim(); if ($s -eq "") { return $null }; return $s }
    return "ANOMALIA:$($v.GetType().Name)"
}

function LerPorTabela($pastaId) {
    $pasta = $ns.GetDefaultFolder($pastaId)
    try {
        $t = $pasta.GetTable()
        try {
            $cols = $t.Columns
            try {
                $cols.RemoveAll()
                [void]$cols.Add("EntryID")
                foreach ($p in $props) { [void]$cols.Add($p.Dasl) }
            } finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($cols) }

            $r = @{}
            while (-not $t.EndOfTable) {
                $a = $t.GetArray(200)
                for ($i = $a.GetLowerBound(0); $i -le $a.GetUpperBound(0); $i++) {
                    $id = "$($a.GetValue($i, 0))"
                    $r[$id] = @{
                        SearchKey = Hex $a.GetValue($i, 1)
                        RecordKey = Hex $a.GetValue($i, 2)
                        MessageID = Hex $a.GetValue($i, 3)
                    }
                }
            }
            return $r
        } finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($t) }
    } finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pasta) }
}

# ------------------------------------------------------------------
Write-Host "1. TABLE x PROPERTYACCESSOR"
Write-Host ("-" * 60)

$conferidos = 0; $divergentes = 0; $semAcesso = 0
$exemplos = @()

foreach ($idDaPasta in @(6, 5, 16, 3)) {
    $tab = LerPorTabela $idDaPasta
    # amostra estratificada: 15 por pasta, espalhados
    $ids = @($tab.Keys)
    $passo = [Math]::Max(1, [int]($ids.Count / 15))
    for ($k = 0; $k -lt $ids.Count; $k += $passo) {
        $id = $ids[$k]
        $item = $null
        try { $item = $ns.GetItemFromID($id) } catch { $semAcesso++; continue }
        try {
            $pa = $item.PropertyAccessor
            try {
                foreach ($p in $props) {
                    $viaPa = $null
                    try { $viaPa = Hex $pa.GetProperty($p.Dasl) } catch { $viaPa = $null }
                    $viaTab = $tab[$id][$p.Nome]
                    $conferidos++
                    if ($viaPa -ne $viaTab) {
                        $divergentes++
                        if ($exemplos.Count -lt 5) {
                            $exemplos += "$($p.Nome): tabela=[$viaTab] acessor=[$viaPa]"
                        }
                    }
                }
            } finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pa) }
        } finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($item) }
    }
}

Write-Host "  valores conferidos : $conferidos"
Write-Host "  divergentes        : $divergentes"
Write-Host "  itens sem acesso   : $semAcesso"
foreach ($e in $exemplos) { Write-Host "     $e" }
if ($divergentes -gt 0) {
    Write-Host "  ATENCAO: os dois caminhos discordam. Os numeros do corpus nao valem."
}

# ------------------------------------------------------------------
Write-Host ""
Write-Host "2. DERIVA: duas leituras seguidas da mesma pasta"
Write-Host ("-" * 60)

foreach ($idDaPasta in @(6, 3)) {
    $a = LerPorTabela $idDaPasta
    $b = LerPorTabela $idDaPasta

    $soEmA = @($a.Keys | Where-Object { -not $b.ContainsKey($_) })
    $soEmB = @($b.Keys | Where-Object { -not $a.ContainsKey($_) })
    $mudou = 0
    foreach ($id in $a.Keys) {
        if (-not $b.ContainsKey($id)) { continue }
        foreach ($p in $props) {
            if ($a[$id][$p.Nome] -ne $b[$id][$p.Nome]) { $mudou++ }
        }
    }
    Write-Host ("  pasta {0}: {1} -> {2} itens | sumiram {3} | surgiram {4} | chaves mudadas {5}" -f `
        $idDaPasta, $a.Count, $b.Count, $soEmA.Count, $soEmB.Count, $mudou)
}
Write-Host ""
Write-Host "Deriva diferente de zero nao invalida sozinha: significa que o corpus"
Write-Host "e uma mistura de instantes, e que numero exato nao e afirmavel."
