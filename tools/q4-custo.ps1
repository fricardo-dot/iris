# Q4, parte 1: custo de enumerar CHAVE + LastModificationTime.
#
# SOMENTE LEITURA.
#
# A Q4 pergunta "quanto custa obter chave + LastModificationTime de todos
# os itens de uma pasta, pelo caminho que a Q1 indicar" — e a Q1 indicou
# Table.
#
# Isto NAO e a pergunta principal da Q4 (que e consistencia sob mutacao),
# mas e a que da para responder sem escrever nada, e ela calibra a
# seguinte: se enumerar a caixa inteira for barato, "exigir DUAS
# observacoes completas e compativeis" (opcao 2 da Q4) fica viavel; se for
# caro, so sobram as opcoes 1 e 3.
#
# Tres execucoes por pasta. Uma medicao so nao distingue custo de ruido.

$ErrorActionPreference = "Stop"
$PT = "http://schemas.microsoft.com/mapi/proptag/"
$P_LASTMOD = $PT + "0x30080040"   # PR_LAST_MODIFICATION_TIME

$ol = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
$ns = $ol.GetNamespace("MAPI")

function Enumerar($pasta) {
    # Devolve (quantidade, milissegundos). So EntryID + LastModificationTime:
    # e o par minimo que a varredura por geracao precisa.
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $n = 0
    $t = $null
    try {
        $t = $pasta.GetTable()
        $cols = $t.Columns
        try {
            $cols.RemoveAll()
            [void]$cols.Add("EntryID")
            [void]$cols.Add($P_LASTMOD)
        } finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($cols) }

        $chaves = New-Object 'System.Collections.Generic.HashSet[string]'
        while (-not $t.EndOfTable) {
            $a = $t.GetArray(200)
            for ($r = $a.GetLowerBound(0); $r -le $a.GetUpperBound(0); $r++) {
                [void]$chaves.Add("$($a.GetValue($r,0))")
                $n++
            }
        }
        $sw.Stop()
        return @{ N = $n; Ms = $sw.Elapsed.TotalMilliseconds; Distintas = $chaves.Count }
    } finally {
        if ($t) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($t) }
    }
}

$total = @{ N = 0; Ms = 0.0 }
Write-Host ("{0,-34} | {1,6} | {2,8} | {3,8} | {4,8} | chaves" -f `
    "pasta", "itens", "1a (ms)", "2a (ms)", "3a (ms)")
Write-Host ("-" * 92)

$pastas = @(
    @{ Id = 6;  Nome = "Caixa de Entrada" },
    @{ Id = 5;  Nome = "Itens Enviados" },
    @{ Id = 16; Nome = "Rascunhos" },
    @{ Id = 3;  Nome = "Itens Excluidos" },
    @{ Id = 23; Nome = "Lixo Eletronico" }
)

foreach ($p in $pastas) {
    $pasta = $ns.GetDefaultFolder($p.Id)
    try {
        $r = @()
        for ($i = 0; $i -lt 3; $i++) { $r += Enumerar $pasta }
        $ok = if ($r[0].Distintas -eq $r[0].N) { "todas unicas" } else { "REPETIDA!" }
        Write-Host ("{0,-34} | {1,6} | {2,8:N1} | {3,8:N1} | {4,8:N1} | {5}" -f `
            $p.Nome, $r[0].N, $r[0].Ms, $r[1].Ms, $r[2].Ms, $ok)
        $total.N += $r[0].N
        $total.Ms += ($r[1].Ms + $r[2].Ms) / 2   # descarta a 1a, que paga o aquecimento
    } finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pasta) }
}

Write-Host ("-" * 92)
Write-Host ("{0,-34} | {1,6} | {2,8:N1} ms somando as 5 pastas" -f "TOTAL", $total.N, $total.Ms)
Write-Host ""

# --- a caixa INTEIRA, que e o que a opcao 2 da Q4 exigiria ---
$sw = [Diagnostics.Stopwatch]::StartNew()
$n = 0
$pastas = 0

function Varrer($pasta, [int]$prof) {
    if ($prof -gt 12) { return }
    $script:pastas++
    $t = $null
    try {
        $t = $pasta.GetTable()
        $cols = $t.Columns
        try {
            $cols.RemoveAll()
            [void]$cols.Add("EntryID")
            [void]$cols.Add($P_LASTMOD)
        } finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($cols) }
        while (-not $t.EndOfTable) {
            $a = $t.GetArray(200)
            $script:n += ($a.GetUpperBound(0) - $a.GetLowerBound(0) + 1)
        }
    } catch {
    } finally {
        if ($t) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($t) }
    }
    $filhas = $null
    try { $filhas = $pasta.Folders } catch { return }
    try {
        for ($k = 1; $k -le $filhas.Count; $k++) {
            $f = $filhas.Item($k)
            try { Varrer $f ($prof + 1) } finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($f) }
        }
    } finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($filhas) }
}

for ($rodada = 1; $rodada -le 3; $rodada++) {
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $n = 0; $pastas = 0
    $stores = $ns.Stores
    for ($s = 1; $s -le $stores.Count; $s++) {
        $store = $stores.Item($s)
        $raiz = $null
        try { $raiz = $store.GetRootFolder(); Varrer $raiz 0 }
        catch { }
        finally {
            if ($raiz) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($raiz) }
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($store)
        }
    }
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($stores)
    $sw.Stop()
    Write-Host ("CAIXA INTEIRA, rodada {0}: {1} pastas, {2} itens, {3:N0} ms" -f `
        $rodada, $pastas, $n, $sw.Elapsed.TotalMilliseconds)
}

Write-Host ""
Write-Host "LEITURA: se a caixa inteira sai em poucos segundos, a opcao 2 da"
Write-Host "Q4 — exigir DUAS observacoes completas e compativeis antes de"
Write-Host "confirmar ausencia — e viavel. E ela e exatamente o que a Q2 pediu"
Write-Host "para confirmar NAO COEXISTENCIA antes de correlacionar."
