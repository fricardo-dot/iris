# Q2: onde estao as manifestacoes dos dois itens do experimento, ENTRE AS
# PASTAS QUE ESTE ROTEIRO CONSEGUIU PERCORRER.
#
# O cabecalho dizia "TODAS as manifestacoes", e a travessia corta na
# profundidade 12 e podia perder um ramo em silencio -- entao o "TOTAL de
# copias" do fim seria zero sobre o que nao foi lido. Agora os dois casos
# sao contados e aparecem na conclusao.
#
# SOMENTE LEITURA.
#
# PR_CREATION_TIME nao serve para achar a copia: o Copy() preserva o
# valor do original. Mas a SearchKey e copiada junto — foi exatamente o que
# o experimento mediu — entao ela localiza original e copia de uma vez.
#
# Esperado se a limpeza tiver funcionado:
#   A (1345298c...) : 2 itens  — o enviado, em Itens Enviados,
#                                e o recebido, em Itens Excluidos
#   B (6e27c63c...) : 1 item   — em Lixo Eletronico
#
# Qualquer item a mais e uma copia que o experimento deixou.

$ErrorActionPreference = "Stop"
$PT = "http://schemas.microsoft.com/mapi/proptag/"

$alvos = @{
    "1345298cfcc2bf4d83e5fbbfeff17e5e" = "A: [IRIS-SPIKE-C] (esperado 2)"
    "6e27c63c795e034da35ae7e823bed6e1" = "B: mais antiga do Lixo (esperado 1)"
}

$ol = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
$ns = $ol.GetNamespace("MAPI")

$achados = @{}
foreach ($k in $alvos.Keys) { $achados[$k] = @() }

$script:cortados = 0
$script:ramosCegos = 0

function Varrer($pasta, [string]$caminho, [int]$prof) {
    if ($prof -gt 12) { $script:cortados++; return }
    $t = $null
    try {
        $t = $pasta.GetTable()
        $cols = $t.Columns
        try {
            $cols.RemoveAll()
            [void]$cols.Add("EntryID")
            [void]$cols.Add("Subject")
            [void]$cols.Add(($PT + "0x300B0102"))
            [void]$cols.Add(($PT + "0x0FF90102"))
        } finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($cols) }

        while (-not $t.EndOfTable) {
            $a = $t.GetArray(200)
            for ($r = $a.GetLowerBound(0); $r -le $a.GetUpperBound(0); $r++) {
                $sk = $a.GetValue($r, 2)
                if ($null -eq $sk -or -not ($sk -is [byte[]])) { continue }
                $hex = (($sk | ForEach-Object { $_.ToString("x2") }) -join "")
                if (-not $script:achados.ContainsKey($hex)) { continue }
                $rk = $a.GetValue($r, 3)
                $rkHex = if ($rk -is [byte[]]) { (($rk | ForEach-Object { $_.ToString("x2") }) -join "") } else { "?" }
                $script:achados[$hex] += [pscustomobject]@{
                    Pasta = $caminho
                    Id = "$($a.GetValue($r,0))"
                    Assunto = "$($a.GetValue($r,1))"
                    Rk = $rkHex
                }
            }
        }
    } catch {
        Write-Host "  FALHOU $caminho : $($_.Exception.Message)"
    } finally {
        if ($t) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($t) }
    }

    $filhas = $null
    try { $filhas = $pasta.Folders } catch { $script:ramosCegos++; return }
    try {
        for ($k = 1; $k -le $filhas.Count; $k++) {
            $f = $filhas.Item($k)
            try { Varrer $f "$caminho/$($f.Name)" ($prof + 1) }
            finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($f) }
        }
    } finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($filhas) }
}

$stores = $ns.Stores
for ($s = 1; $s -le $stores.Count; $s++) {
    $store = $stores.Item($s)
    $raiz = $null
    try { $raiz = $store.GetRootFolder(); Varrer $raiz $store.DisplayName 0 }
    catch { Write-Host "store inacessivel: $($_.Exception.Message)" }
    finally {
        if ($raiz) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($raiz) }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($store)
    }
}
[void][Runtime.InteropServices.Marshal]::ReleaseComObject($stores)

$sobra = 0
foreach ($k in $alvos.Keys) {
    Write-Host ("=" * 74)
    Write-Host $alvos[$k]
    Write-Host ("=" * 74)
    $lista = @($achados[$k])
    Write-Host ("encontrados: {0}" -f $lista.Count)
    foreach ($it in $lista) {
        Write-Host ("   {0}" -f $it.Pasta.Split("/", 2)[-1])
        Write-Host ("      RecordKey ...{0}" -f $it.Rk.Substring([Math]::Max(0, $it.Rk.Length - 16)))
    }
    $esperado = if ($k.StartsWith("1345")) { 2 } else { 1 }
    if ($lista.Count -gt $esperado) {
        $sobra += ($lista.Count - $esperado)
        Write-Host ("   >>> {0} A MAIS que o esperado: copia(s) do experimento." -f ($lista.Count - $esperado))
    }
    Write-Host ""
}

Write-Host ("TOTAL de copias deixadas pelo experimento, ENTRE O QUE FOI LIDO: {0}" -f $sobra)
if ($script:cortados -gt 0 -or $script:ramosCegos -gt 0) {
    Write-Host ("  RESSALVA: {0} ramo(s) cortados na profundidade 12 e {1} que nao" -f $script:cortados, $script:ramosCegos) -ForegroundColor DarkYellow
    Write-Host "  consegui abrir. Zero aqui nao prova que nao sobrou copia."
}
