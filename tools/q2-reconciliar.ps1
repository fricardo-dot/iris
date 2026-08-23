# Q2: achar as copias que o experimento criou, e so elas.
#
# Antes do experimento: Itens Excluidos 141, Lixo Eletronico 172.
# Depois da limpeza:    Itens Excluidos 142, Lixo Eletronico 173.
#
# Sobrou uma copia em cada. O log do q2-move.ps1 nao explica as duas, e
# contagem que nao fecha nao pode ser declarada limpa.
#
# Este script SO LISTA. Nao apaga nada. Identifica candidatos a copia por
# assunto duplicado dentro da MESMA pasta, e mostra RecordKey e
# PR_LAST_MODIFICATION_TIME de cada um para eu decidir qual e qual.
#
# A copia e distinguivel: o Copy() de hoje tem PR_CREATION_TIME de hoje.

$ErrorActionPreference = "Stop"
$PT = "http://schemas.microsoft.com/mapi/proptag/"

$ol = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
$ns = $ol.GetNamespace("MAPI")

$hoje = (Get-Date).Date

foreach ($id in @(3, 23)) {
    $pasta = $ns.GetDefaultFolder($id)
    Write-Host ("=" * 78)
    Write-Host ("{0} — {1} itens" -f $pasta.Name, $pasta.Items.Count)
    Write-Host ("=" * 78)

    $t = $pasta.GetTable()
    $cols = $t.Columns
    try {
        $cols.RemoveAll()
        [void]$cols.Add("EntryID")
        [void]$cols.Add("Subject")
        [void]$cols.Add(($PT + "0x30070040"))   # PR_CREATION_TIME
        [void]$cols.Add(($PT + "0x30080040"))   # PR_LAST_MODIFICATION_TIME
    } finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($cols) }

    $linhas = @()
    while (-not $t.EndOfTable) {
        $a = $t.GetArray(200)
        for ($r = $a.GetLowerBound(0); $r -le $a.GetUpperBound(0); $r++) {
            $linhas += [pscustomobject]@{
                Id      = "$($a.GetValue($r,0))"
                Assunto = "$($a.GetValue($r,1))"
                Criado  = $a.GetValue($r,2)
                Alterado= $a.GetValue($r,3)
            }
        }
    }
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($t)

    # criados HOJE: o Copy() de hoje aparece aqui; mensagem antiga nao.
    $novos = @($linhas | Where-Object {
        $null -ne $_.Criado -and ([datetime]$_.Criado).Date -eq $hoje
    })
    Write-Host ("itens com PR_CREATION_TIME de HOJE: {0}" -f $novos.Count)
    foreach ($n in $novos) {
        $s = $n.Assunto
        if ($s.Length -gt 46) { $s = $s.Substring(0,46) + "..." }
        Write-Host ("   criado {0:HH:mm:ss}  {1}" -f [datetime]$n.Criado, $s)
    }

    # e assuntos repetidos dentro da mesma pasta
    $rep = @($linhas | Group-Object Assunto | Where-Object { $_.Count -gt 1 })
    Write-Host ""
    Write-Host ("assuntos repetidos dentro desta pasta: {0} grupo(s)" -f $rep.Count)
    foreach ($g in $rep) {
        $s = $g.Name
        if ($s.Length -gt 42) { $s = $s.Substring(0,42) + "..." }
        Write-Host ("   {0}x  {1}" -f $g.Count, $s)
        foreach ($it in $g.Group) {
            Write-Host ("        criado {0}  alterado {1}" -f $it.Criado, $it.Alterado)
        }
    }
    Write-Host ""
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pasta)
}
