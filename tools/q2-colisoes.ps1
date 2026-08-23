# Q2, parte 2: as COLISOES sao o mesmo item ou itens diferentes?
#
# SOMENTE LEITURA.
#
# PRIVACIDADE: assunto sai truncado em 30 caracteres e endereco so como
# dominio. E o minimo para eu decidir se dois itens sao a mesma mensagem;
# menos que isso nao da para julgar, e mais nao e preciso.
#
# Esta e a pergunta que decide a Q2. Um grupo com dois itens significa que
# a evidencia NAO os distingue. Se eles forem mensagens DIFERENTES, usar
# aquela evidencia sozinha funde as duas — o R2-G, que e o pior risco da
# fase.

$ol = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
$ns = $ol.GetNamespace("MAPI")

$PR_SEARCH_KEY = "http://schemas.microsoft.com/mapi/proptag/0x300B0102"
$PR_MSG_ID     = "http://schemas.microsoft.com/mapi/proptag/0x1035001E"

function Levantar($id, [string]$rotulo) {
    $pasta = $ns.GetDefaultFolder($id)
    $t = $pasta.GetTable()
    $t.Columns.RemoveAll()
    foreach ($c in @("EntryID","Subject","SenderName","ReceivedTime","Size",
                     "MessageClass",$PR_SEARCH_KEY,$PR_MSG_ID)) { [void]$t.Columns.Add($c) }

    $itens = @()
    while (-not $t.EndOfTable) {
        $a = $t.GetArray(200)
        for ($r = $a.GetLowerBound(0); $r -le $a.GetUpperBound(0); $r++) {
            $sk = $a.GetValue($r, 6)
            $skTexto = if ($null -eq $sk) { "" } else {
                if ($sk -is [array]) { ($sk | ForEach-Object { $_.ToString("x2") }) -join "" } else { "$sk" }
            }
            $itens += [pscustomobject]@{
                Pasta    = $rotulo
                Id       = "$($a.GetValue($r,0))"
                Assunto  = "$($a.GetValue($r,1))"
                Remet    = "$($a.GetValue($r,2))"
                Quando   = $a.GetValue($r,3)
                Tamanho  = $a.GetValue($r,4)
                Classe   = "$($a.GetValue($r,5))"
                Sk       = $skTexto
                Mid      = "$($a.GetValue($r,7))".Trim()
            }
        }
    }
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($t)
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pasta)
    return ,$itens
}

$todos = @()
foreach ($p in @(@(6,"Entrada"), @(5,"Enviados"), @(16,"Rascunhos"), @(3,"Excluidos"))) {
    $todos += Levantar $p[0] $p[1]
}
Write-Host "levantados: $($todos.Count) itens"
Write-Host ""

function Curto([string]$s, [int]$n = 30) {
    if ([string]::IsNullOrEmpty($s)) { return "(vazio)" }
    if ($s.Length -le $n) { return $s }
    return $s.Substring(0, $n) + "..."
}

function Mostrar($grupos, [string]$titulo) {
    Write-Host "=============================================================="
    Write-Host $titulo
    Write-Host "=============================================================="
    if (-not $grupos -or @($grupos).Count -eq 0) { Write-Host "  (nenhum)"; return }

    $n = 0
    foreach ($g in @($grupos)) {
        $n++
        Write-Host ""
        Write-Host ("GRUPO {0} — {1} itens" -f $n, $g.Count)
        foreach ($it in $g.Group) {
            Write-Host ("  [{0,-9}] {1,-33} | {2} | {3} bytes | {4}" -f `
                $it.Pasta, (Curto $it.Assunto), `
                $(if ($null -eq $it.Quando) { "sem data          " } else { ([datetime]$it.Quando).ToString("dd/MM/yyyy HH:mm") }), `
                $it.Tamanho, (Curto $it.Remet 18))
        }
        # Mesmo item aparecendo duas vezes, ou dois itens diferentes?
        $ids = @($g.Group | Select-Object -ExpandProperty Id -Unique)
        $tam = @($g.Group | Select-Object -ExpandProperty Tamanho -Unique)
        Write-Host ("  -> EntryIDs distintos: {0}   tamanhos distintos: {1}" -f $ids.Count, $tam.Count)
    }
}

$colMid = @($todos | Where-Object { $_.Mid -ne "" } | Group-Object Mid | Where-Object { $_.Count -gt 1 })
$colSk  = @($todos | Where-Object { $_.Sk  -ne "" } | Group-Object Sk  | Where-Object { $_.Count -gt 1 })

Mostrar $colMid "COLISOES DE Message-ID  ($($colMid.Count) grupos)"
Write-Host ""
Mostrar $colSk  "COLISOES DE SearchKey   ($($colSk.Count) grupos)"

Write-Host ""
Write-Host "=============================================================="
Write-Host "SearchKey distingue itens que o Message-ID confunde?"
Write-Host "=============================================================="
foreach ($g in $colMid) {
    $sks = @($g.Group | Select-Object -ExpandProperty Sk -Unique)
    Write-Host ("  grupo de {0} itens com o mesmo Message-ID -> {1} SearchKey distintas" -f `
        $g.Count, $sks.Count)
}
