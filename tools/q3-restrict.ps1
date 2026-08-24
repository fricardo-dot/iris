# Q3 com o Restrict de verdade, E com o checkpoint formatado direito.
# Mais: confirmar se um Move durante a varredura TRUNCA a Table.
#
# ESCREVE. So em pasta criada aqui e itens com marcador GUID desta execucao.
#
# ------------------------------------------------------------------
# O BUG DA EXECUCAO ANTERIOR, que quase virou conclusao publicada
#
# O checkpoint era formatado com .ToString("g"), que NAO tem segundos.
# PR_LAST_MODIFICATION_TIME volta em UTC pelo PropertyAccessor (medido:
# 23:54:46 UTC para 20:54:46 local, fuso -3), e o Restrict com a sintaxe
# [Campo] compara em hora LOCAL. Com o checkpoint truncado para 20:54:00,
# QUALQUER item modificado naquele minuto passava no filtro.
#
# O Restrict "achou" o item por truncamento, nao porque o Move tenha
# mexido no carimbo. Duas armadilhas de fuso e formato no mesmo teste, que
# e exatamente a familia de erro que a Q1 ja tinha custado caro.
#
# Agora: segundos no formato, e uma folga de 5 s de cada lado.

$ErrorActionPreference = "Stop"
$PT = "http://schemas.microsoft.com/mapi/proptag/"
$P_LASTMOD = $PT + "0x30080040"
$MARCA = "IRISQ3-" + ([guid]::NewGuid().ToString("N").Substring(0, 8))

$ol = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
$ns = $ol.GetNamespace("MAPI")
function Solta($o) { if ($null -ne $o) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } }

$raiz = $ns.GetDefaultFolder(6).Parent
$origem = $null; $destino = $null

try {
    $origem = $raiz.Folders.Add("Iris Q3R origem")
    $destino = $raiz.Folders.Add("Iris Q3R destino")

    $rasc = $ns.GetDefaultFolder(16)
    $li = $rasc.Items
    for ($i = 1; $i -le 40; $i++) {
        $m = $li.Add("IPM.Note")
        $m.Subject = ("{0} {1:d3}" -f $MARCA, $i)
        $m.Save()
        $mv = $m.Move($origem); Solta $mv; Solta $m
    }
    Solta $li; Solta $rasc
    Write-Host ("origem: {0} itens" -f $origem.Items.Count)

    # =============================================================
    # PARTE 1 - o Move TRUNCA a varredura? (o "24 de 40" da execucao
    # anterior era forte demais para aceitar sem repetir)
    # =============================================================
    Write-Host ""
    Write-Host ("=" * 70)
    Write-Host "PARTE 1 - um Move durante a varredura trunca a Table?"
    Write-Host ("=" * 70)

    function Varrer($pasta, [int]$lote, [scriptblock]$gancho) {
        $n = 0; $total = 0
        $t = $pasta.GetTable()
        try {
            $cols = $t.Columns
            $cols.RemoveAll()
            $c = $cols.Add("Subject"); Solta $c
            Solta $cols
            $t.Sort("Subject", $false)
            while (-not $t.EndOfTable) {
                $a = $t.GetArray($lote)
                $total += ($a.GetUpperBound(0) - $a.GetLowerBound(0) + 1)
                $n++
                if ($gancho) { & $gancho $n }
            }
        } finally { Solta $t }
        return @{ Total = $total; Lotes = $n }
    }

    # 1a. controle: sem mutacao
    $base = Varrer $origem 10 $null
    Write-Host ("  1a. sem mutacao        : {0} itens em {1} lotes" -f $base.Total, $base.Lotes)

    # 1b. move um item do FIM da ordem, no primeiro lote
    for ($rep = 1; $rep -le 3; $rep++) {
        $itens = $origem.Items
        $alvo = $null
        for ($i = 1; $i -le $itens.Count; $i++) {
            $it = $itens.Item($i)
            if ("$($it.Subject)" -eq ("{0} {1:d3}" -f $MARCA, (40 - $rep + 1))) { $alvo = $it; break }
            Solta $it
        }
        Solta $itens
        if ($null -eq $alvo) { Write-Host "     (alvo nao encontrado, pulando)"; continue }

        $script:moveuJa = $false
        $script:alvoRep = $alvo
        $r = Varrer $origem 10 {
            param($lote)
            if ($lote -eq 1 -and -not $script:moveuJa) {
                $script:moveuJa = $true
                $mv = $script:alvoRep.Move($script:destino); Solta $mv
            }
        }
        $antes = $base.Total - ($rep - 1)
        Write-Host ("  1b.{0} move no lote 1  : {1} itens (a pasta tinha {2}, esperado {3})" -f `
            $rep, $r.Total, $antes, ($antes - 1))
        Solta $alvo
    }

    # =============================================================
    # PARTE 2 - o Restrict, com checkpoint COM SEGUNDOS
    # =============================================================
    Write-Host ""
    Write-Host ("=" * 70)
    Write-Host "PARTE 2 - Items.Restrict com checkpoint formatado direito"
    Write-Host ("=" * 70)

    foreach ($modo in @("Move", "Copy")) {
        $itens = $origem.Items
        $alvo = $null
        $procurado = ("{0} {1:d3}" -f $MARCA, $(if ($modo -eq "Move") { 5 } else { 6 }))
        for ($i = 1; $i -le $itens.Count; $i++) {
            $it = $itens.Item($i)
            if ("$($it.Subject)" -eq $procurado) { $alvo = $it; break }
            Solta $it
        }
        Solta $itens
        if ($null -eq $alvo) { Write-Host "  ($procurado nao encontrado)"; continue }

        $pa = $alvo.PropertyAccessor
        $lmAntes = [datetime]$pa.GetProperty($P_LASTMOD)
        Solta $pa

        Start-Sleep -Seconds 5
        $checkpoint = Get-Date            # LOCAL: e o que [Campo] usa
        Start-Sleep -Seconds 5

        $novo = if ($modo -eq "Move") { $alvo.Move($destino) } else {
            $c = $alvo.Copy(); $t2 = $c.Move($destino); Solta $c; $t2
        }
        $pa = $novo.PropertyAccessor
        $lmDepois = [datetime]$pa.GetProperty($P_LASTMOD)
        Solta $pa
        $chave = $novo.EntryID
        Solta $novo

        # COM SEGUNDOS. E imprimindo os dois lados na MESMA base de tempo.
        $q = $checkpoint.ToString("MM/dd/yyyy HH:mm:ss")
        $filtro = "[LastModificationTime] > '$q'"

        Write-Host ""
        Write-Host ("  --- {0}-in ---" -f $modo)
        Write-Host ("     LMT antes  (UTC)   : {0:yyyy-MM-dd HH:mm:ss}" -f $lmAntes)
        Write-Host ("     LMT depois (UTC)   : {0:yyyy-MM-dd HH:mm:ss}" -f $lmDepois)
        Write-Host ("     LMT depois (LOCAL) : {0:yyyy-MM-dd HH:mm:ss}" -f $lmDepois.ToLocalTime())
        Write-Host ("     checkpoint (LOCAL) : {0:yyyy-MM-dd HH:mm:ss}" -f $checkpoint)
        Write-Host ("     o carimbo MUDOU?   : {0}" -f $(if ($lmDepois -ne $lmAntes) { "SIM" } else { "NAO" }))
        Write-Host ("     LMT local > checkpoint? {0}" -f $(if ($lmDepois.ToLocalTime() -gt $checkpoint) { "SIM" } else { "NAO" }))

        $itensD = $destino.Items
        try {
            $restritos = $itensD.Restrict($filtro)
            try {
                $achou = $false
                for ($i = 1; $i -le $restritos.Count; $i++) {
                    $r = $restritos.Item($i)
                    try { if ($r.EntryID -eq $chave) { $achou = $true } } finally { Solta $r }
                }
                Write-Host ("     filtro: {0}" -f $filtro)
                Write-Host ("     Restrict devolveu {0}; ACHOU o alvo? {1}" -f `
                    $restritos.Count, $(if ($achou) { "SIM" } else { "NAO" }))
                if (-not $achou -and $modo -eq "Move") {
                    Write-Host "     => o incremental por LMT NAO descobre o item que chegou."
                }
            } finally { Solta $restritos }
        } finally { Solta $itensD }
        Solta $alvo
    }

} catch {
    Write-Host ""
    Write-Host "!!! FALHA !!!"
    Write-Host $_.Exception.Message
    Write-Host $_.ScriptStackTrace
} finally {
    Write-Host ""
    Write-Host ("=" * 70)
    Write-Host "LIMPEZA"
    $rasc = $ns.GetDefaultFolder(16)
    $li = $rasc.Items
    $n = 0
    for ($i = $li.Count; $i -ge 1; $i--) {
        $it = $li.Item($i)
        try { if ("$($it.Subject)".StartsWith($MARCA, [StringComparison]::Ordinal)) { $it.Delete(); $n++ } }
        finally { Solta $it }
    }
    Solta $li; Solta $rasc
    Write-Host ("  Rascunhos: {0} com o marcador removidos" -f $n)
    foreach ($f in @($raiz.Folders)) {
        try { if ($f.Name -like "Iris Q3R *") {
            Write-Host ("  removendo {0} ({1} itens)" -f $f.Name, $f.Items.Count); $f.Delete() } }
        catch { } finally { Solta $f }
    }
    Solta $origem; Solta $destino; Solta $raiz
}
