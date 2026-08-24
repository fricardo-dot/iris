# Q3 com o Restrict de verdade, E com o checkpoint formatado direito.
# Mais: confirmar se um Move durante a varredura TRUNCA a Table.
#
# ESCREVE. So em pasta criada aqui e itens com marcador GUID desta
# execucao. Nenhuma mensagem do usuario e criada, movida ou apagada.
#
# RESSALVA HONESTA: a limpeza PERCORRE os Rascunhos lendo Subject, para
# achar os proprios artefatos que possam ter ficado entre o Save e o
# Move. O marcador GUID torna a exclusao segura contra colisao; nao
# torna verdadeira a frase "nao le nada do usuario".
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
    # PARTE 2 - o Restrict, com CONTROLE POSITIVO e DOIS formatos
    #
    # A execucao anterior falhou o controle positivo: o filtro
    # "[LastModificationTime] > '08/23/2026 21:07:15'" devolveu ZERO
    # inclusive para um item criado DEPOIS do checkpoint. Ou seja: o filtro
    # nao funcionava, e o "nao achou" nao provava nada sobre Move-in.
    #
    # Hipotese: a sintaxe [Campo] do Restrict nao aceita SEGUNDOS. Se for
    # isso, o .ToString("g") que eu tinha chamado de bug era o formato
    # CERTO, e eu "consertei" o teste quebrando o filtro.
    #
    # Por isso agora:
    #   - dois formatos, um com segundos e um sem;
    #   - e tambem o caminho @SQL com o proptag, que a Q1 mediu ser em UTC;
    #   - separacao de MAIS DE UM MINUTO entre o carimbo do alvo e o
    #     checkpoint, para o truncamento ao minuto nao criar ambiguidade;
    #   - controle positivo obrigatorio em todos.
    # =============================================================
    Write-Host ""
    Write-Host ("=" * 70)
    Write-Host "PARTE 2 - Restrict com controle positivo, tres filtros"
    Write-Host ("=" * 70)

    $itens = $origem.Items
    $alvo = $null
    for ($i = 1; $i -le $itens.Count; $i++) {
        $it = $itens.Item($i)
        if ("$($it.Subject)" -eq ("{0} {1:d3}" -f $MARCA, 5)) { $alvo = $it; break }
        Solta $it
    }
    Solta $itens

    $pa = $alvo.PropertyAccessor
    $lmAlvo = [datetime]$pa.GetProperty($P_LASTMOD)   # UTC
    Solta $pa
    Write-Host ("  LMT do alvo (UTC)  : {0:yyyy-MM-dd HH:mm:ss}" -f $lmAlvo)
    Write-Host ("  LMT do alvo (LOCAL): {0:yyyy-MM-dd HH:mm:ss}" -f $lmAlvo.ToLocalTime())

    Write-Host "  aguardando 70 s para o checkpoint ficar a MAIS DE UM MINUTO..."
    Start-Sleep -Seconds 70
    $checkpoint = Get-Date
    Write-Host ("  checkpoint (LOCAL) : {0:yyyy-MM-dd HH:mm:ss}" -f $checkpoint)
    Start-Sleep -Seconds 5

    # controle positivo: criado DEPOIS do checkpoint
    $rasc2 = $ns.GetDefaultFolder(16)
    $li2 = $rasc2.Items
    $ctrl = $li2.Add("IPM.Note")
    $ctrl.Subject = "$MARCA CONTROLE"
    $ctrl.Save()
    $ctrlMov = $ctrl.Move($destino)
    $chaveCtrl = $ctrlMov.EntryID
    $pa = $ctrlMov.PropertyAccessor
    $lmCtrl = [datetime]$pa.GetProperty($P_LASTMOD)
    Solta $pa
    Solta $ctrlMov; Solta $ctrl; Solta $li2; Solta $rasc2
    Write-Host ("  LMT do controle (UTC): {0:yyyy-MM-dd HH:mm:ss}" -f $lmCtrl)

    # o alvo entra por MOVE, com carimbo velho
    $movido = $alvo.Move($destino)
    $chaveAlvo = $movido.EntryID
    $pa = $movido.PropertyAccessor
    $lmDepois = [datetime]$pa.GetProperty($P_LASTMOD)
    Solta $pa; Solta $movido; Solta $alvo
    Write-Host ("  LMT do alvo APOS o Move (UTC): {0:yyyy-MM-dd HH:mm:ss}  mudou? {1}" -f `
        $lmDepois, $(if ($lmDepois -ne $lmAlvo) { "SIM" } else { "NAO" }))
    Write-Host ""

    $filtros = @(
        @{ Nome = "[Campo] 12h SEM segundos"
           F = "[LastModificationTime] > '" + $checkpoint.ToString("MM/dd/yyyy hh:mm tt") + "'" },
        @{ Nome = "[Campo] 24h COM segundos"
           F = "[LastModificationTime] > '" + $checkpoint.ToString("MM/dd/yyyy HH:mm:ss") + "'" },
        @{ Nome = "@SQL proptag em UTC"
           F = '@SQL="' + $P_LASTMOD + '" > ' + "'" + $checkpoint.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss") + "'" }
    )

    foreach ($f in $filtros) {
        $itensD = $destino.Items
        try {
            $ok = $true; $achouCtrl = $false; $achouAlvo = $false; $n = 0
            try {
                $rs = $itensD.Restrict($f.F)
                $n = $rs.Count
                for ($i = 1; $i -le $n; $i++) {
                    $r = $rs.Item($i)
                    try {
                        if ($r.EntryID -eq $chaveCtrl) { $achouCtrl = $true }
                        if ($r.EntryID -eq $chaveAlvo) { $achouAlvo = $true }
                    } finally { Solta $r }
                }
                Solta $rs
            } catch { $ok = $false; $erro = $_.Exception.Message }

            Write-Host ("  --- {0} ---" -f $f.Nome)
            Write-Host ("     {0}" -f $f.F)
            if (-not $ok) {
                Write-Host ("     Restrict LANCOU: {0}" -f $erro)
            } else {
                Write-Host ("     devolveu {0}; controle={1}  alvo-do-Move={2}" -f `
                    $n, $(if ($achouCtrl) { "ACHADO" } else { "nao" }), `
                    $(if ($achouAlvo) { "achado" } else { "nao" }))
                if (-not $achouCtrl) {
                    Write-Host "     => FILTRO INVALIDO (nem o controle apareceu). Nada a concluir."
                } elseif (-not $achouAlvo) {
                    Write-Host "     => PROVADO: filtro funciona, e o Move-in NAO e descoberto."
                } else {
                    Write-Host "     => o Move-in FOI descoberto por este filtro."
                }
            }
        } finally { Solta $itensD }
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
