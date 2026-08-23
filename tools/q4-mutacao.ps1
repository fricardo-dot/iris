# Q4 (consistencia sob mutacao), Q5 (o que a ausencia prova) e a lacuna
# Move-in da Q3.
#
# ESCREVE. So em pastas que este script cria e em itens que ele cria.
# Nenhuma mensagem do usuario e lida, movida ou apagada.
#
# ------------------------------------------------------------------
# FASES E ORACULOS SEPARADOS, e isso nao e formalidade
#
# Q4 pergunta: "esta geracao pode ser declarada VALIDA?"
# Q5 pergunta: "dadas observacoes validas, que conclusao e SEGURA?"
#
# Sao niveis diferentes. Uma mutacao concorrente pode justamente INVALIDAR
# a geracao — ai ela e evidencia para a Q4 e NAO PODE alimentar a Q5. Um
# experimento so, com um oraculo so, misturaria as duas e produziria
# conclusao semantica a partir de observacao que nem devia ter sido aceita.
#
# ------------------------------------------------------------------
# A LACUNA DA Q3, que so apareceu depois da Q2
#
# A Q3 testa item movido para FORA e exclusao. Falta o caso perigoso: item
# movido ou copiado PARA DENTRO. Se o Move preservar um
# LastModificationTime anterior ao checkpoint, "Restrict > X" NUNCA
# descobre a nova associacao — e janela de sobreposicao nao cobre timestamp
# arbitrariamente antigo.

$ErrorActionPreference = "Stop"
$PT = "http://schemas.microsoft.com/mapi/proptag/"
$P_LASTMOD = $PT + "0x30080040"
$P_ENTRYID = $PT + "0x66700102"

$ol = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
$ns = $ol.GetNamespace("MAPI")

function Solta($o) { if ($null -ne $o) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } }

function Hex($v) {
    if ($v -is [byte[]]) { return (($v | ForEach-Object { $_.ToString("X2") }) -join "") }
    return "$v"
}

# ------------------------------------------------------------------
# Enumeracao por Table, com um GANCHO entre lotes. O gancho e o que
# permite mutar a caixa COM O CURSOR ABERTO, que e o cenario real.
# ------------------------------------------------------------------
function Enumerar($pasta, [int]$lote = 25, [scriptblock]$entreLotes = $null) {
    $chaves = New-Object 'System.Collections.Generic.List[string]'
    $t = $null
    try {
        $t = $pasta.GetTable()
        $cols = $t.Columns
        try {
            $cols.RemoveAll()
            $c1 = $cols.Add($P_ENTRYID); Solta $c1
            $c2 = $cols.Add("Subject");  Solta $c2
        } finally { Solta $cols }

        $n = 0
        while (-not $t.EndOfTable) {
            $a = $t.GetArray($lote)
            for ($r = $a.GetLowerBound(0); $r -le $a.GetUpperBound(0); $r++) {
                $chaves.Add((Hex $a.GetValue($r, 0)))
            }
            $n++
            if ($entreLotes) { & $entreLotes $n }
        }
        # A virgula NAO e enfeite: sem ela o PowerShell DESENROLA a lista
        # no retorno, e lista vazia vira $null. Compare-Object recebendo
        # $null aborta com uma mensagem que nao fala de lista nenhuma.
        return ,$chaves
    } finally { Solta $t }
}

# ATENCAO ao idioma: Items.Add numa pasta qualquer CRIA o item, mas o
# Save() de uma mensagem nao enviada a deposita em RASCUNHOS, nao na pasta.
# Medido: m.Parent.Name devolve "Rascunhos" e a pasta alvo fica com 0.
# Por isso cria-se e MOVE-SE em seguida — e por isso a limpeza tambem varre
# Rascunhos, para o caso de falhar entre o Save e o Move.
function CriarItens($pasta, [int]$quantos, [string]$prefixo) {
    $rascunhos = $ns.GetDefaultFolder(16)
    try {
        $itens = $rascunhos.Items
        try {
            for ($i = 1; $i -le $quantos; $i++) {
                $m = $itens.Add("IPM.Note")
                try {
                    $m.Subject = "$prefixo $i"
                    $m.Body = "item de teste do Iris, pode apagar"
                    $m.Save()
                    $movido = $m.Move($pasta)
                    Solta $movido
                } finally { Solta $m }
            }
        } finally { Solta $itens }
    } finally { Solta $rascunhos }
}

function AcharPorAssunto($pasta, [string]$assunto) {
    $itens = $pasta.Items
    try {
        for ($i = 1; $i -le $itens.Count; $i++) {
            $it = $itens.Item($i)
            if ("$($it.Subject)" -eq $assunto) { return $it }
            Solta $it
        }
    } finally { Solta $itens }
    return $null
}

$raiz = $ns.GetDefaultFolder(6).Parent
$origem = $null; $destino = $null
$achados = @()

try {
    # =============================================================
    # FASE 0 - preparar
    # =============================================================
    Write-Host "FASE 0: preparando"
    $origem = $raiz.Folders.Add("Iris Q4 origem")
    $destino = $raiz.Folders.Add("Iris Q4 destino")
    CriarItens $origem 60 "Q4-item"
    Write-Host ("  origem: {0} itens | destino: {1}" -f $origem.Items.Count, $destino.Items.Count)

    # =============================================================
    # FASE 1 - Q4: a geracao e valida?
    # Oraculo: eu SEI quantos itens existem e o que eu mutei.
    # =============================================================
    Write-Host ""
    Write-Host ("=" * 74)
    Write-Host "FASE 1 - Q4: consistencia da ENUMERACAO sob mutacao"
    Write-Host ("=" * 74)

    # 1a. sem mutacao: duas passadas tem de dar o mesmo
    $p1 = Enumerar $origem
    $p2 = Enumerar $origem
    $iguais = (@(Compare-Object $p1 $p2).Count -eq 0)
    Write-Host ("  1a. sem mutacao: {0} e {1} chaves, manifestos {2}" -f `
        $p1.Count, $p2.Count, $(if ($iguais) { "IDENTICOS" } else { "DIFERENTES" }))
    $achados += @{ Caso = "1a estavel"; Ok = $iguais; Nota = "$($p1.Count) chaves" }

    # 1b. item CRIADO durante a enumeracao
    $criouNoMeio = $false
    $r = Enumerar $origem 25 {
        param($lote)
        if ($lote -eq 2 -and -not $script:criouNoMeio) {
            $script:criouNoMeio = $true
            CriarItens $origem 1 "Q4-NOVO-no-meio"
        }
    }
    $depois = Enumerar $origem
    Write-Host ("  1b. item criado no meio: a varredura viu {0}; a seguinte, {1}" -f `
        $r.Count, $depois.Count)
    Write-Host ("      (esperado antes: 60; depois: 61)")
    $achados += @{ Caso = "1b criado no meio"; Ok = ($depois.Count -eq 61)
                   Nota = "varredura corrente viu $($r.Count)" }

    # 1d. item MOVIDO de origem para destino COM O CURSOR ABERTO,
    #     e as duas pastas sendo varridas.
    $alvo = AcharPorAssunto $origem "Q4-item 40"
    if ($null -eq $alvo) { throw "nao achei o item alvo" }
    # PR_LONGTERM_ENTRYID_FROM_TABLE so existe como COLUNA DE TABLE: pelo
    # PropertyAccessor do item ela e "desconhecida ou nao encontrada". O
    # equivalente no item e o proprio .EntryID, que a secao 12.2 mediu ter
    # exatamente o mesmo valor.
    $chaveAntes = $alvo.EntryID

    $moveu = $false
    $naOrigem = Enumerar $origem 25 {
        param($lote)
        if ($lote -eq 2 -and -not $script:moveu) {
            $script:moveu = $true
            $m = $script:alvo.Move($script:destino)
            $script:chaveDepois = $m.EntryID
            Solta $m
        }
    }
    $noDestino = Enumerar $destino

    $viuNaOrigem = $naOrigem.Contains($chaveAntes)
    $viuNoDestino = $noDestino.Contains($chaveDepois)
    Write-Host ""
    Write-Host "  1d. item movido origem->destino COM O CURSOR ABERTO:"
    Write-Host ("      chave ANTES do move : ...{0}" -f $chaveAntes.Substring($chaveAntes.Length - 16))
    Write-Host ("      chave DEPOIS        : ...{0}" -f $chaveDepois.Substring($chaveDepois.Length - 16))
    Write-Host ("      mudou? {0}" -f $(if ($chaveAntes -ne $chaveDepois) { "SIM" } else { "nao" }))
    Write-Host ("      apareceu na varredura da ORIGEM (chave antiga)? {0}" -f $(if ($viuNaOrigem) { "SIM" } else { "nao" }))
    Write-Host ("      aparece no DESTINO (chave nova)?                {0}" -f $(if ($viuNoDestino) { "SIM" } else { "nao" }))
    if ($viuNaOrigem -and $viuNoDestino) {
        Write-Host "      => CONTADO DUAS VEZES no mesmo instante logico."
    } elseif (-not $viuNaOrigem -and -not $viuNoDestino) {
        Write-Host "      => PERDIDO: nao apareceu em nenhuma das duas."
    } else {
        Write-Host "      => apareceu numa so."
    }
    $achados += @{ Caso = "1d movido durante varredura"
                   Ok = $true
                   Nota = "origem=$viuNaOrigem destino=$viuNoDestino chave mudou=$($chaveAntes -ne $chaveDepois)" }

    # =============================================================
    # FASE 2 - Q3: Move-in atualiza LastModificationTime?
    # =============================================================
    Write-Host ""
    Write-Host ("=" * 74)
    Write-Host "FASE 2 - Q3: a lacuna do Move-in"
    Write-Host ("=" * 74)

    $sujeito = AcharPorAssunto $origem "Q4-item 10"
    $pa = $sujeito.PropertyAccessor
    $lmAntes = $pa.GetProperty($P_LASTMOD)
    Solta $pa
    Write-Host ("  LastModificationTime ANTES do move : {0}" -f $lmAntes)

    Start-Sleep -Seconds 2   # separar os instantes
    $checkpoint = (Get-Date).ToUniversalTime()

    $movido = $sujeito.Move($destino)
    $pa = $movido.PropertyAccessor
    $lmDepois = $pa.GetProperty($P_LASTMOD)
    Solta $pa
    Write-Host ("  checkpoint (UTC)                   : {0}" -f $checkpoint)
    Write-Host ("  LastModificationTime DEPOIS        : {0}" -f $lmDepois)

    $atualizou = ([datetime]$lmDepois -gt [datetime]$lmAntes)
    Write-Host ("  o Move ATUALIZOU o LastModificationTime? {0}" -f $(if ($atualizou) { "SIM" } else { "NAO" }))

    # A pergunta que decide o incremental: um Restrict "> checkpoint" no
    # DESTINO acha o item que acabou de chegar?
    $itensDestino = $destino.Items
    $q = [datetime]$checkpoint
    $filtro = "@SQL=" + '"' + "urn:schemas:httpmail:hasattachment" + '"' + " IS NOT NULL"
    Solta $itensDestino

    $tab = $destino.GetTable()
    try {
        $cols = $tab.Columns
        try {
            $cols.RemoveAll()
            $c = $cols.Add($P_LASTMOD); Solta $c
            $c = $cols.Add("Subject"); Solta $c
        } finally { Solta $cols }
        $achouPeloIncremental = $false
        while (-not $tab.EndOfTable) {
            $a = $tab.GetArray(50)
            for ($r = $a.GetLowerBound(0); $r -le $a.GetUpperBound(0); $r++) {
                $lm = $a.GetValue($r, 0)
                $assunto = "$($a.GetValue($r, 1))"
                if ($assunto -eq "Q4-item 10" -and $null -ne $lm) {
                    if ([datetime]$lm -gt $q) { $achouPeloIncremental = $true }
                }
            }
        }
    } finally { Solta $tab }

    Write-Host ("  um incremental 'LastModificationTime > checkpoint' acharia" )
    Write-Host ("  o item que ACABOU de chegar no destino? {0}" -f `
        $(if ($achouPeloIncremental) { "SIM" } else { "NAO — E ELE SUMIRIA DO CACHE" }))
    $achados += @{ Caso = "Q3 move-in"; Ok = $achouPeloIncremental
                   Nota = "LMT atualizou=$atualizou" }
    Solta $movido

    # =============================================================
    # FASE 3 - Q5: o que a ausencia prova (so sobre geracao valida)
    # =============================================================
    Write-Host ""
    Write-Host ("=" * 74)
    Write-Host "FASE 3 - Q5: o que a ausencia prova"
    Write-Host ("=" * 74)
    Write-Host "  (so sobre geracao ESTAVEL: nada esta sendo mutado agora)"

    $antesDaAusencia = Enumerar $origem
    $some = AcharPorAssunto $origem "Q4-item 30"
    $chaveSome = $some.EntryID
    $m2 = $some.Move($destino)
    Solta $m2
    $depoisDaAusencia = Enumerar $origem

    $sumiu = (-not $depoisDaAusencia.Contains($chaveSome))
    Write-Host ("  item movido para outra pasta. Sumiu da origem? {0}" -f $(if ($sumiu) { "SIM" } else { "nao" }))
    Write-Host ("  {0} -> {1} chaves" -f $antesDaAusencia.Count, $depoisDaAusencia.Count)
    Write-Host ""
    Write-Host "  A ORIGEM ve exatamente o mesmo que veria se o item tivesse"
    Write-Host "  sido APAGADO. E a chave dele mudou no move, entao procurar"
    Write-Host "  pela chave antiga no resto da caixa NAO o encontra."
    $achados += @{ Caso = "Q5 ausencia"; Ok = $sumiu
                   Nota = "ausencia na pasta nao distingue movido de apagado" }

} catch {
    Write-Host ""
    Write-Host "!!! FALHA !!!"
    Write-Host $_.Exception.Message
    Write-Host $_.ScriptStackTrace
} finally {
    Write-Host ""
    Write-Host ("=" * 74)
    Write-Host "LIMPEZA"
    Write-Host ("=" * 74)
    # Rascunhos primeiro: se o script morrer entre o Save e o Move, o item
    # fica LA, na pasta do usuario. Nao pode ficar.
    $rasc = $ns.GetDefaultFolder(16)
    try {
        $sobrou = 0
        $li = $rasc.Items
        try {
            for ($i = $li.Count; $i -ge 1; $i--) {
                $it = $li.Item($i)
                try {
                    if ("$($it.Subject)".StartsWith("Q4-", [StringComparison]::Ordinal)) {
                        $it.Delete()   # soft
                        $sobrou++
                    }
                } finally { Solta $it }
            }
        } finally { Solta $li }
        if ($sobrou -gt 0) {
            Write-Host ("  {0} item(ns) de teste removidos de RASCUNHOS" -f $sobrou)
        } else {
            Write-Host "  Rascunhos: nenhum item de teste sobrou"
        }
    } finally { Solta $rasc }

    foreach ($f in @($raiz.Folders)) {
        try {
            if ($f.Name -like "Iris Q4 *") {
                Write-Host ("  removendo {0} ({1} itens) -> vai para Itens Excluidos" -f $f.Name, $f.Items.Count)
                $f.Delete()
            }
        } catch { Write-Host ("  nao consegui: {0}" -f $_.Exception.Message) }
        finally { Solta $f }
    }
    Solta $origem; Solta $destino; Solta $raiz
}

Write-Host ""
Write-Host ("=" * 74)
Write-Host "RESUMO"
Write-Host ("=" * 74)
foreach ($a in $achados) {
    Write-Host ("  {0,-30} {1,-5} {2}" -f $a.Caso, $(if ($a.Ok) { "ok" } else { "!!" }), $a.Nota)
}
