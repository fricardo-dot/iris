# Q4, o que faltava: existe SINAL OPERACIONAL de varredura truncada?
#
# ESCREVE. So em pasta criada aqui e itens com marcador GUID desta execucao.
#
# ------------------------------------------------------------------
# A PERGUNTA
#
# A secao 16.1 mediu que um unico Move durante a varredura trunca a Table
# em 16 itens, TRES VEZES SEGUIDAS, e ela reporta EndOfTable normalmente.
# Isso deixa o Iris devolvendo 60% de uma pasta achando que leu tudo.
#
# Mas nos experimentos quem provocou a mutacao fui EU — entao eu sabia. Em
# producao ninguem avisa. A pergunta que fecha a Q4 e:
#
#   Que observacao BARATA, feita pelo proprio Iris, distingue uma
#   varredura completa de uma truncada?
#
# Candidatos, do mais barato ao mais caro:
#   S-a) Items.Count antes e depois, comparado com o que foi lido
#   S-b) PR_CONTENT_COUNT da pasta (0x36020003)
#   S-c) PR_CHANGE_KEY da pasta — muda quando o conteudo muda?
#
# O criterio de um sinal util: tem de acusar a truncagem E tem de ficar
# quieto quando nao houve nenhuma. Sinal que acusa sempre nao serve; sinal
# que nunca acusa, menos ainda. Por isso o script roda os DOIS casos.

$ErrorActionPreference = "Stop"
$PT = "http://schemas.microsoft.com/mapi/proptag/"
$P_CONTENT = $PT + "0x36020003"   # PR_CONTENT_COUNT
$P_CHANGE  = $PT + "0x65E20102"   # PR_CHANGE_KEY
$MARCA = "IRISQ4S-" + ([guid]::NewGuid().ToString("N").Substring(0, 8))

$ol = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
$ns = $ol.GetNamespace("MAPI")
function Solta($o) { if ($null -ne $o) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } }

function Sinais($pasta) {
    $r = [ordered]@{}
    $itens = $pasta.Items
    try { $r.Count = $itens.Count } finally { Solta $itens }
    $pa = $pasta.PropertyAccessor
    try {
        foreach ($p in @(@{ N = "ContentCount"; T = $P_CONTENT }, @{ N = "ChangeKey"; T = $P_CHANGE })) {
            try {
                $v = $pa.GetProperty($p.T)
                $r[$p.N] = if ($v -is [byte[]]) { (($v | ForEach-Object { $_.ToString("X2") }) -join "") } else { "$v" }
            } catch { $r[$p.N] = "(ERRO)" }
        }
    } finally { Solta $pa }
    return $r
}

function Varrer($pasta, [int]$lote, [scriptblock]$gancho) {
    $total = 0; $n = 0
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
    return $total
}

$raiz = $ns.GetDefaultFolder(6).Parent
$origem = $null; $destino = $null

try {
    $origem = $raiz.Folders.Add("Iris Q4S origem")
    $destino = $raiz.Folders.Add("Iris Q4S destino")
    $rasc = $ns.GetDefaultFolder(16)
    $li = $rasc.Items
    for ($i = 1; $i -le 40; $i++) {
        $m = $li.Add("IPM.Note")
        $m.Subject = ("{0} {1:d3}" -f $MARCA, $i)
        $m.Save()
        $mv = $m.Move($origem); Solta $mv; Solta $m
    }
    Solta $li; Solta $rasc

    Write-Host ("{0,-24} | {1,6} | {2,6} | {3,7} | {4,7} | {5}" -f `
        "cenario", "lidos", "antes", "depois", "content", "ChangeKey mudou")
    Write-Host ("-" * 82)

    $resultados = @()

    foreach ($cenario in @("SEM mutacao", "COM um Move no lote 1")) {
        # devolve tudo para a origem
        foreach ($f in @($destino.Items)) { $mv = $f.Move($origem); Solta $mv; Solta $f }

        $antes = Sinais $origem

        $script:moveu = $false
        $script:alvo = $null
        if ($cenario -like "COM*") {
            $itens = $origem.Items
            for ($i = 1; $i -le $itens.Count; $i++) {
                $it = $itens.Item($i)
                if ("$($it.Subject)" -eq ("{0} {1:d3}" -f $MARCA, 40)) { $script:alvo = $it; break }
                Solta $it
            }
            Solta $itens
        }

        $lidos = Varrer $origem 10 {
            param($lote)
            if ($lote -eq 1 -and $null -ne $script:alvo -and -not $script:moveu) {
                $script:moveu = $true
                $mv = $script:alvo.Move($script:destino); Solta $mv
            }
        }
        if ($script:alvo) { Solta $script:alvo }

        $depois = Sinais $origem
        $mudouCk = ($antes.ChangeKey -ne $depois.ChangeKey)

        Write-Host ("{0,-24} | {1,6} | {2,6} | {3,7} | {4,7} | {5}" -f `
            $cenario, $lidos, $antes.Count, $depois.Count, $depois.ContentCount,
            $(if ($mudouCk) { "SIM" } else { "nao" }))

        $resultados += @{
            Cenario = $cenario; Lidos = $lidos
            Antes = $antes.Count; Depois = $depois.Count
            MudouCk = $mudouCk
            ContentAntes = $antes.ContentCount; ContentDepois = $depois.ContentCount
        }
    }

    Write-Host ""
    Write-Host ("=" * 82)
    Write-Host "OS TRES CANDIDATOS, avaliados nos DOIS cenarios"
    Write-Host ("=" * 82)

    $sem = $resultados[0]; $com = $resultados[1]

    function Avaliar([string]$nome, [bool]$acusaSem, [bool]$acusaCom) {
        $veredito = if ($acusaCom -and -not $acusaSem) { "SERVE" }
                    elseif ($acusaCom -and $acusaSem)  { "acusa SEMPRE - inutil" }
                    elseif (-not $acusaCom)            { "NAO acusa a truncagem" }
                    else { "?" }
        Write-Host ("  {0,-34} sem mutacao: {1,-5} com mutacao: {2,-5} => {3}" -f `
            $nome, $(if ($acusaSem) { "acusa" } else { "quieto" }),
            $(if ($acusaCom) { "acusa" } else { "quieto" }), $veredito)
    }

    Avaliar "S-a  lidos <> Count ANTES" `
        ($sem.Lidos -ne $sem.Antes) ($com.Lidos -ne $com.Antes)
    Avaliar "S-a' lidos <> Count DEPOIS" `
        ($sem.Lidos -ne $sem.Depois) ($com.Lidos -ne $com.Depois)
    Avaliar "S-b  Count antes <> Count depois" `
        ($sem.Antes -ne $sem.Depois) ($com.Antes -ne $com.Depois)
    Avaliar "S-c  ChangeKey da pasta mudou" $sem.MudouCk $com.MudouCk

    Write-Host ""
    Write-Host ("  PR_CONTENT_COUNT: sem={0}/{1}  com={2}/{3}" -f `
        $sem.ContentAntes, $sem.ContentDepois, $com.ContentAntes, $com.ContentDepois)

} catch {
    Write-Host ""
    Write-Host "!!! FALHA !!!"; Write-Host $_.Exception.Message; Write-Host $_.ScriptStackTrace
} finally {
    Write-Host ""
    Write-Host "LIMPEZA"
    $rasc = $ns.GetDefaultFolder(16)
    $li = $rasc.Items; $n = 0
    for ($i = $li.Count; $i -ge 1; $i--) {
        $it = $li.Item($i)
        try { if ("$($it.Subject)".StartsWith($MARCA, [StringComparison]::Ordinal)) { $it.Delete(); $n++ } }
        finally { Solta $it }
    }
    Solta $li; Solta $rasc
    Write-Host ("  Rascunhos: {0} com o marcador removidos" -f $n)
    foreach ($f in @($raiz.Folders)) {
        try { if ($f.Name -like "Iris Q4S *") {
            Write-Host ("  removendo {0} ({1} itens)" -f $f.Name, $f.Items.Count); $f.Delete() } }
        catch { } finally { Solta $f }
    }
    Solta $origem; Solta $destino; Solta $raiz
}
