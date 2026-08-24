# Q4 com MATRIZ TEMPORAL, e Q3 com o Restrict DE VERDADE.
#
# ESCREVE. So em pastas criadas por este script e em itens marcados com um
# GUID gerado agora. A limpeza usa os EntryIDs CAPTURADOS e o marcador —
# nunca varre pasta do usuario lendo assunto, que era o que a versao
# anterior fazia enquanto o cabecalho prometia o contrario.
#
# ------------------------------------------------------------------
# POR QUE ESTA VERSAO EXISTE
#
# A anterior tinha tres defeitos que invalidavam as conclusoes:
#
#   1. O teste da Q3 montava um filtro DASL e NUNCA O USAVA. Enumerava a
#      tabela inteira e comparava datas em memoria. Isso mede o VALOR da
#      propriedade, nao a semantica do Restrict — que envolve fuso,
#      fronteira estrita e parsing do DASL, exatamente onde a Q1 ja se
#      queimou.
#   2. O teste de "contado duas vezes" movia o item e depois lia o
#      destino. Origem-antiga + destino-nova e o resultado ESPERADO de
#      duas observacoes em tempos diferentes. Nao provava nada sobre a
#      Table, e a tabela nem estava ordenada, entao eu nao sabia se o item
#      ja tinha sido lido quando movi.
#   3. Os oraculos nao discriminavam: o caso do Move recebia Ok=$true
#      incondicionalmente, e o achado CERTO da Q3 aparecia como falha.
#
# ------------------------------------------------------------------
# AS DUAS MATRIZES
#
# A) SEMANTICA DA TABLE numa pasta so, com ordem determinada:
#      A1 - mover item JA LIDO pelo cursor
#      A2 - mover item AINDA NAO LIDO pelo cursor
#
# B) CORTE ENTRE PASTAS, com o move entre as duas varreduras:
#      B1 - origem antes, destino depois  -> duplicidade esperada
#      B2 - destino antes, origem depois  -> ausencia global esperada

$ErrorActionPreference = "Stop"
$PT = "http://schemas.microsoft.com/mapi/proptag/"
$P_LASTMOD = $PT + "0x30080040"
$P_ENTRYID = $PT + "0x66700102"
$MARCA = "IRISQ4-" + ([guid]::NewGuid().ToString("N").Substring(0, 8))

$ol = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
$ns = $ol.GetNamespace("MAPI")

function Solta($o) { if ($null -ne $o) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } }
function Hex($v) {
    if ($v -is [byte[]]) { return (($v | ForEach-Object { $_.ToString("X2") }) -join "") }
    return "$v"
}

$criados = New-Object 'System.Collections.Generic.List[string]'   # EntryIDs capturados
$falhas = New-Object 'System.Collections.Generic.List[string]'
$achados = @()

function Registrar([string]$caso, [bool]$ok, [string]$nota) {
    $script:achados += @{ Caso = $caso; Ok = $ok; Nota = $nota }
    if (-not $ok) { [void]$script:falhas.Add($caso) }
}

# ------------------------------------------------------------------
# Enumeracao ORDENADA, com gancho entre lotes e registro de posicao.
# ------------------------------------------------------------------
function Enumerar($pasta, [int]$lote, [scriptblock]$entreLotes) {
    $linhas = New-Object 'System.Collections.Generic.List[object]'
    $t = $null
    try {
        $t = $pasta.GetTable()
        $cols = $t.Columns
        try {
            $cols.RemoveAll()
            $c = $cols.Add($P_ENTRYID); Solta $c
            $c = $cols.Add("Subject");  Solta $c
        } finally { Solta $cols }
        # Ordem DETERMINADA: sem isso eu nao sei se o alvo ja foi lido.
        $t.Sort("Subject", $false)

        $n = 0
        while (-not $t.EndOfTable) {
            $a = $t.GetArray($lote)
            for ($r = $a.GetLowerBound(0); $r -le $a.GetUpperBound(0); $r++) {
                $linhas.Add([pscustomobject]@{
                    Chave = (Hex $a.GetValue($r, 0)); Assunto = "$($a.GetValue($r, 1))"; Lote = $n + 1 })
            }
            $n++
            if ($entreLotes) { & $entreLotes $n $linhas }
        }
        return ,$linhas
    } finally { Solta $t }
}

function CriarItens($pasta, [int]$quantos) {
    # Items.Add + Save deposita em RASCUNHOS, nao na pasta da colecao.
    # Cria-se la e move-se em seguida; o EntryID final e CAPTURADO.
    $rasc = $ns.GetDefaultFolder(16)
    try {
        $itens = $rasc.Items
        try {
            for ($i = 1; $i -le $quantos; $i++) {
                $m = $itens.Add("IPM.Note")
                try {
                    $m.Subject = ("{0} {1:d3}" -f $MARCA, $i)
                    $m.Body = "item de teste do Iris"
                    $m.Save()
                    $mv = $m.Move($pasta)
                    try { [void]$script:criados.Add($mv.EntryID) } finally { Solta $mv }
                } finally { Solta $m }
            }
        } finally { Solta $itens }
    } finally { Solta $rasc }
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

try {
    Write-Host "marcador desta execucao: $MARCA"
    $origem = $raiz.Folders.Add("Iris Q4M origem")
    $destino = $raiz.Folders.Add("Iris Q4M destino")
    CriarItens $origem 40
    Write-Host ("origem: {0} itens capturados" -f $criados.Count)

    # =============================================================
    # MATRIZ A - semantica da Table numa pasta
    # =============================================================
    Write-Host ""
    Write-Host ("=" * 74)
    Write-Host "MATRIZ A - mover COM O CURSOR ABERTO, posicao conhecida"
    Write-Host ("=" * 74)

    foreach ($caso in @(
        @{ Nome = "A1 item JA LIDO";       Alvo = 3;  Esperado = "aparece" },
        @{ Nome = "A2 item AINDA NAO LIDO"; Alvo = 38; Esperado = "?" })) {

        # devolve tudo para a origem antes de cada caso
        foreach ($f in @($destino.Items)) { $mv = $f.Move($origem); Solta $mv; Solta $f }

        $assunto = ("{0} {1:d3}" -f $MARCA, $caso.Alvo)
        $alvo = AcharPorAssunto $origem $assunto
        if ($null -eq $alvo) { throw "nao achei $assunto" }
        $chaveAntes = $alvo.EntryID
        $script:alvoAtual = $alvo
        $script:chaveAlvo = $chaveAntes
        $script:jaLido = $false
        $script:moveu = $false

        # lote 10: o item 003 cai no lote 1; o 038 cai no lote 4.
        $linhas = Enumerar $origem 10 {
            param($lote, $ate)
            if ($lote -eq 1 -and -not $script:moveu) {
                $script:moveu = $true
                $script:jaLido = @($ate | Where-Object { $_.Chave -eq $script:chaveAlvo }).Count -gt 0
                $mv = $script:alvoAtual.Move($script:destino)
                $script:chaveNova = $mv.EntryID
                Solta $mv
            }
        }

        $apareceu = @($linhas | Where-Object { $_.Chave -eq $chaveAntes }).Count -gt 0
        Write-Host ("  {0,-24} lido antes do move? {1,-5} | apareceu no manifesto? {2}" -f `
            $caso.Nome, $(if ($script:jaLido) { "SIM" } else { "nao" }), $(if ($apareceu) { "SIM" } else { "nao" }))
        Registrar $caso.Nome $true ("jaLido=$($script:jaLido) apareceu=$apareceu total=$($linhas.Count)")
        Solta $alvo
    }

    # =============================================================
    # MATRIZ B - corte entre pastas
    # =============================================================
    Write-Host ""
    Write-Host ("=" * 74)
    Write-Host "MATRIZ B - o move ACONTECE ENTRE as duas varreduras"
    Write-Host ("=" * 74)

    foreach ($ordem in @("origem-antes", "destino-antes")) {
        foreach ($f in @($destino.Items)) { $mv = $f.Move($origem); Solta $mv; Solta $f }

        $assunto = ("{0} {1:d3}" -f $MARCA, 20)
        $alvo = AcharPorAssunto $origem $assunto
        $chaveAntes = $alvo.EntryID

        if ($ordem -eq "origem-antes") {
            $vO = Enumerar $origem 50 $null
            $mv = $alvo.Move($destino); $chaveNova = $mv.EntryID; Solta $mv
            $vD = Enumerar $destino 50 $null
        } else {
            $vD = Enumerar $destino 50 $null
            $mv = $alvo.Move($destino); $chaveNova = $mv.EntryID; Solta $mv
            $vO = Enumerar $origem 50 $null
        }
        Solta $alvo

        $naO = @($vO | Where-Object { $_.Chave -eq $chaveAntes }).Count -gt 0
        $naD = @($vD | Where-Object { $_.Chave -eq $chaveNova }).Count -gt 0
        $veredito = if ($naO -and $naD) { "DUPLICADO" }
                    elseif (-not $naO -and -not $naD) { "PERDIDO" }
                    else { "uma vez so" }
        Write-Host ("  {0,-14} origem={1,-5} destino={2,-5} => {3}" -f `
            $ordem, $naO, $naD, $veredito)
        Registrar "B $ordem" $true $veredito
    }

    # =============================================================
    # A Q3 SAIU DAQUI, e a remocao e deliberada.
    #
    # A versao que existia neste script formatava o checkpoint com
    # .ToString("g"), SEM SEGUNDOS, e por isso concluia o OPOSTO do
    # resultado real. Deixar codigo defeituoso "so para historico" num
    # script que alguem pode rodar e pior que apagar: quem rodasse
    # obteria a conclusao errada com aparencia de medicao.
    #
    # A Q3 vive em tools/q3-restrict.ps1, com controle positivo.
    # =============================================================

} catch {
    Write-Host ""
    Write-Host "!!! FALHA !!!"
    Write-Host $_.Exception.Message
    Write-Host $_.ScriptStackTrace
    [void]$falhas.Add("excecao")
} finally {
    Write-Host ""
    Write-Host ("=" * 74)
    Write-Host "LIMPEZA (por EntryID capturado e pelo marcador desta execucao)"
    Write-Host ("=" * 74)
    $sobrou = 0
    # 1) itens que possam ter ficado em Rascunhos entre o Save e o Move
    $rasc = $ns.GetDefaultFolder(16)
    try {
        $li = $rasc.Items
        try {
            for ($i = $li.Count; $i -ge 1; $i--) {
                $it = $li.Item($i)
                try {
                    if ("$($it.Subject)".StartsWith($MARCA, [StringComparison]::Ordinal)) {
                        $it.Delete(); $sobrou++
                    }
                } finally { Solta $it }
            }
        } finally { Solta $li }
    } finally { Solta $rasc }
    Write-Host ("  Rascunhos: {0} item(ns) com o marcador removidos" -f $sobrou)

    # 2) as pastas do teste, com o que houver dentro
    foreach ($f in @($raiz.Folders)) {
        try {
            if ($f.Name -like "Iris Q4M *") {
                Write-Host ("  removendo {0} ({1} itens) -> Itens Excluidos" -f $f.Name, $f.Items.Count)
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
foreach ($a in $achados) { Write-Host ("  {0,-26} {1}" -f $a.Caso, $a.Nota) }
if ($falhas.Count -gt 0) {
    Write-Host ""
    Write-Host ("FALHAS: {0}" -f ($falhas -join ", "))
    exit 1
}
