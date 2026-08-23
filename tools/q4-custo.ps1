# Q4, parte 1: custo de enumerar LOCALIZADOR ENTRE SESSOES +
# LastModificationTime.
#
# NAO se chama "chave duravel": a 11.1 mediu que ela MUDA no Move. Ela
# sobrevive ao fim da SESSAO, e nada alem disso. O nome errado convida a
# usa-la como chave primaria no 2.1, que e o erro que a Q2 existe para
# impedir.
#
# SOMENTE LEITURA.
#
# 2a versao. A 1a tinha tres defeitos que o Codex pegou, e o primeiro e
# ironico:
#
#   1. Usava a coluna "EntryID", que a secao 12 do FASE2 acabou de
#      descobrir que nao sobrevive a sessao — ela devolve o EntryID de
#      curto prazo, valido so na sessao. Medir o custo de enumerar uma
#      chave que nao serve nao mede nada.
#   2. Engolia qualquer falha de pasta num catch vazio, e ainda contava a
#      pasta como percorrida. "129 pastas" podia incluir pasta nao lida.
#   3. Conferia unicidade so nas 5 pastas padrao, e com um HashSet POR
#      PASTA — o que nao responde a pergunta, que e se a chave e unica na
#      CAIXA.
#
# E ele nao comparava as duas rodadas. Custo estavel com CONTEUDO diferente
# seria pior que custo instavel.

$ErrorActionPreference = "Stop"
$PT = "http://schemas.microsoft.com/mapi/proptag/"
$P_LASTMOD  = $PT + "0x30080040"   # PR_LAST_MODIFICATION_TIME
$P_ENTRYID  = $PT + "0x66700102"   # PR_LONGTERM_ENTRYID_FROM_TABLE

$ol = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
$ns = $ol.GetNamespace("MAPI")

function Add-Coluna($colunas, [string]$dasl) {
    # Columns.Add DEVOLVE um objeto COM. Ignorar o retorno deixa RCW sem dono.
    $c = $null
    try { $c = $colunas.Add($dasl) }
    finally { if ($c) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($c) } }
}

function Enumerar($pasta) {
    # Devolve as chaves e o tempo. LANCA se a pasta nao puder ser lida:
    # numero que ignora pasta ilegivel nao e custo, e amostra.
    $chaves = New-Object 'System.Collections.Generic.List[string]'
    $t = $null
    $sw = [Diagnostics.Stopwatch]::StartNew()
    try {
        $t = $pasta.GetTable()
        $cols = $t.Columns
        try {
            $cols.RemoveAll()
            Add-Coluna $cols $P_ENTRYID
            Add-Coluna $cols $P_LASTMOD
        } finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($cols) }

        while (-not $t.EndOfTable) {
            $a = $t.GetArray(200)
            for ($r = $a.GetLowerBound(0); $r -le $a.GetUpperBound(0); $r++) {
                $v = $a.GetValue($r, 0)
                if ($v -is [byte[]]) {
                    $chaves.Add((($v | ForEach-Object { $_.ToString("X2") }) -join ""))
                } else {
                    # Coluna aceita devolvendo outra coisa e o falso positivo
                    # que a Q1 teve com Permission. Nao conta como chave.
                    $chaves.Add("__NAO_BINARIO__")
                }
            }
        }
        $sw.Stop()
        return @{ Chaves = $chaves; Ms = $sw.Elapsed.TotalMilliseconds }
    } finally {
        if ($t) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($t) }
    }
}

# ------------------------------------------------------------------
Write-Host "PASTAS PADRAO, tres execucoes cada"
Write-Host ("{0,-22} | {1,6} | {2,8} | {3,8} | {4,8}" -f "pasta", "itens", "1a (ms)", "2a (ms)", "3a (ms)")
Write-Host ("-" * 66)

foreach ($p in @(
    @{ Id = 6;  Nome = "Caixa de Entrada" },
    @{ Id = 5;  Nome = "Itens Enviados"   },
    @{ Id = 16; Nome = "Rascunhos"        },
    @{ Id = 3;  Nome = "Itens Excluidos"  },
    @{ Id = 23; Nome = "Lixo Eletronico"  })) {

    $pasta = $ns.GetDefaultFolder($p.Id)
    try {
        $r = @(1..3 | ForEach-Object { Enumerar $pasta })
        Write-Host ("{0,-22} | {1,6} | {2,8:N1} | {3,8:N1} | {4,8:N1}" -f `
            $p.Nome, $r[0].Chaves.Count, $r[0].Ms, $r[1].Ms, $r[2].Ms)
    } finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pasta) }
}
Write-Host ""

# ------------------------------------------------------------------
# A CAIXA INTEIRA. Pasta que falha e REGISTRADA, nao engolida, e a
# contagem de pastas so conta as que foram lidas ate o fim.
# ------------------------------------------------------------------
$script:manifesto = $null
$script:lidas = 0
$script:falhas = New-Object 'System.Collections.Generic.List[string]'
$script:profundidadeMax = 0
$script:truncadas = 0

function Varrer($pasta, [string]$caminho, [int]$prof) {
    if ($prof -gt $script:profundidadeMax) { $script:profundidadeMax = $prof }
    if ($prof -gt 12) {
        $script:truncadas++
        return
    }
    try {
        $r = Enumerar $pasta
        foreach ($k in $r.Chaves) { [void]$script:manifesto.Add($k) }
        $script:lidas++
    } catch {
        [void]$script:falhas.Add("$caminho :: $($_.Exception.Message)")
    }

    $filhas = $null
    try { $filhas = $pasta.Folders } catch { return }
    try {
        for ($k = 1; $k -le $filhas.Count; $k++) {
            $f = $filhas.Item($k)
            try { Varrer $f "$caminho/$($f.Name)" ($prof + 1) }
            finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($f) }
        }
    } finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($filhas) }
}

$rodadas = @()
for ($n = 1; $n -le 3; $n++) {
    $script:manifesto = New-Object 'System.Collections.Generic.List[string]'
    $script:lidas = 0
    $script:falhas = New-Object 'System.Collections.Generic.List[string]'
    $script:truncadas = 0

    $sw = [Diagnostics.Stopwatch]::StartNew()
    $stores = $ns.Stores
    for ($s = 1; $s -le $stores.Count; $s++) {
        $store = $stores.Item($s)
        $raiz = $null
        try { $raiz = $store.GetRootFolder(); Varrer $raiz $store.DisplayName 0 }
        catch { [void]$script:falhas.Add("store $s :: $($_.Exception.Message)") }
        finally {
            if ($raiz) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($raiz) }
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($store)
        }
    }
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($stores)
    $sw.Stop()

    $rodadas += @{
        Ms = $sw.Elapsed.TotalMilliseconds
        Lidas = $script:lidas
        Falhas = @($script:falhas)
        Itens = $script:manifesto.Count
        Manifesto = $script:manifesto
    }
    Write-Host ("CAIXA INTEIRA, rodada {0}: {1} pastas LIDAS, {2} falharam, {3} itens, {4:N0} ms" -f `
        $n, $script:lidas, $script:falhas.Count, $script:manifesto.Count, $sw.Elapsed.TotalMilliseconds)
}

Write-Host ""
Write-Host ("profundidade maxima alcancada: {0} (truncadas: {1})" -f $profundidadeMax, $truncadas)

if ($rodadas[0].Falhas.Count -gt 0) {
    Write-Host ""
    Write-Host "PASTAS QUE FALHARAM (o custo acima NAO as inclui):"
    $rodadas[0].Falhas | Select-Object -First 8 | ForEach-Object { Write-Host "   $_" }
}

# ------------------------------------------------------------------
Write-Host ""
Write-Host "UNICIDADE, na CAIXA inteira e nao por pasta:"
$distintas = New-Object 'System.Collections.Generic.HashSet[string]'
$repetidas = 0
$naoBinarias = 0
foreach ($k in $rodadas[0].Manifesto) {
    if ($k -eq "__NAO_BINARIO__") { $naoBinarias++; continue }
    if (-not $distintas.Add($k)) { $repetidas++ }
}
Write-Host ("   itens {0} | chaves distintas {1} | REPETIDAS {2} | nao binarias {3}" -f `
    $rodadas[0].Itens, $distintas.Count, $repetidas, $naoBinarias)

Write-Host ""
Write-Host "AS DUAS RODADAS OBSERVARAM O MESMO?"
Write-Host "(custo estavel sobre conteudo diferente nao mede o mesmo trabalho)"
$a = New-Object 'System.Collections.Generic.HashSet[string]'
foreach ($k in $rodadas[0].Manifesto) { [void]$a.Add($k) }
$b = New-Object 'System.Collections.Generic.HashSet[string]'
foreach ($k in $rodadas[2].Manifesto) { [void]$b.Add($k) }
$soEmA = ($a | Where-Object { -not $b.Contains($_) }).Count
$soEmB = ($b | Where-Object { -not $a.Contains($_) }).Count
Write-Host ("   so na 1a: {0}   so na 3a: {1}" -f $soEmA, $soEmB)
if ($soEmA -eq 0 -and $soEmB -eq 0) {
    Write-Host "   => manifestos IDENTICOS: o custo se refere ao mesmo trabalho."
} else {
    Write-Host "   => a caixa mudou entre as rodadas; o custo nao e comparavel item a item."
}

Write-Host ""
Write-Host "LEITURA: este e o custo BRUTO em PowerShell, do subconjunto que"
Write-Host "deu para ler. Nao prova que duas varreduras cabem no orcamento da"
Write-Host "STA do Iris, onde elas dividem a fila com a UI."
