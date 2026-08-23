# Q2: os itens de conflito sao o MESMO item, ou objetos novos com linhagem
# comum?
#
# SOMENTE LEITURA.
#
# Eu escrevi na secao 10 que eles eram "duas manifestacoes do mesmo item" e
# chamei isso do caso positivo que faltava no corpus. Se forem objetos
# NOVOS, nao ha positivo nenhum, e o unico jeito de obter um continua sendo
# o teste de Move.
#
# Da para decidir lendo tres propriedades:
#
#   PR_CONFLICT_ITEMS (0x10981102) - lista de EntryIDs dos itens em
#       conflito. Se o vencedor apontar para o perdedor, o vinculo esta
#       PROVADO, e provado como VINCULO, nao como identidade.
#   PR_PREDECESSOR_CHANGE_LIST (0x65E30102) - ancestralidade de versoes.
#   PR_RESOLVE_METHOD (0x3FE70003) - como o conflito foi (ou nao) resolvido.
#
# E confere PR_SOURCE_KEY pelo PropertyAccessor: a Table devolveu nulo nos
# 2281, mas "nulo pela Table" nao e a mesma coisa que "ausente no objeto".

$ErrorActionPreference = "Stop"
$PT = "http://schemas.microsoft.com/mapi/proptag/"

$ol = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
$ns = $ol.GetNamespace("MAPI")

function Hex($v) {
    if ($null -eq $v) { return $null }
    if ($v -is [byte[]]) {
        if ($v.Length -eq 0) { return "(vazio)" }
        return (($v | ForEach-Object { $_.ToString("x2") }) -join "")
    }
    # PT_MV_BINARY volta como array DE arrays. Sem este ramo, "$v" devolve
    # a string "System.Byte[]" e o teste de vinculo da "nao" por falha de
    # marshaling — um falso negativo que parece resultado.
    if ($v -is [Array]) {
        $partes = @()
        foreach ($e in $v) {
            if ($e -is [byte[]]) { $partes += (($e | ForEach-Object { $_.ToString("x2") }) -join "") }
            else { $partes += "$e" }
        }
        if ($partes.Count -eq 0) { return "(lista vazia)" }
        return ($partes -join " , ")
    }
    return "$v"
}

# ---------- achar as pastas de conflito e as demais ----------
$conflito = New-Object System.Collections.ArrayList
$outros   = New-Object System.Collections.ArrayList

function Varrer($pasta, [string]$caminho, [int]$prof) {
    if ($prof -gt 12) { Write-Host "  TRUNCADO em $caminho"; return }
    $ehConflito = $caminho -match "Conflitos"
    $t = $null
    try {
        $t = $pasta.GetTable()
        $cols = $t.Columns
        try {
            $cols.RemoveAll()
            [void]$cols.Add("EntryID")
            [void]$cols.Add("Subject")
            [void]$cols.Add(($PT + "0x300B0102"))
        } finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($cols) }

        while (-not $t.EndOfTable) {
            $a = $t.GetArray(200)
            for ($r = $a.GetLowerBound(0); $r -le $a.GetUpperBound(0); $r++) {
                $reg = [pscustomobject]@{
                    Pasta   = $caminho
                    Id      = "$($a.GetValue($r,0))"
                    Assunto = "$($a.GetValue($r,1))"
                    Sk      = Hex $a.GetValue($r,2)
                }
                if ($ehConflito) { [void]$script:conflito.Add($reg) }
                else { [void]$script:outros.Add($reg) }
            }
        }
    } catch {
        Write-Host "  FALHOU $caminho : $($_.Exception.Message)"
    } finally {
        if ($t) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($t) }
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

Write-Host "itens em pastas de Conflitos: $($conflito.Count)"
Write-Host "demais itens: $($outros.Count)"
Write-Host ""

$props = [ordered]@{
    "ConflictItems" = "0x10981102"
    "PredChangeList" = "0x65E30102"
    "ResolveMethod" = "0x3FE70003"
    "SourceKey"     = "0x65E00102"
    "RecordKey"     = "0x0FF90102"
    "MsgFlags"      = "0x0E070003"
}

function Ler($id) {
    $item = $ns.GetItemFromID($id)
    try {
        $pa = $item.PropertyAccessor
        try {
            $r = [ordered]@{}
            foreach ($k in $props.Keys) {
                try { $r[$k] = Hex $pa.GetProperty($PT + $props[$k]) }
                catch { $r[$k] = "(ERRO AO LER: " + $_.Exception.Message.Replace([char]13," ").Replace([char]10," ") + ")" }
            }
            return $r
        } finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pa) }
    } finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($item) }
}

# ---------- para cada item de conflito, achar o parceiro pela SearchKey ----------
$porSk = @{}
foreach ($o in $outros) {
    if ($null -eq $o.Sk) { continue }
    if (-not $porSk.ContainsKey($o.Sk)) { $porSk[$o.Sk] = New-Object System.Collections.ArrayList }
    [void]$porSk[$o.Sk].Add($o)
}

$provados = 0
foreach ($c in $conflito) {
    Write-Host ("=" * 84)
    Write-Host ("ITEM DE CONFLITO: {0}" -f $c.Assunto.Substring(0, [Math]::Min(50, $c.Assunto.Length)))
    Write-Host ("=" * 84)

    $parceiros = if ($null -ne $c.Sk -and $porSk.ContainsKey($c.Sk)) { @($porSk[$c.Sk]) } else { @() }
    Write-Host ("  parceiros com a mesma SearchKey fora de Conflitos: " + @($parceiros).Count)

    $dadosC = Ler $c.Id
    Write-Host "  --- item de conflito ---"
    foreach ($k in $props.Keys) {
        $v = "$($dadosC[$k])"
        if ($v.Length -gt 60) { $v = $v.Substring(0,57) + "..." }
        Write-Host ("     {0,-15} {1}" -f $k, $v)
    }

    foreach ($p in $parceiros) {
        $dadosP = Ler $p.Id
        Write-Host ("  --- parceiro em {0} ---" -f $p.Pasta.Split("/")[-1])
        foreach ($k in $props.Keys) {
            $v = "$($dadosP[$k])"
            if ($v.Length -gt 60) { $v = $v.Substring(0,57) + "..." }
            Write-Host ("     {0,-15} {1}" -f $k, $v)
        }

        # O VINCULO. Comparar com o EntryID nao funciona: PR_CONFLICT_ITEMS
        # guarda um ID de LONGO PRAZO (46 bytes, mesma forma da RecordKey),
        # e o EntryID da Table e de curto prazo (24 bytes). Comparei com o
        # errado na primeira vez e o "nao" pareceu resultado.
        #
        # Os dois IDs longos diferem num byte de TIPO na posicao 20
        # (RecordKey traz 01, o ID de conflito traz 07), entao a comparacao
        # ignora esse byte e exige o resto identico.
        function Nucleo([string]$hex) {
            if ($hex.Length -lt 46) { return $hex }
            return $hex.Substring(0, 40) + "__" + $hex.Substring(42)
        }

        $ciC = Nucleo "$($dadosC['ConflictItems'])".ToLower()
        $ciP = Nucleo "$($dadosP['ConflictItems'])".ToLower()
        $rkC = Nucleo "$($dadosC['RecordKey'])".ToLower()
        $rkP = Nucleo "$($dadosP['RecordKey'])".ToLower()

        Write-Host ""
        Write-Host ("     ConflictItems do conflito : {0}" -f $ciC)
        Write-Host ("     RecordKey    do conflito  : {0}" -f $rkC)
        Write-Host ("     ConflictItems do parceiro : {0}" -f $ciP)
        Write-Host ("     RecordKey    do parceiro  : {0}" -f $rkP)
        Write-Host ""

        # Igualdade tambem nao serve. O blob de PR_CONFLICT_ITEMS e
        # COMPOSTO: cabecalho + par (GUID de replica + contador) da PASTA +
        # par da MENSAGEM. A RecordKey traz o cabecalho e SO o par da
        # mensagem. Entao o vinculo aparece como SUFIXO.
        #
        # Isso, de quebra, mostra o formato da RecordKey: 4 + 16 (GUID do
        # store) + 2 + 16 (GUID de replica) + 8 (contador) = 46 bytes, sem
        # nenhum par de pasta. E argumento de FORMATO, bem melhor que a
        # constancia de bytes que eu tinha usado.
        $parC = $rkC.Substring(44)   # GUID de replica + contador do conflito
        $parP = $rkP.Substring(44)   # idem, do parceiro

        $apontaCP = $ciC.EndsWith($parP)
        $apontaPC = $ciP.EndsWith($parC)
        $apontamSi = ($ciC.EndsWith($parC) -or $ciP.EndsWith($parP))
        Write-Host ("     conflito  aponta para o parceiro : {0}" -f $(if ($apontaCP) { "SIM" } else { "nao" }))
        Write-Host ("     parceiro  aponta para o conflito : {0}" -f $(if ($apontaPC) { "SIM" } else { "nao" }))
        Write-Host ("     algum aponta para SI MESMO       : {0}" -f $(if ($apontamSi) { "SIM" } else { "nao" }))
        if ($apontaCP -or $apontaPC) { $provados++ }
    }
    Write-Host ""
}

Write-Host ("=" * 84)
Write-Host ("vinculos de conflito PROVADOS por PR_CONFLICT_ITEMS: {0}" -f $provados)
Write-Host ""
Write-Host "LEITURA: vinculo provado mostra LINHAGEM entre dois objetos, e nao"
Write-Host "que sao o mesmo objeto. Cada um tem EntryID, RecordKey e estado"
Write-Host "proprios; os dois aparecem no Outlook e podem ser abertos e"
Write-Host "apagados separadamente. Unir os dois numa identidade so seria"
Write-Host "POLITICA do Iris, nao consequencia da medicao."
