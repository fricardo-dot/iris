# Duas medicoes que decidem o desenho do ReadPage por cursor.
#
# SOMENTE LEITURA.
#
# 1. ReceivedTime NULO. Qualquer filtro "< data" exclui linha nula, entao
#    depois da primeira pagina um item sem ReceivedTime some da travessia
#    em silencio. A FASE2 ja registrou isso como lacuna; aqui eu meco se e
#    problema real nesta caixa ou hipotese.
#
# 2. MAIOR EMPATE. A pagina do algoritmo drenado NAO tem teto: ela vai ate
#    o fim do grupo do ultimo instante. Se existir um grupo enorme, uma
#    unica pagina trava a fila da STA e a UI para. Preciso do maior grupo
#    de TODAS as pastas, nao so da Entrada.
#
# 3. De quebra: MessageClass x TryCast(MailItem). O caminho por Table nao
#    tem TryCast, e filtrar por MessageClass NAO e equivalente. Conto quais
#    classes existem para dimensionar a politica.

$ErrorActionPreference = "Stop"

$ol = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
$ns = $ol.GetNamespace("MAPI")

$nulos = 0
$total = 0
$porPastaNula = @{}
$classes = @{}
$maiorEmpate = 0
$ondeEmpate = ""
$empatesGrandes = @()

function Varrer($pasta, [string]$caminho, [int]$prof) {
    if ($prof -gt 12) { return }
    $t = $null
    try {
        $t = $pasta.GetTable()
        $cols = $t.Columns
        try {
            $cols.RemoveAll()
            [void]$cols.Add("EntryID")
            [void]$cols.Add("ReceivedTime")
            [void]$cols.Add("MessageClass")
        } finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($cols) }

        $instantes = @{}
        while (-not $t.EndOfTable) {
            $a = $t.GetArray(200)
            for ($r = $a.GetLowerBound(0); $r -le $a.GetUpperBound(0); $r++) {
                $script:total++
                $quando = $a.GetValue($r, 1)
                $classe = "$($a.GetValue($r, 2))"
                $script:classes[$classe] = 1 + [int]$script:classes[$classe]

                if ($null -eq $quando) {
                    $script:nulos++
                    $script:porPastaNula[$caminho] = 1 + [int]$script:porPastaNula[$caminho]
                } else {
                    # Empate = mesmo instante ate o SEGUNDO, que e a
                    # granularidade do filtro DASL.
                    $chave = ([datetime]$quando).ToString("yyyyMMddHHmmss")
                    $instantes[$chave] = 1 + [int]$instantes[$chave]
                }
            }
        }
        foreach ($k in $instantes.Keys) {
            if ($instantes[$k] -gt $script:maiorEmpate) {
                $script:maiorEmpate = $instantes[$k]
                $script:ondeEmpate = $caminho
            }
            if ($instantes[$k] -ge 10) {
                $script:empatesGrandes += [pscustomobject]@{
                    Pasta = $caminho.Split("/")[-1]; Quando = $k; N = $instantes[$k] }
            }
        }
    } catch {
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
    catch { }
    finally {
        if ($raiz) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($raiz) }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($store)
    }
}
[void][Runtime.InteropServices.Marshal]::ReleaseComObject($stores)

Write-Host "itens: $total"
Write-Host ""
Write-Host ("1. ReceivedTime NULO: {0} de {1}" -f $nulos, $total)
if ($nulos -gt 0) {
    foreach ($k in ($porPastaNula.Keys | Sort-Object { -$porPastaNula[$_] })) {
        Write-Host ("      {0,4}x  {1}" -f $porPastaNula[$k], $k.Split("/")[-1])
    }
    Write-Host "   => BLOQUEADOR REAL: filtro '< data' perde estes itens."
} else {
    Write-Host "   => nenhum ENTRE OS ITENS QUE ESTE ROTEIRO LEU. Continua sendo"
    Write-Host "      lacuna do CONTRATO, porque outra caixa pode ter -- e porque"
    Write-Host "      este roteiro engole falha de pasta e de store, entao o"
    Write-Host "      'nenhum' e sobre o que foi lido, e nao sobre a caixa."
}
Write-Host ""

# TODO NUMERO DAQUI PARA BAIXO E SOBRE O CORPUS QUE ESTE ROTEIRO CONSEGUIU
# LER: ele corta em profundidade 12 e engole falha de tabela, de filhas e de
# store. Maximo observado nao e maximo existente.
Write-Host ("2. MAIOR EMPATE OBSERVADO no mesmo segundo: {0} itens" -f $maiorEmpate)
Write-Host ("   em: {0}" -f $ondeEmpate.Split("/")[-1])
if ($empatesGrandes.Count -gt 0) {
    Write-Host "   grupos com 10 ou mais:"
    $empatesGrandes | Sort-Object N -Descending | Select-Object -First 10 | ForEach-Object {
        Write-Host ("      {0,4} itens  {1}  em {2}" -f $_.N, $_.Quando, $_.Pasta)
    }
} else {
    Write-Host "   nenhum grupo com 10 ou mais ENTRE OS ITENS QUE ESTE ROTEIRO LEU."
}
Write-Host "   => a pagina drenada devolve ate PAGINA + (empate - 1) itens."
Write-Host ""

Write-Host "3. MessageClass presentes:"
$classes.Keys | Sort-Object { -$classes[$_] } | Select-Object -First 12 | ForEach-Object {
    Write-Host ("      {0,5}x  {1}" -f $classes[$_], $_)
}
