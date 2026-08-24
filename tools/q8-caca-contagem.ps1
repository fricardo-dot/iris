# Qual propriedade MAPI carrega a contagem do SERVIDOR?
#
# A aba Sincronizacao do Outlook mostra, para "1. Backup":
#   Pasta do servidor contem: 145
#   Pasta offline contem:      35
#
# O OOM devolve 35 por tres caminhos diferentes (Items.Count,
# PR_CONTENT_COUNT, Table) - os tres lendo o mesmo OST. O 145 existe, o
# Outlook o exibe, e nao sei por onde.
#
# Em vez de chutar a proptag pela memoria - que e o modo de errar deste
# projeto, e ja custou o 00036601 - eu PROCURO PELO VALOR: varro as faixas
# de proptag de pasta com tipo PT_LONG e vejo qual devolve o numero da tela.
#
# Duas pastas com alvos diferentes dao o controle: uma propriedade que
# devolva 145 em "1. Backup" E 17.728 na Caixa de Entrada e a resposta.
# Uma que devolva 145 numa e outra coisa na outra e coincidencia.
$ErrorActionPreference = "Stop"
$ol = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
$ns = $ol.GetNamespace("MAPI")
$PT = "http://schemas.microsoft.com/mapi/proptag/"

# Faixas plausiveis para propriedade de pasta:
#   0x0E00-0x0EFF  status de mensagem/pasta
#   0x3600-0x36FF  propriedades de pasta
#   0x6600-0x67FF  provider, replicacao, sincronizacao
$faixas = @(@(0x0E00, 0x0EFF), @(0x3600, 0x36FF), @(0x6600, 0x67FF))
$tipoLong = 0x0003

$alvos = @{ "1. Backup" = 145; "Caixa de Entrada" = 17728 }

function Achar($pasta) {
    $encontrados = @{}
    $pa = $pasta.PropertyAccessor
    try {
        foreach ($f in $faixas) {
            for ($id = $f[0]; $id -le $f[1]; $id++) {
                $tag = "{0}0x{1:X4}{2:X4}" -f $PT, $id, $tipoLong
                try {
                    $v = $pa.GetProperty($tag)
                    if ($v -is [int] -or $v -is [long]) { $encontrados["0x{0:X4}" -f $id] = [int]$v }
                } catch { }
            }
        }
    } finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pa) }
    return $encontrados
}

$porPasta = @{}
function Varrer($pasta, $prof) {
    if ($prof -gt 6) { return }
    if ($alvos.ContainsKey($pasta.Name)) {
        Write-Host ("varrendo '{0}' (alvo: {1})..." -f $pasta.Name, $alvos[$pasta.Name])
        $script:porPasta[$pasta.Name] = Achar $pasta
        Write-Host ("  {0} propriedades PT_LONG legiveis" -f $script:porPasta[$pasta.Name].Count)
    }
    $filhas = $null
    try {
        $filhas = $pasta.Folders; $c = $filhas.Count
        for ($i = 1; $i -le $c; $i++) {
            $f = $null
            try { $f = $filhas.Item($i); Varrer $f ($prof + 1) }
            finally { if ($f) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($f) } }
        }
    } catch { } finally { if ($filhas) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($filhas) } }
}

foreach ($st in $ns.Stores) {
    $raiz = $null
    try { $raiz = $st.GetRootFolder(); Varrer $raiz 0 }
    finally { if ($raiz) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($raiz) } }
}

Write-Host ""
Write-Host "=== propriedades que batem com o alvo, POR PASTA ==="
foreach ($nome in $alvos.Keys) {
    if (-not $porPasta.ContainsKey($nome)) { Write-Host "  '$nome' nao encontrada"; continue }
    $alvo = $alvos[$nome]
    $bate = $porPasta[$nome].GetEnumerator() | Where-Object { $_.Value -eq $alvo }
    if ($bate) {
        foreach ($b in $bate) { Write-Host ("  '{0}': {1} = {2}  <-- ALVO" -f $nome, $b.Key, $b.Value) }
    } else {
        Write-Host ("  '{0}': NENHUMA propriedade devolve {1}" -f $nome, $alvo)
    }
}

Write-Host ""
Write-Host "=== propriedades que batem nas DUAS (a resposta, se houver) ==="
$nomes = @($alvos.Keys)
if ($porPasta.Count -eq 2) {
    $a = $porPasta[$nomes[0]]; $b = $porPasta[$nomes[1]]
    $resp = $a.Keys | Where-Object {
        $a[$_] -eq $alvos[$nomes[0]] -and $b.ContainsKey($_) -and $b[$_] -eq $alvos[$nomes[1]]
    }
    if ($resp) { foreach ($r in $resp) { Write-Host "  $r" } }
    else { Write-Host "  nenhuma" }
}

Write-Host ""
Write-Host "=== diferencas entre as duas pastas (para inspecao) ==="
if ($porPasta.Count -eq 2) {
    $a = $porPasta[$nomes[0]]; $b = $porPasta[$nomes[1]]
    foreach ($k in ($a.Keys | Sort-Object)) {
        if ($b.ContainsKey($k) -and $a[$k] -ne $b[$k]) {
            Write-Host ("  {0}: {1}={2}  {3}={4}" -f $k, $nomes[0], $a[$k], $nomes[1], $b[$k])
        }
    }
}
