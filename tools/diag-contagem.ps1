# A Caixa de Entrada tem 1.004 itens ou 17.668?
#
# SOMENTE LEITURA.
#
# A barra de status do Outlook mostra "Itens: 17.668" com a Caixa de
# Entrada selecionada. Todas as minhas medicoes da Fase 2 usaram
# GetDefaultFolder(6).Items.Count, que devolveu 1.004.
#
# Se o numero certo for 17.668, entao TUDO que eu medi — o ganho de 10,6x
# do ReadPage, os ~3,2 s de varredura completa, o custo por pagina — esta
# medido sobre 6% da caixa, e as conclusoes de dimensionamento nao valem.
#
# Hipoteses a distinguir:
#   H1 - modo cache com JANELA DE SINCRONIZACAO: o OST guarda so parte, e
#        o OOM enumera o OST. A barra mostraria o total do servidor.
#   H2 - a pasta que eu pego com GetDefaultFolder(6) NAO e a que esta
#        selecionada na tela.
#   H3 - ha mais de um store / conta, e eu so vi um.
#   H4 - a barra conta algo diferente (conversas, ou outra pasta).

$ErrorActionPreference = "Stop"
$PT = "http://schemas.microsoft.com/mapi/proptag/"

$ol = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
$ns = $ol.GetNamespace("MAPI")
function Solta($o) { if ($null -ne $o) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } }

Write-Host "=== STORES ==="
$stores = $ns.Stores
Write-Host ("total: {0}" -f $stores.Count)
for ($s = 1; $s -le $stores.Count; $s++) {
    $st = $stores.Item($s)
    $tipo = "?"
    try { $tipo = $st.ExchangeStoreType } catch { }
    $cache = "?"
    try { $cache = $st.IsCachedExchange } catch { }
    Write-Host ("  [{0}] {1}" -f $s, $st.DisplayName)
    Write-Host ("      ExchangeStoreType={0}  IsCachedExchange={1}" -f $tipo, $cache)
    try { Write-Host ("      arquivo: {0}" -f $st.FilePath) } catch { }
    Solta $st
}
Solta $stores
Write-Host ""

Write-Host "=== A PASTA QUE O USUARIO ESTA VENDO ==="
try {
    $exp = $ol.ActiveExplorer()
    if ($null -ne $exp) {
        $atual = $exp.CurrentFolder
        Write-Host ("  nome        : {0}" -f $atual.Name)
        Write-Host ("  caminho     : {0}" -f $atual.FolderPath)
        Write-Host ("  Items.Count : {0}" -f $atual.Items.Count)
        $pa = $atual.PropertyAccessor
        foreach ($p in @(@{N="PR_CONTENT_COUNT";T="0x36020003"},
                         @{N="PR_CONTENT_UNREAD";T="0x36030003"},
                         @{N="PR_ASSOC_CONTENT_COUNT";T="0x36170003"})) {
            try { Write-Host ("  {0,-24}: {1}" -f $p.N, $pa.GetProperty($PT + $p.T)) }
            catch { Write-Host ("  {0,-24}: (erro)" -f $p.N) }
        }
        Solta $pa
        $mesma = ($atual.EntryID -eq $ns.GetDefaultFolder(6).EntryID)
        Write-Host ("  e a MESMA de GetDefaultFolder(6)? {0}" -f $(if ($mesma) { "SIM" } else { "NAO" }))
        Solta $atual
        Solta $exp
    } else { Write-Host "  (sem Explorer ativo)" }
} catch { Write-Host ("  erro: {0}" -f $_.Exception.Message) }
Write-Host ""

Write-Host "=== CAIXA DE ENTRADA por GetDefaultFolder(6) ==="
$in = $ns.GetDefaultFolder(6)
Write-Host ("  caminho     : {0}" -f $in.FolderPath)
$itens = $in.Items
Write-Host ("  Items.Count : {0}" -f $itens.Count)
Solta $itens
$pa = $in.PropertyAccessor
foreach ($p in @(@{N="PR_CONTENT_COUNT";T="0x36020003"},
                 @{N="PR_CONTENT_UNREAD";T="0x36030003"})) {
    try { Write-Host ("  {0,-20}: {1}" -f $p.N, $pa.GetProperty($PT + $p.T)) }
    catch { Write-Host ("  {0,-20}: (erro)" -f $p.N) }
}
Solta $pa

# A Table ve o mesmo que o Items?
$t = $in.GetTable()
$cols = $t.Columns; $cols.RemoveAll(); $c = $cols.Add("EntryID"); Solta $c; Solta $cols
$n = 0
while (-not $t.EndOfTable) { $a = $t.GetArray(500); $n += ($a.GetUpperBound(0) - $a.GetLowerBound(0) + 1) }
Solta $t
Write-Host ("  Table (varredura completa): {0} linhas" -f $n)

# item mais antigo e mais novo: revela janela de sincronizacao
$itens = $in.Items
$itens.Sort("[ReceivedTime]", $false)
try { Write-Host ("  item MAIS ANTIGO no OST : {0}" -f $itens.Item(1).ReceivedTime) } catch { }
$itens.Sort("[ReceivedTime]", $true)
try { Write-Host ("  item MAIS NOVO no OST   : {0}" -f $itens.Item(1).ReceivedTime) } catch { }
Solta $itens
Solta $in
Write-Host ""

Write-Host "=== SUBPASTAS DA CAIXA DE ENTRADA ==="
$in = $ns.GetDefaultFolder(6)
foreach ($f in $in.Folders) {
    Write-Host ("  {0,7} itens  {1}" -f $f.Items.Count, $f.Name)
    Solta $f
}
Solta $in
Write-Host ""
Write-Host "LEITURA: se Items.Count, PR_CONTENT_COUNT e a Table concordarem em"
Write-Host "1.004 e a tela disser 17.668, a diferenca esta FORA do OOM — e a"
Write-Host "hipotese H1 (janela de sincronizacao do modo cache) e a candidata."
