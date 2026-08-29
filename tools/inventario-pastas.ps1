# De onde vieram as "129 pastas"?
#
# SOMENTE LEITURA.
#
# O usuario disse que reduziu a organizacao dele para DUAS pastas, e eu
# tinha reportado 129. Uma das duas afirmacoes esta errada, e a minha e a
# suspeita: eu contei tudo que o OOM enumera, sem separar o que e PASTA DO
# USUARIO do que e infraestrutura do Outlook.
#
# O proprio projeto ja sabe fazer essa distincao — FolderVisibilityPolicy e
# o campo IsHidden (PR_ATTR_HIDDEN) existem desde a Fase 1. A varredura da
# Q4 nao usou nenhum dos dois.
#
# Categorias:
#   - OCULTA          : PR_ATTR_HIDDEN. O proprio Outlook nao mostra.
#   - SISTEMA         : Sync Issues, Conversation Action Settings, etc.
#   - NAO-CORREIO     : calendario, contatos, tarefas, notas, diario
#   - EM EXCLUIDOS    : subpasta de Itens Excluidos (inclui artefato MEU)
#   - ARTEFATO DO IRIS: pasta que EU criei em fases anteriores
#   - DO USUARIO      : o que sobra, e e o unico numero que interessa

$ErrorActionPreference = "Stop"
$PT = "http://schemas.microsoft.com/mapi/proptag/"
$P_HIDDEN = $PT + "0x10F4000B"

$ol = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
$ns = $ol.GetNamespace("MAPI")
function Solta($o) { if ($null -ne $o) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } }

$linhas = New-Object 'System.Collections.Generic.List[object]'

# Nomes que o Outlook cria sozinho. Lista conservadora: na duvida, NAO
# classifico como sistema — inflar "sistema" esconderia pasta do usuario.
$sistema = @(
    "Conversation Action Settings", "Quick Step Settings", "Configuracoes de Etapas Rapidas",
    "Configurações de Etapas Rápidas", "Problemas de Sincronização", "Sync Issues",
    "Conflitos", "Conflicts", "Falhas Locais", "Local Failures",
    "Falhas do Servidor", "Server Failures", "Lixo Eletrônico", "Junk E-mail",
    "Caixa de Saida", "Caixa de Saída", "Outbox", "Feeds RSS", "RSS Feeds",
    "Fontes RSS", "Itens Enviados", "Sent Items", "Rascunhos", "Drafts",
    "Caixa de Entrada", "Inbox", "Itens Excluídos", "Deleted Items",
    "Lixeira", "PersonMetadata", "Yammer Root", "ExternalContacts",
    "Files", "Arquivos", "GraphFilesAndWorkingSetSearchFolder",
    "Reminders", "Lembretes", "Chamadas", "Calls"
)

# FALHA ENGOLIDA E PASTA QUE SOME. Tipo, contagem e filhas eram lidos dentro
# de try/catch mudos, e o rodape depois imprimia "nenhuma" -- que e afirmacao
# sobre o que nao foi lido.
$script:semTipo = 0
$script:semContagem = 0
$script:semFilhas = 0
$script:cortadoPorProfundidade = 0   # a arvore continua abaixo e ninguem olhou
$script:semStore = 0                 # um STORE inteiro nao foi percorrido
$script:semOculta = 0                # PR_ATTR_HIDDEN ilegivel: pode virar "DO USUARIO"

function Varrer($pasta, [string]$caminho, [int]$prof, [bool]$sobExcluidos) {
    if ($prof -gt 12) { $script:cortadoPorProfundidade++; return }

    $nome = $pasta.Name
    $oculta = $false
    $pa = $null
    try {
        $pa = $pasta.PropertyAccessor
        try { $oculta = [bool]$pa.GetProperty($P_HIDDEN) } catch { $script:semOculta++ }
    } catch { $script:semOculta++ } finally { Solta $pa }

    $tipo = -1
    try { $tipo = [int]$pasta.DefaultItemType } catch { $script:semTipo++ }
    $n = 0
    try { $itens = $pasta.Items; $n = $itens.Count; Solta $itens } catch { $script:semContagem++ }

    $cat =
        if ($nome -like "Iris*") { "ARTEFATO DO IRIS" }
        elseif ($oculta) { "OCULTA" }
        elseif ($sobExcluidos) { "EM EXCLUIDOS" }
        elseif ($sistema -contains $nome) { "SISTEMA" }
        elseif ($tipo -ne 0) { "NAO-CORREIO" }
        else { "DO USUARIO" }

    $script:linhas.Add([pscustomobject]@{
        Caminho = $caminho; Nome = $nome; Cat = $cat; Itens = $n; Prof = $prof })

    $filhas = $null
    try { $filhas = $pasta.Folders } catch { $script:semFilhas++; return }
    try {
        for ($k = 1; $k -le $filhas.Count; $k++) {
            $f = $filhas.Item($k)
            try {
                $ehExcluidos = ($f.Name -eq "Itens Excluídos" -or $f.Name -eq "Deleted Items")
                Varrer $f "$caminho/$($f.Name)" ($prof + 1) ($sobExcluidos -or $ehExcluidos)
            } finally { Solta $f }
        }
    } finally { Solta $filhas }
}

$stores = $ns.Stores
for ($s = 1; $s -le $stores.Count; $s++) {
    $store = $stores.Item($s)
    $raiz = $null
    # O catch pega o GetRootFolder E TODA A RECURSAO. Dizer "nao consegui
    # abrir a raiz" descreveria errado uma falha tardia, depois de meia
    # arvore lida -- entao o rodape diz o que se sabe: a travessia parou.
    try { $raiz = $store.GetRootFolder(); Varrer $raiz "" 0 $false }
    catch { $script:semStore++ } finally {
        if ($raiz) { Solta $raiz }
        Solta $store
    }
}
Solta $stores

Write-Host ("TOTAL enumerado pelo OOM: {0} pastas" -f $linhas.Count)
Write-Host ""
Write-Host "POR CATEGORIA:"
foreach ($g in ($linhas | Group-Object Cat | Sort-Object Count -Descending)) {
    Write-Host ("  {0,-18} {1,4} pastas, {2,6} itens" -f `
        $g.Name, $g.Count, (($g.Group | Measure-Object Itens -Sum).Sum))
}

Write-Host ""
# O TOTAL DE CEGUEIRAS E CALCULADO AQUI, antes de qualquer secao consultar.
# Ele estava sendo somado la embaixo, DEPOIS da secao "PASTAS DO USUARIO" que
# o consulta -- entao aquela secao lia $null e imprimia "nenhuma" mesmo com
# falha de leitura. Consertar o inventario de cegueiras e deixar a primeira
# consulta cega e uma piada que a revisao externa nao deixou passar.
$script:cego = $script:semTipo + $script:semContagem + $script:semFilhas +
               $script:cortadoPorProfundidade + $script:semStore + $script:semOculta

Write-Host ("=" * 72)
Write-Host "PASTAS DO USUARIO (as unicas que interessam ao cache)"
Write-Host ("=" * 72)
$doUsuario = @($linhas | Where-Object { $_.Cat -eq "DO USUARIO" })
if ($doUsuario.Count -eq 0) {
    if ($script:cego -gt 0) {
        Write-Host "  nenhuma FOI LIDA -- e houve falha de leitura (ver o rodape)."
    } else {
        Write-Host "  nenhuma"
    }
} else {
    foreach ($l in $doUsuario) {
        Write-Host ("  {0,6} itens  {1}" -f $l.Itens, $l.Caminho)
    }
}

Write-Host ""
if ($script:cego -gt 0) {
    Write-Host "O QUE ESTE INVENTARIO NAO VIU:" -ForegroundColor DarkYellow
    if ($script:semTipo -gt 0)     { Write-Host ("  {0} pasta(s): nao consegui ler o tipo" -f $script:semTipo) }
    if ($script:semContagem -gt 0) { Write-Host ("  {0} pasta(s): nao consegui contar os itens (aparecem com 0)" -f $script:semContagem) }
    if ($script:semFilhas -gt 0)   { Write-Host ("  {0} ramo(s): nao consegui enumerar as filhas" -f $script:semFilhas) }
    if ($script:cortadoPorProfundidade -gt 0) { Write-Host ("  {0} ramo(s): cortados na profundidade 12" -f $script:cortadoPorProfundidade) }
    if ($script:semStore -gt 0)    { Write-Host ("  {0} store(s): a travessia parou no meio ou nem comecou" -f $script:semStore) }
    if ($script:semOculta -gt 0)   { Write-Host ("  {0} pasta(s): PR_ATTR_HIDDEN ilegivel (podem ter virado 'DO USUARIO')" -f $script:semOculta) }
    Write-Host "  Zero nas linhas acima pode ser 'nao contei', e nao 'nao tem'."
    Write-Host ""
}
Write-Host ("=" * 72)
Write-Host "ARTEFATOS QUE **EU** DEIXEI"
Write-Host ("=" * 72)
$meus = @($linhas | Where-Object { $_.Cat -eq "ARTEFATO DO IRIS" })
if ($meus.Count -eq 0) {
    if ($script:cego -gt 0) {
        Write-Host "  nenhum FOI LIDO -- e houve falha de leitura (ver acima)."
    } else {
        Write-Host "  nenhum"
    }
}
else {
    foreach ($l in $meus) { Write-Host ("  {0,6} itens  {1}" -f $l.Itens, $l.Caminho) }
    Write-Host ("  --> {0} pastas, {1} itens, TODOS meus" -f `
        $meus.Count, (($meus | Measure-Object Itens -Sum).Sum))
}

Write-Host ""
Write-Host ("=" * 72)
Write-Host "EM ITENS EXCLUIDOS (subpastas)"
Write-Host ("=" * 72)
foreach ($l in ($linhas | Where-Object { $_.Cat -eq "EM EXCLUIDOS" } | Sort-Object Itens -Descending | Select-Object -First 20)) {
    Write-Host ("  {0,6} itens  {1}" -f $l.Itens, $l.Caminho)
}
