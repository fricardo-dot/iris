# Q2: a PCL preserva ANCESTRALIDADE atraves do Move?
#
# ESCREVE. Mesma autorizacao do q2-move.ps1, e usa SO o artefato meu.
#
# Por que este teste existe: no q2-move.ps1 eu comparei
# PR_PREDECESSOR_CHANGE_LIST por IGUALDADE, vi que mudava, e escrevi
# "muda". A PCL e uma LISTA de change keys antecessoras — comparar por
# igualdade e a pergunta errada. A certa e CONTINENCIA:
#
#     a PCL depois do Move contem a ChangeKey de ANTES do Move?
#
# Se contiver, existe continuidade causal mesmo sem nenhuma chave igual, e
# a secao 11.3 ("item movido e indistinguivel de apagado + chegou") esta
# forte demais.
#
# Controle negativo: a mesma pergunta para o Copy. Se a copia tambem herdar
# a ancestralidade, a PCL nao separa Move de Copy sozinha, e a distincao
# passa a exigir observar COEXISTENCIA.
#
# LIMPEZA: a copia sai por MOVE explicito para Itens Excluidos, nao por
# Delete(). Delete() dentro de Itens Excluidos pode ser permanente, e no
# experimento anterior o Delete() nao teve pos-condicao conferida.

$ErrorActionPreference = "Stop"
$PT = "http://schemas.microsoft.com/mapi/proptag/"
$P_CHANGE = $PT + "0x65E20102"
$P_PCL    = $PT + "0x65E30102"
$P_RECORD = $PT + "0x0FF90102"

$ol = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
$ns = $ol.GetNamespace("MAPI")

function Hex($v) {
    if ($null -eq $v) { throw "propriedade obrigatoria voltou nula" }
    if (-not ($v -is [byte[]])) { throw "tipo inesperado: $($v.GetType().Name)" }
    if ($v.Length -eq 0) { throw "propriedade obrigatoria voltou vazia" }
    return (($v | ForEach-Object { $_.ToString("x2") }) -join "")
}

# Erro NAO vira string aqui. No q2-move.ps1 uma falha de leitura virava
# "(erro)", e dois "(erro)" comparavam IGUAIS — o que produziria um
# "sobreviveu" falso.
function Foto($item) {
    $pa = $item.PropertyAccessor
    try {
        return [pscustomobject]@{
            EntryID   = $item.EntryID
            ChangeKey = Hex $pa.GetProperty($P_CHANGE)
            Pcl       = Hex $pa.GetProperty($P_PCL)
            RecordKey = Hex $pa.GetProperty($P_RECORD)
            Pasta     = $item.Parent.Name
        }
    } finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pa) }
}

function Analisar($antes, $depois, [string]$oQue) {
    Write-Host ""
    Write-Host ("=" * 74)
    Write-Host $oQue
    Write-Host ("=" * 74)
    Write-Host ("  pasta      : {0} -> {1}" -f $antes.Pasta, $depois.Pasta)
    Write-Host ("  ChangeKey  : {0}" -f $(if ($antes.ChangeKey -eq $depois.ChangeKey) { "IGUAL" } else { "mudou" }))
    Write-Host ("  RecordKey  : {0}" -f $(if ($antes.RecordKey -eq $depois.RecordKey) { "IGUAL" } else { "mudou" }))
    Write-Host ""
    Write-Host ("  ChangeKey ANTES : {0}" -f $antes.ChangeKey)
    Write-Host ("  PCL antes ({0,3} bytes): {1}" -f ($antes.Pcl.Length/2), $antes.Pcl)
    Write-Host ("  PCL depois({0,3} bytes): {1}" -f ($depois.Pcl.Length/2), $depois.Pcl)
    Write-Host ""

    # A PERGUNTA CERTA: continencia, nao igualdade.
    $contemChave = $depois.Pcl.Contains($antes.ChangeKey)
    $contemPcl   = $depois.Pcl.Contains($antes.Pcl)
    $cresceu     = $depois.Pcl.Length -gt $antes.Pcl.Length

    Write-Host ("  PCL depois CONTEM a ChangeKey de antes : {0}" -f $(if ($contemChave) { "SIM" } else { "nao" }))
    Write-Host ("  PCL depois CONTEM a PCL de antes       : {0}" -f $(if ($contemPcl) { "SIM" } else { "nao" }))
    Write-Host ("  PCL cresceu                            : {0}" -f $(if ($cresceu) { "SIM" } else { "nao" }))
    return $contemChave
}

# ---------- alvo: SO o artefato meu ----------
$excluidos = $ns.GetDefaultFolder(3)
$raiz = $ns.GetDefaultFolder(6).Parent

$alvo = $null
$itens = $excluidos.Items
try {
    for ($i = 1; $i -le $itens.Count; $i++) {
        $it = $itens.Item($i)
        $ok = $false
        try { $ok = "$($it.Subject)".StartsWith("[IRIS-SPIKE-C]", [StringComparison]::Ordinal) } catch { }
        if ($ok) { $alvo = $it; break }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($it)
    }
} finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($itens) }

if ($null -eq $alvo) { Write-Host "ABORTADO: nao achei o artefato."; exit 1 }
Write-Host ("alvo: {0}" -f $alvo.Subject)
Write-Host ("pasta: {0}" -f $alvo.Parent.Name)
Write-Host ""

$temp = $raiz.Folders.Add("Iris Q2 causal")
Write-Host "pasta 'Iris Q2 causal' criada"

$copia = $null
try {
    $f0 = Foto $alvo

    # ---- ida ----
    $movido = $alvo.Move($temp)
    $f1 = Foto $movido
    $moveHerda = Analisar $f0 $f1 "MOVE: Itens Excluidos -> Iris Q2 causal"

    # ---- copy, controle negativo ----
    $copia = $movido.Copy()
    $f2 = Foto $copia
    $copyHerda = Analisar $f1 $f2 "COPY (controle negativo): a copia coexiste com o original"

    # ---- volta ----
    $devolvido = $movido.Move($excluidos)
    $f3 = Foto $devolvido
    $voltaHerda = Analisar $f1 $f3 "MOVE de volta: Iris Q2 causal -> Itens Excluidos"

    # POS-CONDICAO conferida, nao presumida.
    if ($f3.Pasta -ne $excluidos.Name) { throw "o item NAO voltou: esta em $($f3.Pasta)" }
    Write-Host ""
    Write-Host ("pos-condicao: item confirmado em '{0}'" -f $f3.Pasta)

    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($devolvido)
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($movido)

    Write-Host ""
    Write-Host ("=" * 74)
    Write-Host "VEREDITO"
    Write-Host ("=" * 74)
    Write-Host ("  Move  preserva ancestralidade na PCL : {0}" -f $(if ($moveHerda) { "SIM" } else { "NAO" }))
    Write-Host ("  volta preserva ancestralidade na PCL : {0}" -f $(if ($voltaHerda) { "SIM" } else { "NAO" }))
    Write-Host ("  Copy  preserva ancestralidade na PCL : {0}" -f $(if ($copyHerda) { "SIM" } else { "NAO" }))
    Write-Host ""
    if ($moveHerda -and $copyHerda) {
        Write-Host "  => Ha aresta causal, mas ela NAO separa Move de Copy."
        Write-Host "     Separar exige observar COEXISTENCIA: se o original"
        Write-Host "     sumiu e so um descendente apareceu, e Move; se os dois"
        Write-Host "     coexistem, e Copy."
    } elseif ($moveHerda) {
        Write-Host "  => Ha aresta causal e ela distingue Move de Copy."
    } else {
        Write-Host "  => A PCL nao carrega a ChangeKey anterior atraves do Move."
        Write-Host "     A secao 11.3 fica como esta."
    }
} finally {
    # A copia sai por MOVE, com pos-condicao conferida.
    if ($copia) {
        try {
            $movidaCopia = $copia.Move($excluidos)
            Write-Host ""
            Write-Host ("copia movida para '{0}' (confirmado)" -f $movidaCopia.Parent.Name)
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($movidaCopia)
        } catch {
            Write-Host "AVISO: nao consegui mover a copia: $($_.Exception.Message)"
        }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($copia)
    }

    # Reler a pasta DO ZERO antes de decidir apagar.
    $conferir = $null
    foreach ($f in $raiz.Folders) {
        if ($f.Name -eq "Iris Q2 causal") { $conferir = $f; break }
    }
    if ($conferir) {
        $dentro = $conferir.Items
        $n = $dentro.Count
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($dentro)
        Write-Host ("'Iris Q2 causal' contem {0} item(ns) na releitura" -f $n)
        if ($n -eq 0) {
            $conferir.Delete()
            Write-Host "pasta removida da raiz (soft: vai para Itens Excluidos)"
        } else {
            Write-Host "pasta MANTIDA: ainda ha item dentro."
        }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($conferir)
    }
}
