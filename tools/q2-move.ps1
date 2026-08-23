# Q2, experimento decisivo: as chaves sobrevivem a um Move?
#
# ESTE SCRIPT ESCREVE NA CAIXA. Autorizado pelo usuario.
#
# O que ele faz, e so isso:
#   - cria UMA pasta temporaria "Iris Q2 (temp)" na raiz da caixa;
#   - move DOIS itens para ela e DE VOLTA para a pasta de origem;
#   - faz UM Copy como controle negativo, e apaga a copia;
#   - apaga a pasta temporaria.
#
# Os dois itens:
#   A) o [IRIS-SPIKE-C] recebido, artefato meu da Fase 0, em Itens
#      Excluidos. E mensagem entregue pelo servidor de verdade.
#   B) a mensagem MAIS ANTIGA do Lixo Eletronico. E a pasta onde um erro
#      custa menos, e o item volta para la no fim.
#
# NUNCA usa Delete permanente. Delete() em item vai para Itens Excluidos.
#
# POR QUE O CONTROLE NEGATIVO E OBRIGATORIO: uma chave que sobrevive ao
# Move mas tambem e duplicada numa copia NAO serve como identidade — duas
# manifestacoes coexistentes teriam a mesma. Sem o Copy, um "sobreviveu"
# pareceria resposta e nao seria.
#
# SEM RETRY. Mutacao que falha no meio nao pode ser repetida as cegas: o
# script para e diz onde o item esta.

$ErrorActionPreference = "Stop"
$PT = "http://schemas.microsoft.com/mapi/proptag/"

$chaves = [ordered]@{
    "RecordKey"  = "0x0FF90102"
    "SearchKey"  = "0x300B0102"
    "MessageID"  = "0x1035001E"
    "ChangeKey"  = "0x65E20102"
    "PredChange" = "0x65E30102"
}

$ol = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
$ns = $ol.GetNamespace("MAPI")

function Hex($v) {
    if ($null -eq $v) { return "(nulo)" }
    if ($v -is [byte[]]) {
        if ($v.Length -eq 0) { return "(vazio)" }
        return (($v | ForEach-Object { $_.ToString("x2") }) -join "")
    }
    if ($v -is [Array]) {
        return (@($v | ForEach-Object {
            if ($_ -is [byte[]]) { ($_ | ForEach-Object { $_.ToString("x2") }) -join "" } else { "$_" }
        }) -join " , ")
    }
    return "$v"
}

function Fotografar($item, [string]$rotulo) {
    $r = [ordered]@{ Rotulo = $rotulo; EntryID = $item.EntryID }
    $pa = $item.PropertyAccessor
    try {
        foreach ($k in $chaves.Keys) {
            try { $r[$k] = Hex $pa.GetProperty($PT + $chaves[$k]) }
            catch { $r[$k] = "(erro)" }
        }
    } finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pa) }
    return $r
}

function Comparar($antes, $depois, [string]$oQue) {
    Write-Host ""
    Write-Host ("--- {0} ---" -f $oQue)
    Write-Host ("{0,-11} | {1,-8} | {2}" -f "chave", "mudou?", "detalhe")
    Write-Host ("-" * 74)
    $campos = @("EntryID") + @($chaves.Keys)
    foreach ($k in $campos) {
        $a = "$($antes[$k])"; $b = "$($depois[$k])"
        $igual = ($a -eq $b)
        $det = if ($igual) {
            if ($a.Length -gt 46) { $a.Substring(0, 43) + "..." } else { $a }
        } else {
            $n = [Math]::Min($a.Length, $b.Length)
            $comum = 0
            while ($comum -lt $n -and $a[$comum] -eq $b[$comum]) { $comum++ }
            "prefixo comum de $comum de $($a.Length) chars"
        }
        Write-Host ("{0,-11} | {1,-8} | {2}" -f $k, $(if ($igual) { "IGUAL" } else { "MUDOU" }), $det)
    }
}

# ---------------------------------------------------------------
$raiz = $ns.GetDefaultFolder(6).Parent    # raiz da caixa
$excluidos = $ns.GetDefaultFolder(3)
$lixo = $ns.GetDefaultFolder(23)          # olFolderJunk

Write-Host "ORIGENS"
Write-Host ("  Excluidos    : {0} itens" -f $excluidos.Items.Count)
Write-Host ("  Lixo         : {0} itens" -f $lixo.Items.Count)
Write-Host ""

# --- achar os dois alvos ---
function AcharPorPrefixo($pasta, [string]$prefixo) {
    $itens = $pasta.Items
    try {
        for ($i = 1; $i -le $itens.Count; $i++) {
            $it = $itens.Item($i)
            $ok = $false
            try { $ok = "$($it.Subject)".StartsWith($prefixo, [StringComparison]::Ordinal) } catch { }
            # So o RECEBIDO: o enviado esta em Itens Enviados, mas se algum
            # outro artefato casar o prefixo eu quero o que tem cabecalho de
            # transporte, que e a marca do que passou pelo servidor.
            if ($ok) {
                $temHdr = $false
                try {
                    $pa = $it.PropertyAccessor
                    try { $temHdr = ("$($pa.GetProperty($PT + '0x007D001E'))".Length -gt 100) }
                    finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pa) }
                } catch { }
                if ($temHdr) { return $it }
            }
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($it)
        }
    } finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($itens) }
    return $null
}

function MaisAntigo($pasta) {
    $itens = $pasta.Items
    try {
        $itens.Sort("[ReceivedTime]", $false)
        if ($itens.Count -lt 1) { return $null }
        return $itens.Item(1)
    } finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($itens) }
}

$alvoA = AcharPorPrefixo $excluidos "[IRIS-SPIKE-C]"
$alvoB = MaisAntigo $lixo

if ($null -eq $alvoA) { Write-Host "ABORTADO: nao achei o [IRIS-SPIKE-C] recebido."; exit 1 }
if ($null -eq $alvoB) { Write-Host "ABORTADO: Lixo Eletronico vazio."; exit 1 }

$alvos = @(
    @{ Item = $alvoA; Origem = $excluidos; Nome = "A: [IRIS-SPIKE-C] recebido (meu, Fase 0)" },
    @{ Item = $alvoB; Origem = $lixo;      Nome = "B: mais antiga do Lixo Eletronico" }
)

Write-Host "ALVOS"
foreach ($a in $alvos) {
    $s = "$($a.Item.Subject)"
    if ($s.Length -gt 44) { $s = $s.Substring(0,44) + "..." }
    Write-Host ("  {0}" -f $a.Nome)
    Write-Host ("     assunto: {0}" -f $s)
    Write-Host ("     origem : {0}" -f $a.Origem.Name)
}
Write-Host ""

# --- pasta temporaria ---
$temp = $null
foreach ($f in $raiz.Folders) {
    if ($f.Name -eq "Iris Q2 (temp)") { $temp = $f; break }
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($f)
}
if ($null -eq $temp) {
    $temp = $raiz.Folders.Add("Iris Q2 (temp)")
    Write-Host "pasta temporaria CRIADA: Iris Q2 (temp)"
} else {
    Write-Host "pasta temporaria ja existia: Iris Q2 (temp)"
}
Write-Host ""

$erro = $null
try {
    foreach ($a in $alvos) {
        Write-Host ("=" * 74)
        Write-Host $a.Nome
        Write-Host ("=" * 74)

        $antes = Fotografar $a.Item "antes"

        # ---- MOVE de ida ----
        $movido = $a.Item.Move($temp)
        $depoisMove = Fotografar $movido "depois do Move"
        Comparar $antes $depoisMove "MOVE: $($a.Origem.Name) -> Iris Q2 (temp)"

        # ---- COPY, controle negativo ----
        $copia = $movido.Copy()
        $daCopia = Fotografar $copia "copia"
        Comparar $depoisMove $daCopia "COPY (controle negativo): a copia coexiste com o original"

        $copia.Delete()   # soft: vai para Itens Excluidos
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($copia)

        # ---- MOVE de volta ----
        $devolvido = $movido.Move($a.Origem)
        $depoisVolta = Fotografar $devolvido "de volta"
        Comparar $antes $depoisVolta "IDA E VOLTA: comparado com o estado ORIGINAL"

        Write-Host ""
        Write-Host ("  item de volta em: {0}" -f $a.Origem.Name)
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($devolvido)
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($movido)
        Write-Host ""
    }
} catch {
    $erro = $_
    Write-Host ""
    Write-Host "!!! FALHA NO MEIO DA MUTACAO !!!"
    Write-Host $_.Exception.Message
    Write-Host "Sem retry. Confira 'Iris Q2 (temp)' — pode haver item la."
}

# --- limpeza ---
Write-Host ("=" * 74)
$sobrou = $temp.Items.Count
Write-Host ("itens restantes em Iris Q2 (temp): {0}" -f $sobrou)
if ($sobrou -eq 0 -and $null -eq $erro) {
    $temp.Delete()
    Write-Host "pasta temporaria APAGADA (soft, vai para Itens Excluidos)."
} else {
    Write-Host "pasta temporaria MANTIDA para inspecao."
}
[void][Runtime.InteropServices.Marshal]::ReleaseComObject($temp)

Write-Host ""
Write-Host "A copia do controle negativo foi para Itens Excluidos (Delete soft)."
if ($erro) { exit 1 }
