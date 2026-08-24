# Remove os artefatos que EU deixei na caixa, e nada mais.
#
# ESCREVE, e desta vez de forma PERMANENTE: as pastas estao dentro de
# Itens Excluidos, e Delete() ali nao vai para lugar nenhum. Autorizado
# pelo usuario.
#
# ------------------------------------------------------------------
# O DESENHO E TODO SOBRE NAO APAGAR COISA DO USUARIO
#
# A pasta Itens Excluidos contem mensagens EXCLUIDAS PELO USUARIO junto
# com os meus artefatos. Um erro aqui e irreversivel.
#
# Por isso:
#   1. So pastas cujo nome comeca com "Iris " E que estejam sob Itens
#      Excluidos. Nome de pasta sozinho nao basta.
#   2. TODO item dentro da pasta tem de casar com um marcador meu. Um
#      unico item desconhecido e a pasta INTEIRA e preservada e reportada.
#   3. Itens soltos em Itens Excluidos so saem se casarem com marcador.
#   4. Passo 1 e INVENTARIO. Nada e apagado antes de a lista sair.
#
# Assunto vazio conta como DESCONHECIDO. Um item sem assunto pode ser do
# usuario, e "provavelmente e meu" nao e criterio para exclusao permanente.

$ErrorActionPreference = "Stop"
$ol = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
$ns = $ol.GetNamespace("MAPI")
function Solta($o) { if ($null -ne $o) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } }

# Marcadores dos itens que este projeto criou, ao longo das fases.
$marcadores = @(
    "[IRIS-SPIKE",
    "[IRIS-TESTE",
    "Iris 1.5 - verificacao",
    "Iris 1.5 - verificação",
    "Q4-item",
    "Q4-NOVO",
    "IRISQ4-",
    "IRISQ3-",
    "IRISQ4S-"
)

function EhMeu([string]$assunto) {
    if ([string]::IsNullOrWhiteSpace($assunto)) { return $false }
    foreach ($m in $marcadores) {
        if ($assunto.StartsWith($m, [StringComparison]::OrdinalIgnoreCase)) { return $true }
    }
    return $false
}

$excluidos = $ns.GetDefaultFolder(3)

# ==================================================================
# PASSO 1 - INVENTARIO. Nada e apagado aqui.
# ==================================================================
Write-Host ("=" * 76)
Write-Host "PASSO 1 - INVENTARIO (nada e apagado)"
Write-Host ("=" * 76)

$aRemover = New-Object 'System.Collections.Generic.List[object]'
$preservadas = New-Object 'System.Collections.Generic.List[object]'

function Examinar($pasta, [string]$caminho) {
    $desconhecidos = New-Object 'System.Collections.Generic.List[string]'
    $meus = 0
    $itens = $pasta.Items
    try {
        for ($i = 1; $i -le $itens.Count; $i++) {
            $it = $itens.Item($i)
            try {
                $a = ""
                try { $a = "$($it.Subject)" } catch { }
                if (EhMeu $a) { $meus++ }
                else { $desconhecidos.Add($(if ($a -eq "") { "(sem assunto)" } else { $a })) }
            } finally { Solta $it }
        }
    } finally { Solta $itens }

    # subpastas contam junto: pasta so sai se a arvore inteira for minha
    $filhas = $pasta.Folders
    try {
        for ($k = 1; $k -le $filhas.Count; $k++) {
            $f = $filhas.Item($k)
            try {
                $r = Examinar $f "$caminho/$($f.Name)"
                $meus += $r.Meus
                foreach ($d in $r.Desconhecidos) { $desconhecidos.Add($d) }
            } finally { Solta $f }
        }
    } finally { Solta $filhas }

    return @{ Meus = $meus; Desconhecidos = $desconhecidos }
}

foreach ($f in @($excluidos.Folders)) {
    try {
        if (-not $f.Name.StartsWith("Iris ", [StringComparison]::Ordinal)) {
            Solta $f; continue
        }
        $r = Examinar $f $f.Name
        if ($r.Desconhecidos.Count -eq 0) {
            $aRemover.Add(@{ Nome = $f.Name; Meus = $r.Meus; Id = $f.EntryID })
            Write-Host ("  REMOVER   {0,-34} {1,4} itens, todos meus" -f $f.Name, $r.Meus)
        } else {
            $preservadas.Add(@{ Nome = $f.Name; Desconhecidos = $r.Desconhecidos })
            Write-Host ("  PRESERVAR {0,-34} {1} item(ns) DESCONHECIDO(s)" -f $f.Name, $r.Desconhecidos.Count)
            foreach ($d in ($r.Desconhecidos | Select-Object -First 3)) {
                Write-Host ("              -> {0}" -f $d.Substring(0, [Math]::Min(52, $d.Length)))
            }
        }
    } finally { Solta $f }
}

# itens SOLTOS em Itens Excluidos
$soltosMeus = New-Object 'System.Collections.Generic.List[string]'
$itens = $excluidos.Items
try {
    for ($i = 1; $i -le $itens.Count; $i++) {
        $it = $itens.Item($i)
        try {
            $a = ""
            try { $a = "$($it.Subject)" } catch { }
            if (EhMeu $a) { $soltosMeus.Add($it.EntryID) }
        } finally { Solta $it }
    }
} finally { Solta $itens }

Write-Host ""
Write-Host ("  itens SOLTOS em Itens Excluidos com marcador meu: {0}" -f $soltosMeus.Count)
Write-Host ("  (Itens Excluidos tem {0} itens diretos no total)" -f $excluidos.Items.Count)

Write-Host ""
Write-Host ("=" * 76)
Write-Host ("A REMOVER: {0} pasta(s) e {1} item(ns) solto(s)" -f $aRemover.Count, $soltosMeus.Count)
if ($preservadas.Count -gt 0) {
    Write-Host ("PRESERVADAS por conterem item desconhecido: {0}" -f $preservadas.Count)
}
Write-Host ("=" * 76)

# ==================================================================
# PASSO 2 - REMOVER. Permanente.
# ==================================================================
Write-Host ""
Write-Host "PASSO 2 - removendo (PERMANENTE)"

$okPastas = 0; $erroPastas = 0
foreach ($p in $aRemover) {
    try {
        $f = $ns.GetFolderFromID($p.Id)
        try { $f.Delete(); $okPastas++ } finally { Solta $f }
    } catch {
        $erroPastas++
        Write-Host ("  falhou {0}: {1}" -f $p.Nome, $_.Exception.Message)
    }
}
Write-Host ("  pastas removidas: {0}   falhas: {1}" -f $okPastas, $erroPastas)

$okItens = 0
foreach ($id in $soltosMeus) {
    try {
        $it = $ns.GetItemFromID($id)
        try { $it.Delete(); $okItens++ } finally { Solta $it }
    } catch { }
}
Write-Host ("  itens soltos removidos: {0}" -f $okItens)

Solta $excluidos

Write-Host ""
Write-Host ("=" * 76)
Write-Host "ESTADO FINAL"
Write-Host ("=" * 76)
foreach ($i in @(6,5,16,3,23)) {
    $f = $ns.GetDefaultFolder($i)
    Write-Host ("  {0,-20} {1,6} itens" -f $f.Name, $f.Items.Count)
    Solta $f
}
$exc = $ns.GetDefaultFolder(3)
$sobrou = 0
foreach ($f in $exc.Folders) { if ($f.Name -like "Iris*") { $sobrou++; Write-Host ("  ainda em Excluidos: {0}" -f $f.Name) }; Solta $f }
if ($sobrou -eq 0) { Write-Host "  nenhuma pasta Iris* em Itens Excluidos" }
Solta $exc
$raiz = $ns.GetDefaultFolder(6).Parent
$sobrou = 0
foreach ($f in $raiz.Folders) { if ($f.Name -like "Iris*") { $sobrou++; Write-Host ("  ainda na raiz: {0}" -f $f.Name) }; Solta $f }
if ($sobrou -eq 0) { Write-Host "  nenhuma pasta Iris* na raiz" }
Solta $raiz
