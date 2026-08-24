# Caca EXAUSTIVA: qual proptag devolve a contagem do servidor?
#
# A q8-caca-contagem.ps1 varreu tres faixas plausiveis (0x0E00, 0x3600,
# 0x6600) e nao achou. "Faixa plausivel" e um palpite disfarcado de metodo:
# se o valor estiver fora dela, o resultado negativo diz mais sobre onde eu
# procurei do que sobre onde ele esta.
#
# Aqui varro o espaco INTEIRO de identificadores, 0x0000-0xFFFF, em PT_LONG,
# na pasta menor (alvo 145). Os candidatos sao depois conferidos na Caixa de
# Entrada, onde precisam devolver 17728 - duas pastas, dois alvos, para que
# uma coincidencia nao passe por resposta.
$ErrorActionPreference = "Stop"
$ol = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
$ns = $ol.GetNamespace("MAPI")
$PT = "http://schemas.microsoft.com/mapi/proptag/"

$ALVO_PEQUENA = 145      # "1. Backup",        servidor
$ALVO_GRANDE  = 17728    # "Caixa de Entrada", servidor

function Pasta($nome) {
    $achada = $null
    function Ir($p, $prof) {
        if ($script:achada -or $prof -gt 6) { return }
        if ($p.Name -eq $nome) { $script:achada = $p; return }
        $fs = $null
        try {
            $fs = $p.Folders; $c = $fs.Count
            for ($i = 1; $i -le $c; $i++) {
                $f = $fs.Item($i)
                Ir $f ($prof + 1)
                if (-not $script:achada) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($f) }
            }
        } catch { } finally { if ($fs) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($fs) } }
    }
    foreach ($st in $ns.Stores) {
        $raiz = $st.GetRootFolder()
        Ir $raiz 0
        if ($script:achada) { break }
    }
    return $script:achada
}

$pequena = Pasta "1. Backup"
if (-not $pequena) { Write-Host "nao achei '1. Backup'"; exit 1 }

Write-Host "varrendo 0x0000-0xFFFF em PT_LONG na '1. Backup' (alvo $ALVO_PEQUENA)..."
$pa = $pequena.PropertyAccessor
$candidatos = @()
$legiveis = 0
$sw = [Diagnostics.Stopwatch]::StartNew()
for ($id = 0; $id -le 0xFFFF; $id++) {
    if (($id -band 0x0FFF) -eq 0) {
        Write-Host ("  ... 0x{0:X4}  ({1:N0} legiveis, {2:N0}s)" -f $id, $legiveis, $sw.Elapsed.TotalSeconds)
    }
    $tag = "{0}0x{1:X4}0003" -f $PT, $id
    try {
        $v = $pa.GetProperty($tag)
        if ($v -is [int] -or $v -is [long]) {
            $legiveis++
            if ([int]$v -eq $ALVO_PEQUENA) { $candidatos += ("0x{0:X4}" -f $id) }
        }
    } catch { }
}
[void][Runtime.InteropServices.Marshal]::ReleaseComObject($pa)
$sw.Stop()

Write-Host ""
Write-Host ("{0:N0} propriedades PT_LONG legiveis em {1:N0}s" -f $legiveis, $sw.Elapsed.TotalSeconds)
Write-Host ("candidatos (= {0}): {1}" -f $ALVO_PEQUENA, $(if ($candidatos) { $candidatos -join ", " } else { "NENHUM" }))

if (-not $candidatos) {
    Write-Host ""
    Write-Host "RESULTADO NEGATIVO: nenhuma propriedade PT_LONG da pasta local"
    Write-Host "devolve a contagem do servidor. O numero que o Outlook exibe na aba"
    Write-Host "Sincronizacao nao esta acessivel por PropertyAccessor sobre o Folder."
    exit 0
}

Write-Host ""
Write-Host "=== conferindo candidatos na Caixa de Entrada (alvo $ALVO_GRANDE) ==="
$script:achada = $null
$grande = Pasta "Caixa de Entrada"
$pa2 = $grande.PropertyAccessor
foreach ($c in $candidatos) {
    $tag = "{0}{1}0003" -f $PT, $c
    try {
        $v = $pa2.GetProperty($tag)
        $marca = if ([int]$v -eq $ALVO_GRANDE) { "  <-- RESPOSTA" } else { "" }
        Write-Host ("  {0} = {1}{2}" -f $c, $v, $marca)
    } catch { Write-Host ("  {0} = <nao legivel aqui>" -f $c) }
}
[void][Runtime.InteropServices.Marshal]::ReleaseComObject($pa2)
