# Q4, o que faltava: o que acontece com uma varredura quando o Outlook
# morre no meio dela?
#
# ESCREVE. So pasta criada aqui e itens com marcador GUID desta execucao.
# FECHA E REABRE O OUTLOOK. Autorizado pelo usuario.
#
# ------------------------------------------------------------------
# POR QUE PRECISA SER ASSIM
#
# A varredura por Table leva milissegundos; um reinicio leva segundos.
# Nao da para interceptar naturalmente. Entao leio em lotes PEQUENOS e
# derrubo o Outlook ENTRE DOIS LOTES, com o cursor aberto — que e
# exatamente o cenario real de uma varredura lenta dividindo a fila unica
# da STA com a UI.
#
# ------------------------------------------------------------------
# O QUE SE QUER SABER
#
# O desfecho PERIGOSO seria: GetArray devolve vazio, EndOfTable vira True,
# e a varredura declara conclusao normal com metade dos itens. Isso e o
# mesmo mal da §16.1, agora por outra causa — e o S6 nao pegaria, porque a
# contagem "depois" tambem estaria indisponivel.
#
# O desfecho BOM e uma excecao COM com HRESULT reconhecivel, que o
# OutlookFailurePolicy da Fase 1 ja saiba classificar como NotConnected.
#
# NUNCA Stop-Process -Force. Se o Quit nao resolver, o script reporta e
# para — e isso tambem e um achado: Outlook que nao fecha com varredura
# aberta e informacao util.

$ErrorActionPreference = "Stop"
$PT = "http://schemas.microsoft.com/mapi/proptag/"
$MARCA = "IRISQ4R-" + ([guid]::NewGuid().ToString("N").Substring(0, 8))
$LOTE = 10
$TOTAL = 100

function Solta($o) { if ($null -ne $o) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } }
function Processos { @(Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue) }

function Conectar {
    return [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
}

# ==================================================================
# PREPARAR
# ==================================================================
Write-Host "marcador: $MARCA"
$ol = Conectar
$ns = $ol.GetNamespace("MAPI")
$raiz = $ns.GetDefaultFolder(6).Parent
$pasta = $raiz.Folders.Add("Iris Q4R")

$rasc = $ns.GetDefaultFolder(16)
$li = $rasc.Items
for ($i = 1; $i -le $TOTAL; $i++) {
    $m = $li.Add("IPM.Note")
    $m.Subject = ("{0} {1:d3}" -f $MARCA, $i)
    $m.Save()
    $mv = $m.Move($pasta); Solta $mv; Solta $m
}
Solta $li; Solta $rasc
Write-Host ("pasta 'Iris Q4R' com {0} itens" -f $pasta.Items.Count)

# ==================================================================
# ABRIR O CURSOR E LER O PRIMEIRO LOTE
# ==================================================================
Write-Host ""
Write-Host ("=" * 70)
Write-Host "1. abrindo o cursor e lendo o primeiro lote"
Write-Host ("=" * 70)

$t = $pasta.GetTable()
$cols = $t.Columns
$cols.RemoveAll()
$c = $cols.Add("Subject"); Solta $c
Solta $cols
$t.Sort("Subject", $false)

$lidos = 0
$a = $t.GetArray($LOTE)
$lidos += ($a.GetUpperBound(0) - $a.GetLowerBound(0) + 1)
Write-Host ("   lote 1: {0} linhas   EndOfTable={1}" -f $lidos, $t.EndOfTable)

# ==================================================================
# MATAR O OUTLOOK COM O CURSOR ABERTO
# ==================================================================
Write-Host ""
Write-Host ("=" * 70)
Write-Host "2. chamando Quit() COM O CURSOR ABERTO"
Write-Host ("=" * 70)

$fechou = $false
try {
    $ol.Quit()
    Write-Host "   Quit() aceito"
} catch {
    Write-Host ("   Quit() lancou: {0}" -f $_.Exception.Message)
}

$esperou = 0
while ((Processos).Count -gt 0 -and $esperou -lt 60) {
    Start-Sleep -Seconds 3
    $esperou += 3
}
$fechou = ((Processos).Count -eq 0)
Write-Host ("   Outlook fechou? {0}  (em {1}s)" -f $(if ($fechou) { "SIM" } else { "NAO" }), $esperou)
if (-not $fechou) {
    Write-Host "   ACHADO: o Outlook NAO fecha enquanto ha uma Table aberta."
    Write-Host "   Nao vou forcar. Isso ja e resposta parcial da Q4."
}

# ==================================================================
# TENTAR CONTINUAR A VARREDURA
# ==================================================================
Write-Host ""
Write-Host ("=" * 70)
Write-Host "3. tentando ler o PROXIMO lote com o Outlook fora"
Write-Host ("=" * 70)

$desfecho = "?"
$hresult = 0
try {
    $eot = $t.EndOfTable
    Write-Host ("   EndOfTable respondeu: {0}" -f $eot)
    $a2 = $t.GetArray($LOTE)
    $n2 = if ($null -eq $a2) { 0 } else { $a2.GetUpperBound(0) - $a2.GetLowerBound(0) + 1 }
    $lidos += $n2
    Write-Host ("   GetArray devolveu {0} linhas" -f $n2)
    if ($n2 -eq 0) {
        $desfecho = "VAZIO SEM ERRO"
        Write-Host ""
        Write-Host "   *** O PIOR DESFECHO ***"
        Write-Host "   Devolveu vazio e nao lancou. Uma varredura ingenua"
        Write-Host ("   declararia conclusao normal com {0} de {1} itens." -f $lidos, $TOTAL)
    } else {
        $desfecho = "CONTINUOU LENDO"
        Write-Host "   Continuou lendo mesmo com o processo fora."
    }
} catch [Runtime.InteropServices.COMException] {
    $hresult = $_.Exception.HResult
    $desfecho = "COMException"
    Write-Host ("   COMException, HRESULT = 0x{0:X8}" -f $hresult)
    Write-Host ("   mensagem: {0}" -f $_.Exception.Message.Split([char]13)[0])
    Write-Host ""
    Write-Host "   Este e o desfecho BOM: falha declarada, com HRESULT que o"
    Write-Host "   OutlookFailurePolicy pode classificar."
} catch {
    $desfecho = "outra excecao"
    Write-Host ("   {0}: {1}" -f $_.Exception.GetType().Name, $_.Exception.Message.Split([char]13)[0])
}

Write-Host ""
Write-Host ("   itens lidos ao todo: {0} de {1}" -f $lidos, $TOTAL)

# ==================================================================
# REABRIR E LIMPAR
# ==================================================================
Write-Host ""
Write-Host ("=" * 70)
Write-Host "4. reabrindo o Outlook e limpando"
Write-Host ("=" * 70)

foreach ($o in @($t, $pasta, $raiz, $ns, $ol)) { try { Solta $o } catch { } }
[GC]::Collect(); [GC]::WaitForPendingFinalizers()

if ((Processos).Count -eq 0) {
    $exe = $null
    foreach ($p in @("C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE",
                     "C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE")) {
        if (Test-Path $p) { $exe = $p; break }
    }
    Start-Process -FilePath $exe | Out-Null
    $esperou = 0
    while ($esperou -lt 180) {
        Start-Sleep -Seconds 5
        $esperou += 5
        try { $ol = Conectar; $ns = $ol.GetNamespace("MAPI"); break } catch { }
    }
    Write-Host ("   Outlook de volta em {0}s" -f $esperou)
} else {
    $ol = Conectar; $ns = $ol.GetNamespace("MAPI")
    Write-Host "   Outlook nunca saiu"
}

$raiz = $ns.GetDefaultFolder(6).Parent
$n = 0
foreach ($f in @($raiz.Folders)) {
    try {
        if ($f.Name -eq "Iris Q4R") {
            Write-Host ("   removendo 'Iris Q4R' ({0} itens)" -f $f.Items.Count)
            $f.Delete(); $n++
        }
    } catch { } finally { Solta $f }
}
$rasc = $ns.GetDefaultFolder(16)
$li = $rasc.Items
$sobrou = 0
for ($i = $li.Count; $i -ge 1; $i--) {
    $it = $li.Item($i)
    try { if ("$($it.Subject)".StartsWith($MARCA, [StringComparison]::Ordinal)) { $it.Delete(); $sobrou++ } }
    finally { Solta $it }
}
Solta $li; Solta $rasc
Write-Host ("   Rascunhos: {0} com o marcador removidos" -f $sobrou)

Write-Host ""
Write-Host ("=" * 70)
Write-Host ("DESFECHO: {0}" -f $desfecho)
if ($hresult -ne 0) { Write-Host ("HRESULT : 0x{0:X8}" -f $hresult) }
Write-Host ("LIDOS   : {0} de {1}" -f $lidos, $TOTAL)
Write-Host ("=" * 70)
