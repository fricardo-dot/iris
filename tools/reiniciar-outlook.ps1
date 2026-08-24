# Reinicia o Outlook com educacao, e confere que ele voltou.
#
# REGRA DO PROJETO, e ela nasceu de um erro meu: NUNCA Stop-Process
# -Force. Uma vez eu matei o Outlook do usuario a forca e ele ficou
# minutos num estado quebrado — janela fantasma de 322x18 px em coordenada
# negativa, sem registro no ROT. Quit() e paciencia.
#
# Se o Quit nao resolver, este script NAO escala para forca: ele reporta e
# para. Outlook que nao fecha pode estar com dialogo aberto esperando o
# usuario, e matar isso pode custar trabalho nao salvo dele.

$ErrorActionPreference = "Stop"

function Processos { @(Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue) }

$antes = Processos
if ($antes.Count -eq 0) {
    Write-Host "Outlook nao esta rodando."
} else {
    Write-Host ("Outlook rodando: PID {0}" -f ($antes.Id -join ", "))

    # Estado antes, para comparar depois.
    try {
        $ol = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
        $ns = $ol.GetNamespace("MAPI")
        $in = $ns.GetDefaultFolder(6)
        Write-Host ("ANTES: Caixa de Entrada com {0} itens" -f $in.Items.Count)
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($in)
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($ns)

        Write-Host "chamando Quit()..."
        $ol.Quit()
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($ol)
    } catch {
        Write-Host ("nao consegui falar com o Outlook por COM: {0}" -f $_.Exception.Message)
    }

    [GC]::Collect(); [GC]::WaitForPendingFinalizers()

    $limite = 120
    $t = 0
    while ((Processos).Count -gt 0 -and $t -lt $limite) {
        Start-Sleep -Seconds 3
        $t += 3
        if ($t % 15 -eq 0) { Write-Host ("   ainda fechando... {0}s" -f $t) }
    }

    if ((Processos).Count -gt 0) {
        Write-Host ""
        Write-Host ("O Outlook NAO fechou em {0}s e eu NAO vou forcar." -f $limite)
        Write-Host "Provavel causa: dialogo aberto esperando resposta sua, ou item"
        Write-Host "nao salvo. Feche-o na tela e rode este script de novo."
        exit 1
    }
    Write-Host ("fechou em {0}s" -f $t)
}

# ---- subir de novo ----
$exe = $null
foreach ($p in @(
    "C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE",
    "C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE",
    "C:\Program Files\Microsoft Office\Office16\OUTLOOK.EXE")) {
    if (Test-Path $p) { $exe = $p; break }
}
if (-not $exe) {
    try { $exe = (Get-Command outlook.exe -ErrorAction Stop).Source } catch { }
}
if (-not $exe) { Write-Host "nao achei o OUTLOOK.EXE"; exit 1 }

Write-Host ""
Write-Host ("iniciando: {0}" -f $exe)
Start-Process -FilePath $exe | Out-Null

# A Fase 1 mediu que esta caixa leva de 30 a 90 s para o Outlook ficar
# acessivel por COM. Espero ate 240, porque agora ha resync pela frente.
$limite = 240
$t = 0
$vivo = $false
while ($t -lt $limite) {
    Start-Sleep -Seconds 5
    $t += 5
    try {
        $ol = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
        $ns = $ol.GetNamespace("MAPI")
        $in = $ns.GetDefaultFolder(6)
        $n = $in.Items.Count
        Write-Host ("DEPOIS ({0}s): acessivel por COM, Caixa de Entrada com {1} itens" -f $t, $n)
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($in)
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($ns)
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($ol)
        $vivo = $true
        break
    } catch {
        if ($t % 30 -eq 0) { Write-Host ("   ainda subindo... {0}s" -f $t) }
    }
}

if (-not $vivo) {
    Write-Host ("Outlook nao ficou acessivel por COM em {0}s." -f $limite)
    Write-Host "Ele pode estar pedindo senha ou mostrando dialogo na tela."
    exit 1
}

$ost = "C:\Users\Ricardo\AppData\Local\Microsoft\Outlook\conta.do.dono@empresa.com.ost"
if (Test-Path $ost) {
    Write-Host ("OST agora: {0:N0} MB" -f ((Get-Item $ost).Length / 1MB))
}
Write-Host ""
Write-Host "A sincronizacao da janela nova comeca agora e leva o tempo que levar."
