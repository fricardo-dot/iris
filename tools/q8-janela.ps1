# Q8 - A janela de sincronizacao e legivel pelo Iris?
#
# A EnvironmentPolicy poe a janela na impressao digital do ambiente porque a
# 18.4 mediu que ela muda O QUE EXISTE. Se o Iris nao conseguir ler a janela,
# a politica degrada sempre e o desenho nao serve.
$ErrorActionPreference = "Stop"

Write-Host "=== 1. O que o OOM oferece ==="
$ol = $null
try { $ol = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application") }
catch { Write-Host "Outlook nao esta rodando: $($_.Exception.Message)"; }

if ($ol) {
    $ns = $ol.GetNamespace("MAPI")
    foreach ($st in $ns.Stores) {
        $nome = $st.DisplayName
        $cached = $null; $tipo = $null
        try { $cached = $st.IsCachedExchange } catch { $cached = "ERRO: $($_.Exception.Message)" }
        try { $tipo = $st.ExchangeStoreType } catch { $tipo = "ERRO" }
        Write-Host ("  store='{0}' IsCachedExchange={1} ExchangeStoreType={2}" -f $nome, $cached, $tipo)
        # A janela NAO e propriedade do Store no OOM. Confirmando:
        $temJanela = $st | Get-Member -Name "*Sync*","*Window*" -ErrorAction SilentlyContinue
        if ($temJanela) { Write-Host "    membros Sync/Window: $($temJanela.Name -join ', ')" }
        else { Write-Host "    membros Sync/Window: NENHUM exposto pelo OOM" }
    }
}

Write-Host ""
Write-Host "=== 2. O registro do perfil ==="
$base = "HKCU:\Software\Microsoft\Office"
$vers = Get-ChildItem $base -ErrorAction SilentlyContinue |
        Where-Object { $_.PSChildName -match '^\d+\.\d+$' }
foreach ($v in $vers) {
    $perfis = Join-Path $v.PSPath "Outlook\Profiles"
    if (-not (Test-Path $perfis)) { continue }
    Write-Host "  versao $($v.PSChildName)"
    foreach ($p in Get-ChildItem $perfis -ErrorAction SilentlyContinue) {
        Write-Host "    perfil: $($p.PSChildName)"
        $achou = $false
        Get-ChildItem $p.PSPath -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
            $props = Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue
            foreach ($n in @("00036601","SyncWindowSetting","SyncWindowSettingMonths","00036655")) {
                if ($props.PSObject.Properties.Name -contains $n) {
                    $val = $props.$n
                    if ($val -is [byte[]]) { $val = ($val | ForEach-Object { $_.ToString("X2") }) -join " " }
                    Write-Host "      $n = $val   [$($_.PSChildName)]"
                    $achou = $true
                }
            }
        }
        # -ErrorAction SilentlyContinue acima: chave sem permissao some sem
        # avisar, entao "nenhum" e sobre o que foi LIDO.
        if (-not $achou) { Write-Host "      nenhum valor de janela ENTRE AS CHAVES LIDAS" }
    }
}

Write-Host ""
Write-Host "=== 3. Politica de grupo ==="
foreach ($v in @("16.0","15.0")) {
    $k = "HKCU:\Software\Policies\Microsoft\office\$v\outlook\cached mode"
    if (Test-Path $k) {
        Get-ItemProperty $k | Format-List | Out-String | Write-Host
    } else { Write-Host "  $k : ausente" }
}
