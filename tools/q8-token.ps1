# Q8 - o token da janela e ESTAVEL?
#
# A impressao digital do ambiente exige duas propriedades: ser ESTAVEL
# enquanto o ambiente nao muda, e SENSIVEL quando ele muda. Eu tinha escrito
# no codigo que nao precisava verificar nenhuma das duas porque nao precisava
# DECODIFICAR o blob. Sao coisas diferentes: nao decodificar dispensa saber o
# significado, nao dispensa saber se o valor serve de impressao digital.
#
# Este script mede a metade GRATIS - estabilidade. A sensibilidade exige mover
# o cursor da janela e esta em tools/q8-sensibilidade.md, para o usuario.
$ErrorActionPreference = "Stop"

function LerTokens {
    $r = @{}
    $base = "HKCU:\Software\Microsoft\Office"
    foreach ($v in (Get-ChildItem $base -EA SilentlyContinue |
                    Where-Object { $_.PSChildName -match '^\d+\.\d+$' })) {
        $perfis = Join-Path $v.PSPath "Outlook\Profiles"
        if (-not (Test-Path $perfis)) { continue }
        Get-ChildItem $perfis -Recurse -EA SilentlyContinue | ForEach-Object {
            $props = Get-ItemProperty $_.PSPath -EA SilentlyContinue
            if ($props.PSObject.Properties.Name -contains "00036601") {
                $b = $props."00036601"
                if ($b -is [byte[]]) { $b = ($b | ForEach-Object { $_.ToString("X2") }) -join "-" }
                $r[$_.PSPath.ToString()] = $b
            }
        }
    }
    return $r
}

Write-Host "=== leituras repetidas ==="
$primeira = $null
for ($i = 1; $i -le 5; $i++) {
    $t = LerTokens
    $assinatura = ($t.GetEnumerator() | Sort-Object Name |
                   ForEach-Object { "$($_.Value)" }) -join " / "
    Write-Host ("  leitura {0}: {1}" -f $i, $assinatura)
    if ($null -eq $primeira) { $primeira = $assinatura }
    elseif ($primeira -ne $assinatura) { Write-Host "  *** MUDOU sem nada mudar ***" }
    Start-Sleep -Milliseconds 300
}

Write-Host ""
Write-Host "=== quem e quem ==="
foreach ($kv in (LerTokens).GetEnumerator() | Sort-Object Name) {
    $chave = $kv.Name.Substring($kv.Name.IndexOf("Profiles") + 9)
    Write-Host ("  {0}" -f $chave)
    Write-Host ("      00036601 = {0}" -f $kv.Value)
}

Write-Host ""
Write-Host "=== stores vistos pelo OOM ==="
try {
    $ol = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
    foreach ($st in $ol.GetNamespace("MAPI").Stores) {
        Write-Host ("  '{0}' cached={1} tipo={2}" -f $st.DisplayName, $st.IsCachedExchange, $st.ExchangeStoreType)
    }
} catch { Write-Host "  Outlook nao esta rodando" }
