# Q8 - retrato do perfil inteiro, para DIFERENCIAR em vez de adivinhar.
#
# O q8-token.ps1 procura por quatro nomes conhecidos e achou 00036601. Isso
# nao prova que 00036601 SEJA a janela de sincronizacao - so que existe e tem
# nome parecido com o que a literatura cita. Depois de mover o cursor e o
# token nao mudar em quatro leituras, a hipotese mais provavel deixa de ser
# "a janela nao e legivel" e passa a ser "eu estou lendo a chave errada".
#
# Este script despeja TODOS os valores sob Profiles. Rodando antes e depois
# da mudanca, o diff mostra qual valor realmente carrega a configuracao.
param([string]$Saida = "retrato.txt")
$ErrorActionPreference = "Stop"

$linhas = New-Object System.Collections.Generic.List[string]
$base = "HKCU:\Software\Microsoft\Office"
foreach ($v in (Get-ChildItem $base -EA SilentlyContinue |
                Where-Object { $_.PSChildName -match '^\d+\.\d+$' })) {
    $perfis = Join-Path $v.PSPath "Outlook\Profiles"
    if (-not (Test-Path $perfis)) { continue }
    foreach ($k in (Get-ChildItem $perfis -Recurse -EA SilentlyContinue)) {
        $caminho = $k.PSPath.ToString()
        $caminho = $caminho.Substring($caminho.IndexOf("Profiles"))
        $props = Get-ItemProperty $k.PSPath -EA SilentlyContinue
        if (-not $props) { continue }
        foreach ($n in ($props.PSObject.Properties.Name | Sort-Object)) {
            if ($n -like "PS*") { continue }
            $val = $props.$n
            if ($val -is [byte[]]) {
                # HASH, nunca "<blob N bytes>". A primeira versao truncava
                # blobs longos para so o tamanho, e blob longo e exatamente
                # onde a configuracao tem mais chance de morar: uma mudanca
                # DENTRO dele ficava invisivel para o diff, e o retrato dizia
                # "nada mudou" sobre a parte que ele nao olhou.
                $hex = ($val | ForEach-Object { $_.ToString("X2") }) -join "-"
                if ($val.Length -gt 64) {
                    $sha = [Security.Cryptography.SHA256]::Create()
                    $h = ($sha.ComputeHash($val) | Select-Object -First 8 |
                          ForEach-Object { $_.ToString("X2") }) -join ""
                    $val = "<blob $($val.Length) bytes sha=$h>"
                } else { $val = $hex }
            }
            $linhas.Add("$caminho :: $n = $val")
        }
    }
}
$linhas | Sort-Object | Set-Content -Path $Saida -Encoding utf8
Write-Host ("$($linhas.Count) valores em $Saida")
