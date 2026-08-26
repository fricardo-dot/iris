<#
.SYNOPSIS
    Mostra quem pode escrever no arquivo de ativacao, e o comando para tirar
    quem nao devia estar la. NAO altera nada.

.DESCRIPTION
    O Iris recusa a ativacao se alguem alem de voce, do SYSTEM e dos
    Administradores puder escrever nela. Quem pode escrever pode trocar a
    autorizacao, e trocar a autorizacao e escolher para onde o seu e-mail vai.

    Esta conferencia existe porque a suposicao de que "%LOCALAPPDATA% ja e
    protegido por usuario" se mostrou FALSA nesta maquina: o arquivo tinha
    controle total herdado por um SID de outra maquina e por um SID de
    capability.

    O roteiro nao conserta sozinho. Mexer em permissao e alterar configuracao
    de seguranca do sistema, e essa decisao e sua.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File "C:\Users\Ricardo\Documents\Iris\tools\conferir-permissao.ps1"
#>
[CmdletBinding()]
param(
    [string] $Caminho = (Join-Path $env:LOCALAPPDATA "Iris\ativacao.json")
)

$ErrorActionPreference = 'Stop'

$eu0 = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$elevado = (New-Object System.Security.Principal.WindowsPrincipal $eu0).IsInRole(
    [System.Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not (Test-Path -LiteralPath $Caminho)) {
    Write-Host ""
    Write-Host "Nao achei ativacao." -ForegroundColor Yellow
    Write-Host ""

    # OS CAMINHOS ENTRE COLCHETES, de proposito: espaco no fim de uma variavel
    # de ambiente e invisivel, e produz um caminho que IMPRIME igual e NAO
    # existe. Foi a hipotese que sobrou depois de conta, perfil e elevacao
    # terem sido descartados.
    Write-Host ("  procurei em ... [{0}]" -f $Caminho)
    Write-Host ("  LOCALAPPDATA .. [{0}]  ({1} caracteres)" -f $env:LOCALAPPDATA, $env:LOCALAPPDATA.Length)
    Write-Host ("  quem .......... {0}" -f $eu0.Name)
    Write-Host ("  elevado ....... {0}" -f $elevado)
    Write-Host ""

    # E PROCURA DE VERDADE, em vez de so reclamar.
    $tentativas = @(
        (Join-Path $env:LOCALAPPDATA "Iris\ativacao.json")
        (Join-Path $env:USERPROFILE "AppData\Local\Iris\ativacao.json")
        ("C:\Users\" + $env:USERNAME + "\AppData\Local\Iris\ativacao.json")
    ) | Select-Object -Unique

    $achou = $null
    foreach ($t in $tentativas) {
        $existe = Test-Path -LiteralPath $t
        $marca = if ($existe) { "ACHEI" } else { "nao  " }
        Write-Host ("  {0}  [{1}]" -f $marca, $t)
        if ($existe -and -not $achou) { $achou = $t }
    }
    Write-Host ""

    if ($achou) {
        Write-Host "O arquivo esta la. Rode de novo apontando para ele:" -ForegroundColor Green
        Write-Host ""
        Write-Host ("  ... conferir-permissao.ps1 -Caminho " + [char]34 + $achou + [char]34) -ForegroundColor Green
    } else {
        Write-Host "Nenhum dos caminhos acima tem o arquivo." -ForegroundColor Yellow
        Write-Host "Se souber onde ele esta, passe -Caminho." -ForegroundColor Yellow
    }
    Write-Host ""
    exit 1
}

$eu = [System.Security.Principal.WindowsIdentity]::GetCurrent().User
$acl = Get-Acl $Caminho

$permitidos = @(
    $eu.Value
    (New-Object System.Security.Principal.SecurityIdentifier ([System.Security.Principal.WellKnownSidType]::LocalSystemSid), $null).Value
    (New-Object System.Security.Principal.SecurityIdentifier ([System.Security.Principal.WellKnownSidType]::BuiltinAdministratorsSid), $null).Value
)

# Os mesmos direitos que o Iris considera escrita.
$perigosos = [System.Security.AccessControl.FileSystemRights]::WriteData -bor
             [System.Security.AccessControl.FileSystemRights]::AppendData -bor
             [System.Security.AccessControl.FileSystemRights]::Delete -bor
             [System.Security.AccessControl.FileSystemRights]::ChangePermissions -bor
             [System.Security.AccessControl.FileSystemRights]::TakeOwnership

Write-Host ""
Write-Host "Arquivo: $Caminho"

# O DONO ENTRA NA DECISAO, e comparado por SID.
#
# A primeira versao avisava sobre dono errado e depois decidia so pelas ACEs
# excedentes: com dono errado e nenhuma ACE sobrando, ela imprimia "O Iris
# aceita" e saia com 0 -- dizendo o contrario do que o Iris faria. E comparar
# por NOME quebra com conta renomeada, dominio, ou idioma do Windows.
$donoSid = try {
    $acl.GetOwner([System.Security.Principal.SecurityIdentifier])
} catch { $null }

$donoOk = ($donoSid -ne $null) -and ($donoSid.Value -eq $eu.Value)
Write-Host ("Dono:    {0}" -f $acl.Owner)
if (-not $donoOk) {
    Write-Host "  SOBRA: o dono nao e voce. O Iris vai RECUSAR." -ForegroundColor Red
}
Write-Host ""

$intrusos = @()
foreach ($regra in $acl.Access) {
    if ($regra.AccessControlType -ne 'Allow') { continue }
    if (($regra.FileSystemRights -band $perigosos) -eq 0) { continue }

    $sid = try {
        $regra.IdentityReference.Translate([System.Security.Principal.SecurityIdentifier]).Value
    } catch { $regra.IdentityReference.Value }

    $nome = try {
        ([System.Security.Principal.SecurityIdentifier]$sid).Translate(
            [System.Security.Principal.NTAccount]).Value
    } catch { "(nao resolve nesta maquina)" }

    if ($permitidos -contains $sid) {
        Write-Host ("  OK      {0}  {1}" -f $nome, $sid) -ForegroundColor Green
    } else {
        Write-Host ("  SOBRA   {0}  {1}" -f $nome, $sid) -ForegroundColor Red
        $intrusos += $sid
    }
}

Write-Host ""
if ($intrusos.Count -eq 0 -and $donoOk) {
    Write-Host "Ninguem a mais pode escrever, e o dono e voce. O Iris aceita." -ForegroundColor Green
    exit 0
}

if ($intrusos.Count -gt 0) {
    Write-Host "$($intrusos.Count) identidade(s) a mais podem escrever. O Iris vai RECUSAR." -ForegroundColor Red
}
if (-not $donoOk) {
    Write-Host "O dono do arquivo nao e voce. O Iris vai RECUSAR." -ForegroundColor Red
    Write-Host "Para retomar a posse:" -ForegroundColor Cyan
    Write-Host ("  takeown /f " + [char]34 + $Caminho + [char]34) -ForegroundColor Green
    Write-Host ""
}
Write-Host ""
Write-Host "Para consertar, rode:" -ForegroundColor Cyan
Write-Host ""
$aspas = [char]34
# UM COMANDO SO, E A ORDEM IMPORTA.
#
# A primeira versao mandava /inheritance:r primeiro e as concessoes depois,
# em chamadas separadas. Entre uma e outra o arquivo fica SEM ACE nenhuma:
# quem parasse no meio -- ou tivesse o comando falhando -- perdia o acesso
# ao proprio arquivo. Concede primeiro, corta a heranca depois, numa
# invocacao so.
#
# (M) e nao (R,W) para voce: sem Delete nao da para substituir a ativacao
# depois, e o erro apareceria como uma falha de permissao sem explicacao.
$cmd = "icacls " + $aspas + $Caminho + $aspas +
       " /grant:r " + $aspas + "*" + $eu.Value + ":(M)" + $aspas +
       " " + $aspas + "*S-1-5-18:(F)" + $aspas +
       " " + $aspas + "*S-1-5-32-544:(F)" + $aspas +
       " /inheritance:r"
Write-Host $cmd -ForegroundColor Green
Write-Host ""
Write-Host "Concede primeiro e corta a heranca depois, numa chamada so: entre" -ForegroundColor Yellow
Write-Host "os dois passos o arquivo ficaria sem permissao nenhuma." -ForegroundColor Yellow
Write-Host ""
Write-Host "Depois, confira de novo:" -ForegroundColor Cyan
Write-Host "  dotnet run --project tools\Iris.CrashHarness -- ativacao"
Write-Host ""
exit 1
