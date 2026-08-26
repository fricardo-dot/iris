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

if (-not (Test-Path $Caminho)) {
    Write-Host "Nao existe ativacao em: $Caminho" -ForegroundColor Yellow
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
Write-Host "Dono:    $($acl.Owner)"
if ($acl.Owner -ne [System.Security.Principal.WindowsIdentity]::GetCurrent().Name) {
    Write-Host "  ATENCAO: o dono nao e voce. O Iris vai recusar." -ForegroundColor Red
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
if ($intrusos.Count -eq 0) {
    Write-Host "Ninguem a mais pode escrever. O Iris aceita." -ForegroundColor Green
    exit 0
}

Write-Host "$($intrusos.Count) identidade(s) a mais podem escrever. O Iris vai RECUSAR." -ForegroundColor Red
Write-Host ""
Write-Host "Para consertar, rode:" -ForegroundColor Cyan
Write-Host ""
$aspas = [char]34
$linha1 = "  icacls " + $aspas + $Caminho + $aspas + " /inheritance:r"
$linha2 = "  icacls " + $aspas + $Caminho + $aspas + " /grant:r " + $aspas + "*" + $eu.Value + ":(R,W)" + $aspas
$linha3 = "  icacls " + $aspas + $Caminho + $aspas + " /grant:r " + $aspas + "*S-1-5-18:(F)" + $aspas + " " + $aspas + "*S-1-5-32-544:(F)" + $aspas
Write-Host $linha1 -ForegroundColor Green
Write-Host $linha2 -ForegroundColor Green
Write-Host $linha3 -ForegroundColor Green
Write-Host ""
Write-Host "O primeiro corta a heranca - e de la que vem o que sobra. Os" -ForegroundColor Yellow
Write-Host "outros devolvem so voce, o SYSTEM e os Administradores." -ForegroundColor Yellow
Write-Host ""
Write-Host "Depois, confira de novo:" -ForegroundColor Cyan
Write-Host "  dotnet run --project tools\Iris.CrashHarness -- ativacao"
Write-Host ""
exit 1
