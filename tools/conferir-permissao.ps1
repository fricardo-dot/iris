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
$pasta = Split-Path -Parent $Caminho

$permitidos = @(
    $eu.Value
    (New-Object System.Security.Principal.SecurityIdentifier ([System.Security.Principal.WellKnownSidType]::LocalSystemSid), $null).Value
    (New-Object System.Security.Principal.SecurityIdentifier ([System.Security.Principal.WellKnownSidType]::BuiltinAdministratorsSid), $null).Value
)

# Os mesmos direitos que o Iris considera perigosos, e sao DOIS conjuntos.
$doArquivo = [System.Security.AccessControl.FileSystemRights]::WriteData -bor
             [System.Security.AccessControl.FileSystemRights]::AppendData -bor
             [System.Security.AccessControl.FileSystemRights]::Delete -bor
             [System.Security.AccessControl.FileSystemRights]::ChangePermissions -bor
             [System.Security.AccessControl.FileSystemRights]::TakeOwnership

# NA PASTA o perigo e outro: quem pode CRIAR e APAGAR ali dentro troca o
# ativacao.json sem ter direito nenhum sobre o arquivo.
$daPasta = [System.Security.AccessControl.FileSystemRights]::CreateFiles -bor
           [System.Security.AccessControl.FileSystemRights]::CreateDirectories -bor
           [System.Security.AccessControl.FileSystemRights]::DeleteSubdirectoriesAndFiles -bor
           [System.Security.AccessControl.FileSystemRights]::Delete -bor
           [System.Security.AccessControl.FileSystemRights]::ChangePermissions -bor
           [System.Security.AccessControl.FileSystemRights]::TakeOwnership

function Conferir($alvo, $perigosos, $rotulo, $donoTemDeSerEu = $true) {
    $acl = Get-Acl $alvo
    $donoSid = try { $acl.GetOwner([System.Security.Principal.SecurityIdentifier]) } catch { $null }
    $donoOk = if ($donoTemDeSerEu) {
        ($null -ne $donoSid) -and ($donoSid.Value -eq $eu.Value)
    } else {
        ($null -ne $donoSid) -and ($permitidos -contains $donoSid.Value)
    }

    Write-Host ""
    Write-Host ("{0}: {1}" -f $rotulo, $alvo)
    Write-Host ("  dono: {0}" -f $acl.Owner)
    if (-not $donoOk) {
        Write-Host "  SOBRA: o dono nao e voce. O Iris vai RECUSAR." -ForegroundColor Red
    }

    $sobrando = @()
    foreach ($regra in $acl.Access) {
        if ($regra.AccessControlType -ne 'Allow') { continue }
        # ACE so-de-heranca e molde para o que nascer dentro, e nao vale para
        # o proprio objeto. O Iris a ignora, e aqui tem de ignorar tambem.
        if (($regra.PropagationFlags -band [System.Security.AccessControl.PropagationFlags]::InheritOnly) -ne 0) { continue }
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
            $sobrando += $sid
        }
    }
    return [pscustomobject]@{ DonoOk = $donoOk; Sobrando = $sobrando }
}

$mae = Split-Path -Parent $pasta

$rArquivo = Conferir $Caminho $doArquivo "Arquivo"
$rPasta   = Conferir $pasta   $daPasta   "Pasta  "
# A MAE TAMBEM: quem cria e apaga nela renomeia a pasta inteira e poe outra
# no lugar, com outra ativacao dentro. Nela o dono pode ser o sistema.
$rMae     = Conferir $mae     $daPasta   "Mae    " $false

$intrusos = @($rArquivo.Sobrando) + @($rPasta.Sobrando) + @($rMae.Sobrando)
$donoOk = $rArquivo.DonoOk -and $rPasta.DonoOk -and $rMae.DonoOk

if ((New-Object System.IO.DirectoryInfo $pasta).Attributes -band `
    [System.IO.FileAttributes]::ReparsePoint) {
    Write-Host ""
    Write-Host "A pasta e um link (junction). O Iris vai RECUSAR." -ForegroundColor Red
    $intrusos += "junction"
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

# DOIS COMANDOS, E A PASTA VEM PRIMEIRO.
#
# A pasta primeiro porque /inheritance:r nela pode mexer no que os filhos
# herdam; com o arquivo ja tendo ACE explicita, ele sobrevive. Ao contrario,
# arrumar o arquivo e depois a pasta deixaria uma janela em que a pasta ainda
# permite substituicao.
#
# (OI)(CI) na pasta para o que nascer la dentro herdar a ACL certa -- senao o
# proximo ativacao.json volta a ter o problema.
#
# Em cada um: concede primeiro, corta a heranca depois, numa invocacao so.
# Comecar por /inheritance:r deixaria o objeto sem ACE nenhuma entre os dois
# passos.
$cmdPasta = "icacls " + $aspas + $pasta + $aspas +
            " /grant:r " + $aspas + "*" + $eu.Value + ":(OI)(CI)(M)" + $aspas +
            " " + $aspas + "*S-1-5-18:(OI)(CI)(F)" + $aspas +
            " " + $aspas + "*S-1-5-32-544:(OI)(CI)(F)" + $aspas +
            " /inheritance:r"

$cmdArquivo = "icacls " + $aspas + $Caminho + $aspas +
              " /grant:r " + $aspas + "*" + $eu.Value + ":(M)" + $aspas +
              " " + $aspas + "*S-1-5-18:(F)" + $aspas +
              " " + $aspas + "*S-1-5-32-544:(F)" + $aspas +
              " /inheritance:r"

Write-Host $cmdPasta -ForegroundColor Green
Write-Host ""
Write-Host $cmdArquivo -ForegroundColor Green
Write-Host ""
Write-Host "A pasta PRIMEIRO: quem pode criar e apagar la dentro troca o" -ForegroundColor Yellow
Write-Host "ativacao.json sem ter direito nenhum sobre o arquivo." -ForegroundColor Yellow
Write-Host ""

Write-Host "Depois, confira de novo:" -ForegroundColor Cyan
Write-Host "  dotnet run --project tools\Iris.CrashHarness -- ativacao"
Write-Host ""
exit 1
