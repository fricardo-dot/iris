<#
.SYNOPSIS
    Gera o par de chaves que assina as versões do Iris.

.DESCRIPTION
    RODE ISTO UMA VEZ SÓ, E RODE VOCÊ.

    A chave PRIVADA é o que prova que uma versão é sua. Quem a tiver publica
    atualizações que o Iris aceita como legítimas — dentro de um programa que lê
    o seu e-mail. Ela não entra no repositório, não entra em backup na nuvem sem
    senha, e não passa por mais ninguém.

    A chave PÚBLICA é o oposto: ela existe para ser distribuída. Vai embutida no
    executável, e é com ela que cada cópia do Iris confere a assinatura.

    ECDSA na curva P-256, com SHA-256. Não é a escolha mais moderna que existe —
    é a que o .NET traz pronta, sem dependência externa e sem código de
    criptografia escrito à mão. Para assinar um manifesto, isso basta e sobra.

.PARAMETER Destino
    Onde guardar a chave privada. O padrão fica fora do repositório de propósito.

.EXAMPLE
    .\tools\gerar-chave-de-assinatura.ps1
#>
[CmdletBinding()]
param(
    [string] $Destino = (Join-Path $env:USERPROFILE '.iris\chave-de-assinatura.pem')
)

$ErrorActionPreference = 'Stop'

# O REPOSITORIO E O ULTIMO LUGAR ONDE ELA PODE CAIR.
$raiz = Split-Path -Parent $PSScriptRoot
if ($Destino -like (Join-Path $raiz '*')) {
    throw "RECUSADO: $Destino esta dentro do repositorio. A chave privada nao entra no git."
}

if (Test-Path $Destino) {
    Write-Host ''
    Write-Host "JA EXISTE uma chave em $Destino" -ForegroundColor Yellow
    Write-Host ''
    Write-Host 'Gerar outra INVALIDA todas as versoes ja publicadas: as copias do'
    Write-Host 'Iris que estao por ai tem a chave publica antiga embutida, e vao'
    Write-Host 'recusar tudo o que for assinado com a nova.'
    Write-Host ''
    Write-Host 'Se e isso mesmo que voce quer, apague o arquivo primeiro.'
    exit 1
}

$pasta = Split-Path -Parent $Destino
if (-not (Test-Path $pasta)) { New-Item -ItemType Directory -Force $pasta | Out-Null }

$ecdsa = [System.Security.Cryptography.ECDsa]::Create(
    [System.Security.Cryptography.ECCurve+NamedCurves]::nistP256)

try {
    # A privada em PEM, e o arquivo nasce so seu.
    $privada = $ecdsa.ExportPkcs8PrivateKeyPem()
    Set-Content -Path $Destino -Value $privada -Encoding ascii

    # SO O DONO LE. Sem isto, o arquivo herda as permissoes da pasta -- que num
    # perfil comum ja e restrito, e num redirecionado pode nao ser.
    $acl = Get-Acl $Destino
    $acl.SetAccessRuleProtection($true, $false)
    $acl.SetAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule(
        $env:USERNAME, 'FullControl', 'Allow')))
    Set-Acl -Path $Destino -AclObject $acl

    $publica = [Convert]::ToBase64String($ecdsa.ExportSubjectPublicKeyInfo())

    Write-Host ''
    Write-Host 'Chave privada gravada em:' -ForegroundColor Green
    Write-Host "  $Destino"
    Write-Host ''
    Write-Host 'GUARDE UMA COPIA FORA DESTA MAQUINA. Perder esta chave nao quebra'
    Write-Host 'o Iris que ja esta instalado, mas voce nunca mais consegue publicar'
    Write-Host 'uma atualizacao que ele aceite -- teria de distribuir uma versao'
    Write-Host 'nova a mao, com outra chave publica embutida.'
    Write-Host ''
    Write-Host '--- A CHAVE PUBLICA (esta pode ser mostrada a qualquer um) ---' -ForegroundColor Cyan
    Write-Host ''
    Write-Host $publica
    Write-Host ''
    Write-Host 'Ela vai embutida no executavel: cole-a em'
    Write-Host '  src/Iris.App/ChaveDeAtualizacao.vb'
}
finally {
    $ecdsa.Dispose()
}
