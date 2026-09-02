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
    criptografia escrito à mão.

    A CRIPTOGRAFIA NÃO ACONTECE AQUI. Ela acontece em tools/Iris.Assinatura,
    que é .NET 10. Este script fez tudo em PowerShell na primeira versão e não
    rodava: `powershell.exe` é o Windows PowerShell 5.1, sobre .NET Framework,
    onde ExportPkcs8PrivateKeyPem e ExportSubjectPublicKeyInfo não existem. A
    descoberta teria sido na hora de gerar a chave.

    O que sobrou para o PowerShell é o que ele faz bem: criar o arquivo com a
    ACL certa ANTES de ele ter conteúdo, e conversar com o usuário.

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
$raiz = Split-Path -Parent $PSScriptRoot

# O REPOSITORIO E O ULTIMO LUGAR ONDE ELA PODE CAIR.
#
# A conferencia era textual e um caminho relativo passava por ela: ".\chave.pem"
# nao comeca com o caminho absoluto da raiz, e Set-Content gravava dentro do
# repositorio assim mesmo. Resolver o caminho primeiro e a diferenca entre uma
# barreira e a aparencia de uma.
$pasta = Split-Path -Parent $Destino
if (-not $pasta) { $pasta = (Get-Location).Path }
if (-not (Test-Path $pasta)) { New-Item -ItemType Directory -Force $pasta | Out-Null }
$absoluto = Join-Path ((Resolve-Path $pasta).Path) (Split-Path -Leaf $Destino)

$raizAbsoluta = (Resolve-Path $raiz).Path
if ($absoluto.StartsWith($raizAbsoluta, [StringComparison]::OrdinalIgnoreCase)) {
    throw "RECUSADO: $absoluto esta dentro do repositorio. A chave privada nao entra no git."
}

if ((Test-Path $absoluto) -and (Get-Item $absoluto).Length -gt 0) {
    Write-Host ''
    Write-Host "JA EXISTE uma chave em $absoluto" -ForegroundColor Yellow
    Write-Host ''
    Write-Host 'Gerar outra INVALIDA todas as versoes ja publicadas: as copias do'
    Write-Host 'Iris que estao por ai tem a chave publica antiga embutida, e vao'
    Write-Host 'recusar tudo o que for assinado com a nova.'
    Write-Host ''
    Write-Host 'Se voce so quer ver a chave publica de novo, use:'
    Write-Host "  .\tools\publicar-versao.ps1 -MostrarChavePublica -Chave `"$absoluto`""
    Write-Host ''
    Write-Host 'Se e outra chave que voce quer, apague o arquivo primeiro.'
    exit 1
}

# ---------------------------------------------------------- a ACL vem PRIMEIRO
#
# O arquivo nasce VAZIO e ja restrito; so depois o utilitario escreve nele.
# Gravar e restringir em seguida deixava uma janela com a chave no disco sob a
# ACL herdada da pasta -- e, se o Set-Acl falhasse, ela ficava assim.
if (-not (Test-Path $absoluto)) { New-Item -ItemType File $absoluto | Out-Null }

try {
    # O SID DA IDENTIDADE CORRENTE, e nao $env:USERNAME. Um nome sem dominio
    # pode nao resolver numa maquina ingressada, ou resolver para outra conta
    # quando ha uma local homonima.
    $eu = [System.Security.Principal.WindowsIdentity]::GetCurrent().User
    $acl = Get-Acl $absoluto
    $acl.SetAccessRuleProtection($true, $false)
    $acl.SetAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule(
        $eu, 'FullControl', 'Allow')))
    Set-Acl -Path $absoluto -AclObject $acl
}
catch {
    Remove-Item $absoluto -Force -ErrorAction SilentlyContinue
    throw "Nao consegui restringir a ACL de $absoluto. Nada foi gravado. ($_)"
}

# ------------------------------------------------------------------- a chave

$ferramenta = & (Join-Path $PSScriptRoot 'montar-assinador.ps1')
$publica = & $ferramenta gerar --destino $absoluto
if ($LASTEXITCODE -ne 0) {
    Remove-Item $absoluto -Force -ErrorAction SilentlyContinue
    throw "A geracao da chave falhou com $LASTEXITCODE. Nada ficou no disco."
}

Write-Host ''
Write-Host 'Chave privada gravada em:' -ForegroundColor Green
Write-Host "  $absoluto"
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
Write-Host 'junto com o endereco do manifesto. Ver LANCAR.md.'
