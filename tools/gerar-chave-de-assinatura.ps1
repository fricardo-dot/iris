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
$pastaAbsoluta = (Resolve-Path $pasta).Path
$absoluto = Join-Path $pastaAbsoluta (Split-Path -Leaf $Destino)

# DUAS BARREIRAS, porque a primeira e lexical e a lexical tem furo.
#
# (1) Comparacao de prefixo COM O SEPARADOR. Sem ele, "Iris2" era recusado por
#     comecar igual a "Iris" -- e um irmao de nome parecido nao tem nada a ver
#     com o repositorio.
$raizAbsoluta = (Resolve-Path $raiz).Path.TrimEnd('\')
if ($pastaAbsoluta.TrimEnd('\') -eq $raizAbsoluta -or
    $pastaAbsoluta.StartsWith($raizAbsoluta + '\', [StringComparison]::OrdinalIgnoreCase)) {
    throw "RECUSADO: $absoluto esta dentro do repositorio. A chave privada nao entra no git."
}

# (2) PERGUNTA AO GIT. Comparar texto nao enxerga junction, link simbolico,
#     nome curto 8.3 nem unidade mapeada -- todos deixam o mesmo diretorio
#     acessivel por outro caminho. O git resolve o caminho de verdade, e se ele
#     disser que aquela pasta pertence a uma arvore de trabalho, pertence.
#
#     E ELA FALHA FECHADA. A primeira versao tratava QUALQUER saida diferente de
#     zero como "nao e repositorio" -- e aquilo desligava a barreira sozinho
#     quando o git nao estava no PATH, quando GIT_CEILING_DIRECTORIES ou GIT_DIR
#     atrapalhavam a descoberta, ou quando o git recusava por safe.directory.
#     Uma barreira que some calada e pior que nenhuma: quem escreveu o script
#     acha que ela esta la.
#
#     Agora so ha um jeito de passar: o git RODAR e dizer, com essas palavras,
#     que aquilo nao e um repositorio.
#
#     --git-dir, e nao --show-toplevel: dentro da propria pasta .git o
#     --show-toplevel FALHA com "must be run in a work tree", e aquilo tambem
#     era lido como "nao e repositorio". Uma junction apontando para .git
#     deixava a chave cair dentro dos metadados do repositorio.
#
#     E O GIT E CHAMADO PELO Process DO .NET, e nao pelo operador &. No Windows
#     PowerShell 5.1, stderr de executavel nativo vira ErrorRecord: com
#     $ErrorActionPreference = 'Stop' isso ABORTA o script -- foi o que
#     aconteceu na primeira vez que rodei isto --, e com 'SilentlyContinue' o
#     texto do erro e DESCARTADO, mesmo com 2>&1. E aqui esse texto e a decisao:
#     "not a git repository" e a unica resposta que autoriza seguir.
#
#     Redirecionar pelo ProcessStartInfo nao passa por nada disso.
$oGit = Get-Command git -ErrorAction SilentlyContinue
if (-not $oGit) {
    throw ("RECUSADO: nao achei o git no PATH, e sem ele nao da para saber se " +
           "$absoluto esta dentro de um repositorio. Instale o git ou escolha " +
           "um destino que voce tenha certeza de que esta fora.")
}

$como = New-Object System.Diagnostics.ProcessStartInfo
$como.FileName = $oGit.Source
$como.Arguments = "-C `"$pastaAbsoluta`" rev-parse --git-dir"
$como.RedirectStandardOutput = $true
$como.RedirectStandardError = $true
$como.UseShellExecute = $false
$como.CreateNoWindow = $true

# As variaveis de ambiente do git sao LIMPAS para esta pergunta: GIT_DIR,
# GIT_WORK_TREE e GIT_CEILING_DIRECTORIES mudam a resposta, e quem responde tem
# de ser o disco, e nao o ambiente de quem chamou.
foreach ($nome in 'GIT_DIR', 'GIT_WORK_TREE', 'GIT_COMMON_DIR', 'GIT_CEILING_DIRECTORIES') {
    if ($como.EnvironmentVariables.ContainsKey($nome)) {
        $como.EnvironmentVariables.Remove($nome) | Out-Null
    }
}

$quem = [System.Diagnostics.Process]::Start($como)
$oQueOGitDisse = ($quem.StandardOutput.ReadToEnd() + ' ' +
                  $quem.StandardError.ReadToEnd()).Trim()
$quem.WaitForExit()
$codigoDoGit = $quem.ExitCode
$quem.Dispose()

if ($codigoDoGit -eq 0) {
    throw ("RECUSADO: $absoluto esta dentro de um repositorio git " +
           "($oQueOGitDisse). A chave privada nao entra no git.")
}
if ($oQueOGitDisse -notmatch 'not a git repository') {
    # NAO SEI, ENTAO NAO DEIXO. Qualquer outra falha -- permissao,
    # safe.directory, git quebrado -- e uma pergunta sem resposta, e nao um
    # "nao".
    throw ("RECUSADO: nao consegui perguntar ao git se $absoluto esta dentro de " +
           "um repositorio (saida $codigoDoGit): $oQueOGitDisse")
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

# ------------------------------------------------------- E A PASTA TEM DE SER SUA
#
# A ACL do ARQUIVO nao protege o NOME dentro da pasta: quem puder criar e apagar
# entradas ali pode trocar o arquivo entre a hora em que ele recebe a ACL e a
# hora em que a chave e escrita nele, e a chave acaba num objeto escolhido por
# outra pessoa. Conferir a ACL do arquivo depois nao resolve, porque a conferencia
# tem a mesma corrida.
#
# O que fecha isso e a pasta nao ser gravavel por terceiros.
#
# A CONFERENCIA E CALIBRADA, e a calibragem tem motivo. A primeira versao
# recusava qualquer identidade fora de uma lista curta, e recusou o proprio
# perfil do dono: um perfil do Windows costuma carregar ACEs de SIDs ORFAOS --
# restos de instalacao anterior, que nao resolvem para conta nenhuma e que
# ninguem pode usar para entrar. Recusar por causa deles tornaria o script
# inutil, e um script inutil e um script que se contorna.
#
# Entao: GRUPO AMPLO recusa, porque "Todos pode escrever aqui" e exatamente a
# ameaca. Identidade individual avisa, porque pode ser um SID orfao (inofensivo)
# ou outra pessoa de verdade (nao inofensiva) -- e daqui nao da para distinguir
# uma coisa da outra com confianca suficiente para impedir alguem de gerar a
# propria chave.
$gruposAmplos = @{
    'S-1-1-0'       = 'Todos'
    'S-1-5-11'      = 'Usuarios Autenticados'
    'S-1-5-32-545'  = 'Usuarios'
    'S-1-5-32-546'  = 'Convidados'
    'S-1-5-4'       = 'INTERATIVO'
    'S-1-5-7'       = 'LOGON ANONIMO'
}
$deConfianca = @(
    [System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value,
    'S-1-5-18',      # SYSTEM
    'S-1-5-32-544',  # BUILTIN\Administrators
    'S-1-3-0'        # CREATOR OWNER
)
$escrita = 'WriteData|CreateFiles|Delete|DeleteSubdirectoriesAndFiles|' +
           'ChangePermissions|TakeOwnership|FullControl|Modify|Write'

$avisar = @()
foreach ($regra in (Get-Acl $pastaAbsoluta).Access) {
    if ($regra.AccessControlType -ne 'Allow') { continue }
    if ($regra.FileSystemRights.ToString() -notmatch $escrita) { continue }

    $quem = try {
        $regra.IdentityReference.Translate(
            [System.Security.Principal.SecurityIdentifier]).Value
    } catch { "$($regra.IdentityReference)" }

    if ($deConfianca -contains $quem) { continue }

    if ($gruposAmplos.ContainsKey($quem)) {
        throw ("RECUSADO: '$($gruposAmplos[$quem])' pode escrever em " +
               "$pastaAbsoluta. Quem escreve na pasta pode trocar o arquivo da " +
               "chave depois de ele receber a ACL, e a ACL do arquivo nao " +
               "protege o NOME dentro da pasta. Escolha uma pasta que so voce " +
               "possa alterar -- o padrao, %USERPROFILE%\.iris, e uma.")
    }
    $avisar += "$($regra.IdentityReference) ($($regra.FileSystemRights))"
}

if ($avisar.Count -gt 0) {
    Write-Host ''
    Write-Host 'ATENCAO: estas identidades tambem podem escrever na pasta da chave:' -ForegroundColor Yellow
    $avisar | ForEach-Object { Write-Host "  $_" }
    Write-Host ''
    Write-Host 'Se alguma delas for outra PESSOA, ela pode trocar o arquivo da chave'
    Write-Host 'depois de ele receber a ACL. SIDs que aparecem como numero cru'
    Write-Host 'costumam ser restos de instalacao anterior, e nao pertencem a conta'
    Write-Host 'nenhuma -- esses sao inofensivos.'
    Write-Host ''
}

# ---------------------------------------------------------- a ACL vem PRIMEIRO
#
# O arquivo nasce VAZIO e ja restrito; so depois o utilitario escreve nele.
# Gravar e restringir em seguida deixava uma janela com a chave no disco sob a
# ACL herdada da pasta -- e, se o Set-Acl falhasse, ela ficava assim.
#
# E ELE NASCE MESMO: um arquivo vazio preexistente e APAGADO primeiro.
# SetAccessRuleProtection remove a heranca, mas nao remove ACEs EXPLICITAS que
# ja estivessem la -- entao um arquivo vazio preparado de antemao com uma regra
# permissiva receberia a chave e a manteria legivel para terceiros.
if (Test-Path $absoluto) { Remove-Item $absoluto -Force }
New-Item -ItemType File $absoluto | Out-Null

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
    $oQueSobrou = 'Nada foi gravado.'
    try { Remove-Item $absoluto -Force } catch { $oQueSobrou = "ATENCAO: sobrou um arquivo vazio em $absoluto -- apague-o." }
    throw "Nao consegui restringir a ACL de $absoluto. $oQueSobrou ($_)"
}

# ------------------------------------------------------------------- a chave

$ferramenta = & (Join-Path $PSScriptRoot 'montar-assinador.ps1')
$publica = & $ferramenta gerar --destino $absoluto
if ($LASTEXITCODE -ne 0) {
    # A MENSAGEM DEPENDE DE A LIMPEZA TER FUNCIONADO. Dizer "nada ficou no
    # disco" sem conferir e a mesma classe de erro que esta revisao inteira
    # persegue: afirmar mais do que se sabe.
    $oQueSobrou = 'Nada ficou no disco.'
    try { Remove-Item $absoluto -Force } catch { }
    if (Test-Path $absoluto) {
        $oQueSobrou = "ATENCAO: sobrou um arquivo em $absoluto -- apague-o antes de tentar de novo."
    }
    throw "A geracao da chave falhou com $LASTEXITCODE. $oQueSobrou"
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
