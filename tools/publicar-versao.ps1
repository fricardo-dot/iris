<#
.SYNOPSIS
    Compila, empacota, assina e prepara uma versão do Iris para publicação.

.DESCRIPTION
    O caminho inteiro de uma release, numa ordem que não deixa gerar um pacote
    sem assinatura nem uma assinatura sem pacote:

      1. grava a versão em Directory.Build.props (fonte única do número);
      2. publica autocontido, num .exe só;
      3. calcula o SHA-256 DO ARQUIVO QUE ACABOU DE SAIR;
      4. escreve iris.json com esse hash e o endereço de download;
      5. assina o iris.json inteiro, byte a byte, com a sua chave privada.

    AUTOCONTIDO, e é o que faz a segunda máquina funcionar: o .NET 10 vai dentro
    do executável. O arquivo passa dos 100 MB e essa é a troca — instalar runtime
    numa máquina é um passo a mais que dá errado calado.

    O QUE ESTE SCRIPT NÃO FAZ: publicar. Ele deixa os três arquivos prontos e
    imprime o comando. Subir a release é um ato seu, e um ato público.

.PARAMETER Versao
    A versão nova, como 0.2.0. Tem de ser MAIOR que a que está no
    Directory.Build.props — o Iris recusa oferta de versão que não sobe, e
    publicar uma release que ele ignora não daria erro em lugar nenhum.

.PARAMETER Notas
    O que mudou, em português, para aparecer na tela de quem for atualizar.

.PARAMETER Repositorio
    O repositório do GitHub, como "ricardo/iris". É dele que sai o endereço de
    download que vai assinado dentro do manifesto.

.PARAMETER Chave
    A chave privada gerada por gerar-chave-de-assinatura.ps1.

.PARAMETER Publicar
    Sobe a release pelo gh. SEM ISTO, o script só prepara e mostra o comando.

.EXAMPLE
    .\tools\publicar-versao.ps1 -Versao 0.2.0 -Repositorio ricardo/iris `
        -Notas "Classificação em lote e verificação de versões."
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string] $Versao,
    [Parameter(Mandatory)] [string] $Notas,
    [Parameter(Mandatory)] [string] $Repositorio,
    [string] $Chave = (Join-Path $env:USERPROFILE '.iris\chave-de-assinatura.pem'),
    [switch] $Publicar
)

$ErrorActionPreference = 'Stop'
$raiz = Split-Path -Parent $PSScriptRoot

# --------------------------------------------------------------- conferencias

if ($Versao -notmatch '^\d+\.\d+\.\d+$') {
    throw "A versao tem de ser tres numeros, como 0.2.0. Veio: $Versao"
}
if ($Repositorio -notmatch '^[\w.-]+/[\w.-]+$') {
    throw "O repositorio e 'dono/nome'. Veio: $Repositorio"
}
if (-not (Test-Path $Chave)) {
    throw "Nao achei a chave em $Chave. Rode tools/gerar-chave-de-assinatura.ps1 primeiro."
}

$props = Join-Path $raiz 'Directory.Build.props'
$textoDosProps = Get-Content $props -Raw
if ($textoDosProps -notmatch '<Version>([\d.]+)</Version>') {
    throw "Nao achei <Version> em $props"
}
$atual = [Version] $Matches[1]
$nova = [Version] $Versao

# A VERSAO TEM DE SUBIR. O Iris compara com <= e trata "igual" como ja estar em
# dia: republicar o mesmo numero produziria uma release que ninguem baixa, sem
# erro em lugar nenhum, e a descoberta seria por alguem reclamando que a
# atualizacao nao chega.
if ($nova -le $atual) {
    throw "A versao tem de subir. Instalada aqui: $atual. Pedida: $nova."
}

# ------------------------------------------------------------------ compilacao

Write-Host "Gravando a versao $nova em Directory.Build.props..." -ForegroundColor Cyan
$textoDosProps -replace '<Version>[\d.]+</Version>', "<Version>$Versao</Version>" |
    Set-Content $props -Encoding utf8 -NoNewline

$saida = Join-Path $raiz "artefatos\$Versao"
if (Test-Path $saida) { Remove-Item -Recurse -Force $saida }
New-Item -ItemType Directory -Force $saida | Out-Null

Write-Host 'Publicando autocontido (isto demora)...' -ForegroundColor Cyan
$projeto = Join-Path $raiz 'src\Iris.App\Iris.App.vbproj'
& dotnet publish $projeto -c Release -r win-x64 --self-contained true `
    -p:PublishSingleFile=true `
    -p:IncludeNativeLibrariesForSelfExtract=true `
    -p:EnableCompressionInSingleFile=true `
    -p:DebugType=none `
    -o $saida --nologo
if ($LASTEXITCODE -ne 0) { throw "dotnet publish falhou com $LASTEXITCODE" }

# O NOME QUE O MANIFESTO PROMETE. ProcuraDeVersao.Baixar grava exatamente
# "Iris-<versao>.exe", e o endereco aqui tem de bater com o nome do ativo da
# release -- se divergirem, o download da 404 e a mensagem fala de rede.
$exe = Join-Path $saida 'Iris.exe'
if (-not (Test-Path $exe)) { throw "O publish nao produziu $exe" }
$pacote = Join-Path $saida "Iris-$Versao.exe"
Move-Item $exe $pacote -Force

# Um autocontido de WPF nao tem como ser pequeno; alguns megabytes significam
# que o publish saiu framework-dependent e a segunda maquina nao vai abrir.
$tamanho = (Get-Item $pacote).Length
if ($tamanho -lt 20MB) {
    throw "O pacote tem so $([math]::Round($tamanho/1MB,1)) MB -- isto nao parece autocontido."
}

# ------------------------------------------------------------- hash e manifesto

$sha = (Get-FileHash $pacote -Algorithm SHA256).Hash.ToLowerInvariant()
$endereco = "https://github.com/$Repositorio/releases/download/v$Versao/Iris-$Versao.exe"

$manifesto = [ordered] @{
    versao    = $Versao
    publicada = (Get-Date).ToUniversalTime().ToString('o')
    notas     = $Notas
    endereco  = $endereco
    sha256    = $sha
    bytes     = $tamanho
}

# SEM BOM. O manifesto e assinado byte a byte; tres bytes invisiveis na frente
# fazem parte do que foi assinado e do que sera conferido, entao nao quebram
# nada -- mas JsonDocument.Parse tropeca neles, e o erro sairia como "o
# manifesto nao e JSON legivel" depois de a assinatura conferir, que e o pior
# lugar para procurar.
$json = $manifesto | ConvertTo-Json -Depth 3
$bytesDoManifesto = [System.Text.UTF8Encoding]::new($false).GetBytes($json)

$arquivoDoManifesto = Join-Path $saida 'iris.json'
[System.IO.File]::WriteAllBytes($arquivoDoManifesto, $bytesDoManifesto)

# ------------------------------------------------------------------- assinatura

$assinador = [System.Security.Cryptography.ECDsa]::Create()
try {
    $assinador.ImportFromPem((Get-Content $Chave -Raw))
    $assinatura = $assinador.SignData(
        $bytesDoManifesto, [System.Security.Cryptography.HashAlgorithmName]::SHA256)

    # E CONFERE AQUI MESMO. Assinar e nao verificar deixa descobrir um par
    # trocado la na frente, na maquina de destino, com a mensagem "este arquivo
    # nao foi publicado por voce" -- dita sobre um arquivo que foi.
    if (-not $assinador.VerifyData(
            $bytesDoManifesto, $assinatura,
            [System.Security.Cryptography.HashAlgorithmName]::SHA256)) {
        throw 'A assinatura nao confere contra a propria chave que a fez.'
    }
}
finally {
    $assinador.Dispose()
}

$arquivoDaAssinatura = Join-Path $saida 'iris.json.sig'
[System.IO.File]::WriteAllBytes($arquivoDaAssinatura, $assinatura)

# ----------------------------------------------------------------------- fim

Write-Host ''
Write-Host "Versao $Versao pronta em $saida" -ForegroundColor Green
Write-Host ("  Iris-$Versao.exe   {0:N1} MB" -f ($tamanho / 1MB))
Write-Host "  iris.json          sha256 $sha"
Write-Host "  iris.json.sig      $($assinatura.Length) bytes"
Write-Host ''

$comando = "gh release create v$Versao " +
           "`"$pacote`" `"$arquivoDoManifesto`" `"$arquivoDaAssinatura`" " +
           "--repo $Repositorio --title `"Iris $Versao`" --notes `"$Notas`""

if ($Publicar) {
    Write-Host 'Subindo a release...' -ForegroundColor Cyan
    & gh release create "v$Versao" $pacote $arquivoDoManifesto $arquivoDaAssinatura `
        --repo $Repositorio --title "Iris $Versao" --notes $Notas
    if ($LASTEXITCODE -ne 0) { throw "gh release create falhou com $LASTEXITCODE" }
    Write-Host 'Publicado.' -ForegroundColor Green
}
else {
    Write-Host 'Para publicar (isto torna os arquivos publicos):' -ForegroundColor Yellow
    Write-Host ''
    Write-Host "  $comando"
    Write-Host ''
    Write-Host 'Ou rode este script de novo com -Publicar.'
}

Write-Host ''
Write-Host 'E nao esqueca o commit: Directory.Build.props mudou.'
