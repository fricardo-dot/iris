<#
.SYNOPSIS
    Compila, empacota, assina e prepara uma versão do Iris para publicação.

.DESCRIPTION
    O caminho inteiro de uma release, numa ordem que não deixa gerar um pacote
    sem assinatura nem uma assinatura sem pacote:

      1. publica autocontido, num .exe só, com a versão nova passada ao build;
      2. calcula o SHA-256 DO ARQUIVO QUE ACABOU DE SAIR;
      3. escreve iris.json com esse hash e o endereço de download;
      4. assina o iris.json inteiro, byte a byte, com a sua chave privada;
      5. só então — e, com -Publicar, só depois de a release existir —
         grava a versão em Directory.Build.props.

    A ORDEM DO PASSO 5 IMPORTA. Ele era o primeiro, e uma falha em qualquer
    etapa seguinte deixava a versão gravada — a reexecução com o mesmo número
    era então recusada pela própria conferência de "a versão tem de subir", e o
    conserto exigia editar o arquivo à mão.

    Passou para depois da assinatura, e depois ainda para DEPOIS DO gh: uma
    falha ao subir a release deixava o mesmo rastro, pelo mesmo motivo.

    AUTOCONTIDO, e é o que faz a segunda máquina funcionar: o .NET 10 vai dentro
    do executável. O arquivo passa dos 60 MB e essa é a troca — instalar runtime
    numa máquina é um passo a mais que dá errado calado.

    A CRIPTOGRAFIA NÃO ACONTECE AQUI, e sim em tools/Iris.Assinatura: o Windows
    PowerShell 5.1 roda sobre .NET Framework, onde ImportFromPem não existe.

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

.PARAMETER MostrarChavePublica
    Só imprime a chave pública correspondente a -Chave e sai. Serve para
    recuperá-la sem gerar par novo, que invalidaria tudo o que já foi publicado.

.EXAMPLE
    .\tools\publicar-versao.ps1 -Versao 0.2.0 -Repositorio ricardo/iris `
        -Notas "Classificação em lote e verificação de versões."
#>
[CmdletBinding(DefaultParameterSetName = 'Publicar')]
param(
    [Parameter(Mandatory, ParameterSetName = 'Publicar')] [string] $Versao,
    [Parameter(Mandatory, ParameterSetName = 'Publicar')] [string] $Notas,
    [Parameter(Mandatory, ParameterSetName = 'Publicar')] [string] $Repositorio,
    [Parameter(Mandatory, ParameterSetName = 'Chave')] [switch] $MostrarChavePublica,
    [string] $Chave = (Join-Path $env:USERPROFILE '.iris\chave-de-assinatura.pem'),
    [switch] $Publicar,

    # ONDE OS PACOTES FICAM. Vazio = descobre sozinho, ver abaixo.
    [string] $Saida
)

$ErrorActionPreference = 'Stop'
$raiz = Split-Path -Parent $PSScriptRoot

if (-not (Test-Path $Chave)) {
    throw "Nao achei a chave em $Chave. Rode tools/gerar-chave-de-assinatura.ps1 primeiro."
}

$ferramenta = & (Join-Path $PSScriptRoot 'montar-assinador.ps1')

if ($MostrarChavePublica) {
    & $ferramenta publica --chave $Chave
    if ($LASTEXITCODE -ne 0) { throw "Nao consegui ler a chave publica ($LASTEXITCODE)." }
    exit 0
}

# --------------------------------------------------------------- conferencias

if ($Versao -notmatch '^\d+\.\d+\.\d+$') {
    throw "A versao tem de ser tres numeros, como 0.2.0. Veio: $Versao"
}
if ($Repositorio -notmatch '^[\w.-]+/[\w.-]+$') {
    throw "O repositorio e 'dono/nome'. Veio: $Repositorio"
}

$props = Join-Path $raiz 'Directory.Build.props'

# [IO.File]::ReadAllText, e nao Get-Content -Raw: no 5.1, Get-Content decodifica
# UTF-8 SEM BOM usando a pagina ANSI, e os acentos dos comentarios voltariam
# corrompidos -- para serem regravados corrompidos.
$textoDosProps = [System.IO.File]::ReadAllText($props)

$achadas = [regex]::Matches($textoDosProps, '<Version>([\d.]+)</Version>')
if ($achadas.Count -ne 1) {
    # EXATAMENTE UMA. O -replace troca TODAS as ocorrencias, e a comparacao usa
    # so a primeira: com duas, o script alteraria uma que ninguem pediu.
    throw "Esperava exatamente um <Version> em $props; achei $($achadas.Count)."
}
$atual = [Version] $achadas[0].Groups[1].Value
$nova = [Version] $Versao

# A VERSAO TEM DE SUBIR. O Iris compara com <= e trata "igual" como ja estar em
# dia: republicar o mesmo numero produziria uma release que ninguem baixa, sem
# erro em lugar nenhum, e a descoberta seria por alguem reclamando que a
# atualizacao nao chega.
if ($nova -le $atual) {
    throw "A versao tem de subir. Instalada aqui: $atual. Pedida: $nova."
}

# O TETO DO CLIENTE, CONFERIDO AQUI. ManifestoDeVersao recusa manifesto acima de
# 64 KiB, e a nota e o unico campo de tamanho livre. Descobrir isso depois de
# assinar e publicar seria descobrir pela reclamacao de quem tentou atualizar.
$bytesDasNotas = [System.Text.Encoding]::UTF8.GetByteCount($Notas)
if ($bytesDasNotas -gt 60000) {
    throw "As notas tem $bytesDasNotas bytes; o manifesto inteiro nao pode passar de 64 KiB."
}

# ------------------------------------------------------------------ compilacao

# ONDE GRAVAR, em tres degraus -- e nenhum deles crava uma letra de unidade
# no codigo, porque a segunda maquina nao tem as mesmas.
#
#   1. -Saida, se voce passar;
#   2. a variavel de ambiente IRIS_ARTEFATOS, se existir;
#   3. artefatos\ dentro do repositorio, que e o padrao e esta no .gitignore.
#
# O degrau 2 existe para a maquina que tem um disco melhor: define-se uma vez,
# e todas as publicacoes seguintes vao para la sem ninguem lembrar de nada. Um
# .exe autocontido tem 63 MB, e tres releases enchem um C: apertado.
$ondeGravar = $Saida
if (-not $ondeGravar) { $ondeGravar = $env:IRIS_ARTEFATOS }
if (-not $ondeGravar) { $ondeGravar = Join-Path $raiz 'artefatos' }
$saida = Join-Path $ondeGravar $Versao
if (Test-Path $saida) { Remove-Item -Recurse -Force $saida }
New-Item -ItemType Directory -Force $saida | Out-Null

Write-Host 'Publicando autocontido (isto demora)...' -ForegroundColor Cyan
$projeto = Join-Path $raiz 'src\Iris.App\Iris.App.vbproj'

# -p:Version NA LINHA DE COMANDO. Assim o build ja sai com a versao nova sem
# que Directory.Build.props tenha sido tocado -- e uma falha aqui nao deixa
# rastro nenhum no repositorio.
& dotnet publish $projeto -c Release -r win-x64 --self-contained true `
    -p:Version=$Versao `
    -p:PublishSingleFile=true `
    -p:IncludeNativeLibrariesForSelfExtract=true `
    -p:EnableCompressionInSingleFile=true `
    -p:DebugType=none `
    -o $saida --nologo
if ($LASTEXITCODE -ne 0) { throw "dotnet publish falhou com $LASTEXITCODE" }

# O NOME QUE O MANIFESTO PROMETE. ProcuraDeVersao grava exatamente
# "Iris-<versao>.exe", e o endereco aqui tem de bater com o nome do ativo da
# release -- se divergirem, o download da 404 e a mensagem fala de rede.
$exe = Join-Path $saida 'Iris.exe'
if (-not (Test-Path $exe)) { throw "O publish nao produziu $exe" }
$pacote = Join-Path $saida "Iris-$Versao.exe"
Move-Item $exe $pacote -Force

$tamanho = (Get-Item $pacote).Length

# Um autocontido de WPF nao tem como ser pequeno; alguns megabytes significam
# que o publish saiu framework-dependent e a segunda maquina nao vai abrir.
if ($tamanho -lt 20MB) {
    throw "O pacote tem so $([math]::Round($tamanho/1MB,1)) MB -- isto nao parece autocontido."
}
# E O TETO DO CLIENTE: ManifestoDeVersao.TamanhoMaximo sao 300 MB, e um pacote
# maior produziria um manifesto assinado que o Iris recusa antes de baixar.
if ($tamanho -gt 300MB) {
    throw "O pacote tem $([math]::Round($tamanho/1MB,1)) MB e o Iris recusa acima de 300 MB."
}

# ------------------------------------------------------------- hash e manifesto

$sha = (Get-FileHash $pacote -Algorithm SHA256).Hash.ToLowerInvariant()
$endereco = "https://github.com/$Repositorio/releases/download/v$Versao/Iris-$Versao.exe"

# A GERACAO DO JSON MORA EM OUTRO SCRIPT, e nao por organizacao: e para a
# suite poder CHAMA-LA. Enquanto ela estava aqui dentro, o teste que afirmava
# cobrir "a forma que o ConvertTo-Json escreve" era uma copia literal feita a
# mao -- mudar este arquivo nao derrubava teste nenhum.
$arquivoDoManifesto = Join-Path $saida 'iris.json'
& (Join-Path $PSScriptRoot 'montar-manifesto.ps1') `
    -Versao $Versao -Notas $Notas -Endereco $endereco -Sha256 $sha `
    -Bytes $tamanho -Destino $arquivoDoManifesto
if (-not (Test-Path $arquivoDoManifesto)) { throw "Nao consegui montar $arquivoDoManifesto" }

# ------------------------------------------------------------------- assinatura

& $ferramenta assinar --chave $Chave --arquivo $arquivoDoManifesto | Out-Null
if ($LASTEXITCODE -ne 0) { throw "A assinatura falhou com $LASTEXITCODE" }

$arquivoDaAssinatura = "$arquivoDoManifesto.sig"
if (-not (Test-Path $arquivoDaAssinatura)) { throw "O assinador nao produziu $arquivoDaAssinatura" }

# 64 BYTES, E NAO "EXISTE". Uma assinatura ECDSA P-256 em IeeeP1363 tem
# exatamente isso; qualquer outro tamanho e um arquivo truncado com nome de
# arquivo inteiro, e o Iris o recusaria na maquina do dono com "este arquivo
# nao foi publicado por voce" -- dito sobre um arquivo que foi.
$bytesDaAssinatura = (Get-Item $arquivoDaAssinatura).Length
if ($bytesDaAssinatura -ne 64) {
    throw "A assinatura tem $bytesDaAssinatura bytes; uma P-256 tem 64."
}

# ----------------------------------------------------------------------- fim

# AS NOTAS VAO POR ARQUIVO, e nao por argumento. Uma aspa dentro delas encerra
# o argumento antes da hora na reconstrucao de linha de comando que o
# PowerShell 5.1 faz para programas nativos, e uma quebra de linha vira outro
# comando. --notes-file nao tem nenhum dos dois problemas.
$arquivoDasNotas = Join-Path $saida 'notas.txt'
[System.IO.File]::WriteAllText($arquivoDasNotas, $Notas,
                               [System.Text.UTF8Encoding]::new($false))

Write-Host ''
Write-Host "Versao $Versao pronta em $saida" -ForegroundColor Green
if ($env:IRIS_ARTEFATOS -and -not $Saida) {
    Write-Host "  (por IRIS_ARTEFATOS; use -Saida para mandar para outro lugar)" -ForegroundColor DarkGray
}
Write-Host ("  Iris-$Versao.exe   {0:N1} MB" -f ($tamanho / 1MB))
Write-Host "  iris.json          sha256 $sha"
Write-Host "  iris.json.sig      $((Get-Item $arquivoDaAssinatura).Length) bytes"
Write-Host ''

if ($Publicar) {
    Write-Host 'Subindo a release...' -ForegroundColor Cyan
    # --latest EXPLICITO: o endereco que o Iris consulta e
    # releases/latest/download/iris.json, e "latest" e um metadado do GitHub, e
    # nao a maior versao. Deixar implicito e deixar a conta certa para o acaso.
    & gh release create "v$Versao" $pacote $arquivoDoManifesto $arquivoDaAssinatura `
        --repo $Repositorio --title "Iris $Versao" --notes-file $arquivoDasNotas --latest
    if ($LASTEXITCODE -ne 0) { throw "gh release create falhou com $LASTEXITCODE" }
    Write-Host 'Publicado.' -ForegroundColor Green
}
else {
    Write-Host 'Para publicar (isto torna os arquivos publicos):' -ForegroundColor Yellow
    Write-Host ''
    Write-Host "  gh release create v$Versao ``"
    Write-Host "    `"$pacote`" ``"
    Write-Host "    `"$arquivoDoManifesto`" ``"
    Write-Host "    `"$arquivoDaAssinatura`" ``"
    Write-Host "    --repo $Repositorio --title `"Iris $Versao`" ``"
    Write-Host "    --notes-file `"$arquivoDasNotas`" --latest"
    Write-Host ''
    Write-Host 'Para publicar rodando este script de novo com -Publicar, desfaca'
    Write-Host 'antes a versao gravada -- senao a conferencia de "a versao tem de'
    Write-Host 'subir" vai recusar o mesmo numero:'
    Write-Host ''
    Write-Host '  git checkout -- Directory.Build.props'
    Write-Host "  .\tools\publicar-versao.ps1 -Versao $Versao -Repositorio $Repositorio -Notas ... -Publicar"
}

# --------------------------------------------------- e SO AGORA a versao entra
#
# DEPOIS DO gh, e nao antes. Gravar antes de publicar deixava o numero no
# repositorio quando o `gh release create` falhava -- e a reexecucao com o mesmo
# numero era entao recusada pela propria conferencia de "a versao tem de subir",
# obrigando a editar o arquivo a mao. O ponto de nao-retorno e a release existir.
[System.IO.File]::WriteAllText(
    $props,
    ($textoDosProps -replace '<Version>[\d.]+</Version>', "<Version>$Versao</Version>"),
    [System.Text.UTF8Encoding]::new($false))

Write-Host ''
Write-Host 'E nao esqueca o commit: Directory.Build.props mudou.'
if (-not $Publicar) {
    Write-Host 'A versao ja foi gravada. Se voce NAO for publicar esta release,'
    Write-Host 'desfaca com: git checkout -- Directory.Build.props'
}
