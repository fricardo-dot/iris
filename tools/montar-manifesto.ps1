<#
.SYNOPSIS
    Escreve o iris.json de uma versão. Um lugar só, para poder ser testado.

.DESCRIPTION
    Isto era três linhas dentro de publicar-versao.ps1, e por isso o teste que
    afirmava cobrir "a forma que o ConvertTo-Json escreve" era uma cópia
    literal, feita à mão, de uma saída de 02/09/2026. Mudar o script não
    derrubava o teste — e o comentário do teste dizia que derrubaria.

    Agora a geração mora aqui, e a suíte CHAMA este script e lê o que ele
    escreve com o mesmo ManifestoDeVersao.Ler que roda na máquina do usuário.
    Mudar a forma aqui derruba o teste de verdade.

    A ordem dos campos e o espaçamento não importam para o cliente — JSON é
    JSON. O que importa, e o que o teste passa a proteger, é: `bytes` sair como
    NÚMERO e não string, o arquivo não ter BOM, e os nomes dos campos serem os
    que o leitor procura.

.PARAMETER Destino
    Onde gravar o iris.json.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string] $Versao,
    [Parameter(Mandatory)] [string] $Notas,
    [Parameter(Mandatory)] [string] $Endereco,
    [Parameter(Mandatory)] [string] $Sha256,
    [Parameter(Mandatory)] [string] $Bytes,
    [Parameter(Mandatory)] [string] $Destino,
    [string] $Publicada
)

$ErrorActionPreference = 'Stop'

# [string] E NAO [long] NO PARAMETRO, com a conferencia a mao: no 5.1, [long]
# converte "1.5" em 2 e "1e3" em 1000, calado. O tamanho de um arquivo nunca e
# nenhuma dessas coisas, e um manifesto assinado com o numero errado so seria
# descoberto pelo Iris recusando o download por tamanho.
if ($Bytes -notmatch '^[0-9]+$') {
    throw "Bytes tem de ser um inteiro decimal. Veio: $Bytes"
}
$quantosBytes = [long] $Bytes

if (-not $Publicada) {
    $Publicada = (Get-Date).ToUniversalTime().ToString('o')
}

$manifesto = [ordered] @{
    versao    = $Versao
    publicada = $Publicada
    notas     = $Notas
    endereco  = $Endereco
    sha256    = $Sha256
    bytes     = $quantosBytes
}

# SEM BOM. Tres bytes invisiveis na frente fariam parte do que foi assinado e do
# que sera conferido, entao a assinatura ainda bateria -- mas JsonDocument.Parse
# tropeca neles, e o erro sairia como "o manifesto nao e JSON legivel" DEPOIS de
# a assinatura conferir, que e o pior lugar para procurar.
$json = $manifesto | ConvertTo-Json -Depth 3
$bytesDoManifesto = [System.Text.UTF8Encoding]::new($false).GetBytes($json)

if ($bytesDoManifesto.Length -gt 65536) {
    # O teto do cliente: ManifestoDeVersao.ManifestoMaximo.
    throw "O manifesto tem $($bytesDoManifesto.Length) bytes; o Iris recusa acima de 64 KiB."
}

# COM NOME TEMPORARIO, como o .sig. WriteAllBytes direto no destino cria ou
# trunca antes de terminar: falta de espaco ou erro de I/O deixariam um
# iris.json vazio ou parcial com cara de manifesto pronto -- e este script
# existe justamente para ser chamado de mais de um lugar.
$temporario = "$Destino.$([guid]::NewGuid().ToString('N')).parcial"
try {
    [System.IO.File]::WriteAllBytes($temporario, $bytesDoManifesto)
    # Move-Item -Force, e nao [File]::Move com tres argumentos: a sobrecarga
    # com overwrite nao existe no .NET Framework, que e onde o PowerShell 5.1
    # roda. Ela existe no .NET moderno, e foi de la que eu a copiei.
    #
    # E O QUE ISTO COMPRA E MENOS DO QUE PARECE. O ganho real e nao gravar
    # direto no nome final: uma falha no meio do WriteAllBytes nao deixa um
    # iris.json parcial com cara de pronto. Se a PROMOCAO em si e atomica
    # depende do provider e nao esta demonstrado aqui -- e para este arquivo,
    # que e lido logo depois pelo mesmo script, isso nao muda nada.
    Move-Item -LiteralPath $temporario -Destination $Destino -Force
}
catch {
    if (Test-Path $temporario) { Remove-Item $temporario -Force -ErrorAction SilentlyContinue }
    throw
}
