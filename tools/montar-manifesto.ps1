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
    [Parameter(Mandatory)] [long]   $Bytes,
    [Parameter(Mandatory)] [string] $Destino,
    [string] $Publicada
)

$ErrorActionPreference = 'Stop'

if (-not $Publicada) {
    $Publicada = (Get-Date).ToUniversalTime().ToString('o')
}

$manifesto = [ordered] @{
    versao    = $Versao
    publicada = $Publicada
    notas     = $Notas
    endereco  = $Endereco
    sha256    = $Sha256
    bytes     = $Bytes
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

[System.IO.File]::WriteAllBytes($Destino, $bytesDoManifesto)
