<#
.SYNOPSIS
    Compila tools/Iris.Assinatura e devolve o caminho do executável.

.DESCRIPTION
    Um lugar só, porque os dois scripts precisam da mesma coisa e uma cópia
    divergiria.

    POR QUE COMPILAR E CHAMAR O EXE, e não `dotnet run`: o `dotnet run` escreve
    a saída da compilação na mesma saída padrão do programa, e a primeira coisa
    que este utilitário imprime é a chave pública. Uma linha de build no meio
    viraria parte da chave que alguém vai colar no código.

    Silencioso quando dá certo; a saída da compilação só aparece se ela falhar.

    E o stderr do `dotnet build` NÃO é redirecionado. No Windows PowerShell 5.1,
    `& nativo 2>&1` embrulha cada linha de stderr num ErrorRecord, e com
    `$ErrorActionPreference = 'Stop'` isso ABORTA a atribuição mesmo quando o
    processo sai com código zero — matando o script por um aviso de build.
#>
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

$projeto = Join-Path $PSScriptRoot 'Iris.Assinatura\Iris.Assinatura.vbproj'
if (-not (Test-Path $projeto)) { throw "Nao achei $projeto" }

# Sem 2>&1: ver a nota acima. O stdout vai para $log; o stderr passa direto
# para o console, que e onde ele serve para alguma coisa.
$log = & dotnet build $projeto -c Release --nologo -v quiet
if ($LASTEXITCODE -ne 0) {
    $log | ForEach-Object { Write-Host $_ }
    throw "Nao consegui compilar o assinador (dotnet build saiu com $LASTEXITCODE)."
}

$exe = Join-Path (Split-Path -Parent $projeto) 'bin\Release\net10.0\Iris.Assinatura.exe'
if (-not (Test-Path $exe)) { throw "A compilacao nao produziu $exe" }

# O CAMINHO E SO O CAMINHO. Quem chama faz `& $ferramenta ...`, entao nada
# alem disto pode sair daqui.
$exe
